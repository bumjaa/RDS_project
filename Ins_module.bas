Attribute VB_Name = "Ins_module"
Option Explicit

Private gRowsBySheet As Object   ' Scripting.Dictionary: sheetName -> Variant(2D)
Private gHdrBySheet  As Object   ' Scripting.Dictionary: sheetName -> Scripting.Dictionary(header->col)

Private Sub EnsureUsageStores()
    If gRowsBySheet Is Nothing Then Set gRowsBySheet = CreateObject("Scripting.Dictionary")
    If gHdrBySheet Is Nothing Then Set gHdrBySheet = CreateObject("Scripting.Dictionary")
End Sub

Private Function FetchUsageDumpCached(ByVal sheetName As String) As Boolean
    EnsureUsageStores
    If gRowsBySheet.Exists(sheetName) Then
        FetchUsageDumpCached = True
        Exit Function
    End If

    Dim url As String, raw As String, parsed As Object
    Dim baseUrl As String
    baseUrl = GetApiUrl("usage_url")
    url = baseUrl & "?sheet=" & URLEncodeUTF8(sheetName) & "&dump=true&ts=" & CLng(Timer * 1000)

    raw = HttpGet(url)

    If Len(raw) = 0 Or Left$(Trim$(raw), 1) = "<" Then
        Err.Raise vbObjectError, , "Non-JSON response: " & Left$(raw, 200)
    End If

    Set parsed = JsonConverter.ParseJson(raw)

    Dim hdrC As Collection, rowsC As Collection
    Set hdrC = parsed("header")
    Set rowsC = parsed("rows")

    Dim colN As Long, rowN As Long, r As Long, c As Long
    colN = hdrC.Count
    rowN = rowsC.Count

    Dim hdrIdx As Object
    Set hdrIdx = CreateObject("Scripting.Dictionary")
    For c = 1 To colN
        hdrIdx(CStr(hdrC(c))) = c
    Next

    Dim rows2D() As Variant
    ReDim rows2D(1 To rowN, 1 To colN)

    Dim oneRow As Collection
    For r = 1 To rowN
        Set oneRow = rowsC(r)
        For c = 1 To colN
            rows2D(r, c) = oneRow(c)
        Next
    Next

    gRowsBySheet.Add sheetName, rows2D
    gHdrBySheet.Add sheetName, hdrIdx

    FetchUsageDumpCached = True
End Function

Public Sub ClearUsageSheetCache(Optional ByVal sheetName As String = "")
    EnsureUsageStores
    If sheetName = "" Then
        gRowsBySheet.RemoveAll
        gHdrBySheet.RemoveAll
    Else
        If gRowsBySheet.Exists(sheetName) Then gRowsBySheet.Remove sheetName
        If gHdrBySheet.Exists(sheetName) Then gHdrBySheet.Remove sheetName
    End If
End Sub

Private Function ParseYmdDate(ByVal v As Variant) As Date
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then ParseYmdDate = 0: Exit Function

    If Len(s) >= 10 And Mid$(s, 5, 1) = "-" Then
        ParseYmdDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
        Exit Function
    End If

    If Len(s) >= 10 And Mid$(s, 5, 1) = "." Then
        ParseYmdDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
        Exit Function
    End If

    If IsDate(v) Then
        ParseYmdDate = DateValue(CDate(v))
    Else
        ParseYmdDate = 0
    End If
End Function


Private Function MinValidDate(ByVal rng As Range) As Date
    Dim cell As Range
    Dim d As Date, dMin As Date
    Dim found As Boolean
    Dim v As Variant

    For Each cell In rng.Cells
        v = cell.Value
        If IsDate(v) Then
            d = DateValue(CDate(v))
            If d > 0 Then
                If (Not found) Or d < dMin Then
                    dMin = d
                    found = True
                End If
            End If
        End If
    Next cell

    If found Then
        MinValidDate = dMin
    Else
        MinValidDate = Date
    End If
End Function

Private Function BuildUsageFromCache(ByVal sheetName As String, ByVal locQ As String, ByVal itemQ As String, ByVal cutoffDt As Date) As Variant
    cutoffDt = DateValue(cutoffDt)
    EnsureUsageStores
    If Not FetchUsageDumpCached(sheetName) Then
        BuildUsageFromCache = Empty
        Exit Function
    End If

    Dim rows2D As Variant, hdrIdx As Object
    rows2D = gRowsBySheet(sheetName)
    Set hdrIdx = gHdrBySheet(sheetName)

    Dim ixEff As Long, ixLoc As Long, ixItem As Long, ixCode As Long
    Dim ixInst As Long, ixModel As Long, ixMaker As Long, ixSerial As Long
    Dim ixPrev As Long, ixCal As Long, ixNext As Long, ixPeriod As Long

    ixEff = hdrIdx("EffectiveDate")
    ixLoc = hdrIdx("Location")
    ixItem = hdrIdx("Item")
    ixCode = hdrIdx("Code")
    ixInst = hdrIdx("Instrument_Name")
    ixModel = hdrIdx("Model_Name")
    ixMaker = hdrIdx("Manufacturer")
    ixSerial = hdrIdx("Serial_No")
    ixPrev = hdrIdx("PreviousCal")
    ixCal = hdrIdx("CalDate")
    ixNext = hdrIdx("NextCal")
    ixPeriod = hdrIdx("Cal_Period")

    Dim bestRow As Object, bestEff As Object
    Set bestRow = CreateObject("Scripting.Dictionary")
    Set bestEff = CreateObject("Scripting.Dictionary")

    Dim r As Long, effDate As Date
    Dim rowLoc As String, rowItem As String, code As String

    For r = LBound(rows2D, 1) To UBound(rows2D, 1)
        rowLoc = Trim$(CStr(rows2D(r, ixLoc)))
        rowItem = Trim$(CStr(rows2D(r, ixItem)))
        If rowLoc <> locQ Or rowItem <> itemQ Then GoTo ContinueLoop

        effDate = ParseYmdDate(rows2D(r, ixEff))
        If effDate = 0 Or effDate > cutoffDt Then GoTo ContinueLoop

        code = Trim$(CStr(rows2D(r, ixCode)))
        If code = "" Then GoTo ContinueLoop

        If Not bestRow.Exists(code) Then
            bestRow.Add code, r
            bestEff.Add code, effDate
        Else
            If effDate > bestEff(code) Then
                bestRow(code) = r
                bestEff(code) = effDate
            End If
        End If

ContinueLoop:
    Next r

    If bestRow.Count = 0 Then
        BuildUsageFromCache = Empty
        Exit Function
    End If

    Dim out() As Variant, i As Long
    ReDim out(0 To bestRow.Count - 1, 0 To 6)

    Dim k As Variant, rr As Long
    Dim prevD As Date, calD As Date, nextD As Date
    Dim periodY As Long

    i = 0
    For Each k In bestRow.Keys
        rr = bestRow(k)

        prevD = ParseYmdDate(rows2D(rr, ixPrev))
        calD = ParseYmdDate(rows2D(rr, ixCal))
        nextD = ParseYmdDate(rows2D(rr, ixNext))
        periodY = ToLongSafe(rows2D(rr, ixPeriod))

        Dim calOut As Variant, nextOut As Variant
        calOut = "N/A": nextOut = "N/A"

        If calD <> 0 And DateValue(cutoffDt) < calD Then
            If prevD <> 0 And periodY > 0 Then
                calOut = prevD
                nextOut = DateAdd("d", -1, DateAdd("yyyy", periodY, prevD))
            End If
        Else
            If calD <> 0 Then calOut = calD
            If nextD <> 0 Then nextOut = nextD
        End If

        out(i, 0) = rows2D(rr, ixInst)
        out(i, 1) = rows2D(rr, ixModel)
        out(i, 2) = rows2D(rr, ixMaker)
        out(i, 3) = rows2D(rr, ixSerial)
        out(i, 4) = calOut
        out(i, 5) = nextOut
        out(i, 6) = rows2D(rr, ixPeriod)

        i = i + 1
    Next k

    BuildUsageFromCache = out
End Function


Public Sub UpdateInstruments(ByVal prefix As String)
    Dim ws           As Worksheet
    Dim loc          As String
    Dim cutoffDt     As Date
    Dim hasCutoff    As Boolean
    Dim envRg        As Range, dateRg As Range
    Dim dtMin        As Date
    Dim cutoffStr    As String
    Dim targetRange  As Range
    Dim stnm         As String
    Dim localOut As Variant
    Dim output() As Variant
    Dim idx As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ActiveSheet
    loc = ws.Range(prefix & "_LOCATION").Value
    Set envRg = ws.Range(prefix & "_ENV")
    Set dateRg = envRg.Offset(1, 3).Resize(envRg.Rows.Count - 1, 1)

    Set targetRange = ws.Range(prefix & "_INSTRUMENTS")
    If targetRange.Rows.Count > 1 Then
        targetRange.Offset(1, 0) _
          .Resize(targetRange.Rows.Count - 1, targetRange.Columns.Count - 1).ClearContents
    End If

    If loc = "" Then
        GoTo Cleanup
    End If

    If Application.Count(dateRg) > 0 Then
        cutoffDt = MinValidDate(dateRg)
        hasCutoff = True
        cutoffStr = Format(cutoffDt, "yyyy-MM-dd")
    Else
        hasCutoff = False
        cutoffDt = Date
    End If

    cutoffStr = IIf(hasCutoff, Format(cutoffDt, "yyyy-MM-dd"), "")

    If Left(ActiveSheet.Range("STD").Value, 2) = "KN" Or Left(ActiveSheet.Range("STD").Value, 2) = "KS" Then
        stnm = "Usage"
    ElseIf Left(ActiveSheet.Range("STD").Value, 3) = "FCC" Then
        stnm = "Usage_FCC"
    Else
        stnm = "Usage_FCC"
    End If

    localOut = BuildUsageFromCache(stnm, loc, prefix, cutoffDt)

    If IsEmpty(localOut) Then GoTo Cleanup

    Dim writeRowCount As Long
    Dim i As Long
    Dim startRow As Long

    output = localOut

    startRow = targetRange.Row + 1
    writeRowCount = UBound(output, 1) + 1

    Dim rawcDate   As String
    Dim onlycDate  As Date
    Dim rawnDate   As String
    Dim onlynDate  As Date

    Dim currentDataRows As Long
    Dim additionalRowsNeeded As Long
    currentDataRows = targetRange.Rows.Count - 1
    If currentDataRows < writeRowCount Then
        additionalRowsNeeded = writeRowCount - currentDataRows
        Call ExpandRange(ws, targetRange, prefix & "_INSTRUMENTS", additionalRowsNeeded)
        Set targetRange = ws.Range(prefix & "_INSTRUMENTS")
    End If

    For i = 0 To writeRowCount - 1
        ws.Cells(startRow + i, 2).Value = output(i, 0)
        ws.Cells(startRow + i, 4).Value = output(i, 1)
        ws.Cells(startRow + i, 6).Value = output(i, 2)
        ws.Cells(startRow + i, 7).Value = output(i, 3)
        rawcDate = output(i, 4)
        If IsDate(rawcDate) Then
            onlycDate = CDate(rawcDate)
            ws.Cells(startRow + i, 8).Value = onlycDate
            ws.Cells(startRow + i, 8).NumberFormat = "yyyy-mm-dd"
        Else
            ws.Cells(startRow + i, 8).Value = "N/A"
        End If

        rawnDate = output(i, 5)
        If IsDate(rawnDate) Then
            onlynDate = CDate(rawnDate)
            ws.Cells(startRow + i, 9).Value = onlynDate
            ws.Cells(startRow + i, 9).NumberFormat = "yyyy-mm-dd"
        Else
            ws.Cells(startRow + i, 9).Value = "N/A"
        End If
        ws.Cells(startRow + i, 10).Value = output(i, 6)
        ws.Cells(startRow + i, 11).Value = 0
    Next i

    CheckValidation prefix

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Public Sub CheckValidation(ByVal prefix As String)
    Dim ws As Worksheet
    Dim envRg As Range, insRg As Range
    Dim dateRg As Range, instDateRg As Range, periodRg As Range
    Dim minDate As Date, maxDate As Date
    Dim nextCalDate As Date, calibrationPeriod As Double
    Dim startValidDate As Date, endValidDate As Date
    Dim cell As Range

    Set ws = ActiveSheet
    Set envRg = ws.Range(prefix & "_ENV")
    Set insRg = ws.Range(prefix & "_INSTRUMENTS")

    Set dateRg = envRg.Offset(1, 3).Resize(envRg.Rows.Count - 1, 1)
    Set instDateRg = insRg.Offset(1, 7).Resize(insRg.Rows.Count - 1, 1)
    Set periodRg = insRg.Offset(1, 8).Resize(insRg.Rows.Count - 1, 1)
    instDateRg.Interior.ColorIndex = xlNone

    If Application.Count(dateRg) > 0 Then
        minDate = Application.Min(dateRg)
        maxDate = Application.Max(dateRg)
    Else
        Exit Sub
    End If

    For Each cell In instDateRg
        If IsDate(cell.Value) Then
            nextCalDate = CDate(cell.Value)
            calibrationPeriod = cell.Offset(0, 1).Value

            If IsNumeric(calibrationPeriod) Then
                startValidDate = DateAdd("yyyy", -calibrationPeriod, nextCalDate) + 1
                endValidDate = nextCalDate

                If minDate < startValidDate Or maxDate > endValidDate Then
                    cell.Interior.Color = RGB(255, 0, 0)
                Else
                    cell.Interior.ColorIndex = xlNone
                End If
            End If
        End If
    Next cell
End Sub


Public Sub DebugPick(ByVal sheetName As String, ByVal locQ As String, ByVal itemQ As String, ByVal cutoffDt As Date, ByVal codeQ As String)
    EnsureUsageStores
    Call FetchUsageDumpCached(sheetName)

    Dim rows2D As Variant, hdrIdx As Object
    rows2D = gRowsBySheet(sheetName)
    Set hdrIdx = gHdrBySheet(sheetName)

    Dim ixEff As Long, ixLoc As Long, ixItem As Long, ixCode As Long
    ixEff = hdrIdx("EffectiveDate")
    ixLoc = hdrIdx("Location")
    ixItem = hdrIdx("Item")
    ixCode = hdrIdx("Code")

    Dim r As Long, effDate As Date

    Debug.Print "==== DebugPick ===="
    Debug.Print "sheet=" & sheetName & " loc=[" & locQ & "] item=[" & itemQ & "] code=[" & codeQ & "] cutoff=" & Format(DateValue(cutoffDt), "yyyy-mm-dd")

    For r = LBound(rows2D, 1) To UBound(rows2D, 1)
        If NormText(rows2D(r, ixLoc)) = NormText(locQ) _
           And NormText(rows2D(r, ixItem)) = NormText(itemQ) _
           And NormText(rows2D(r, ixCode)) = NormText(codeQ) Then

            effDate = ParseYmdDate(rows2D(r, ixEff))

            Debug.Print "r=" & r & _
                        " effRaw=" & CStr(rows2D(r, ixEff)) & _
                        " eff=" & IIf(effDate = 0, "0", Format(effDate, "yyyy-mm-dd")) & _
                        " <=cutoff? " & IIf(effDate <> 0 And effDate <= DateValue(cutoffDt), "OK", "SKIP")
        End If
    Next r

    Debug.Print "==== /DebugPick ===="
End Sub

Public Function NormText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ChrW(160), " ")
    s = Trim$(s)
    NormText = s
End Function

Private Function ToLongSafe(ByVal v As Variant) As Long
    Dim s As String, i As Long, ch As String, num As String
    s = Trim$(CStr(v))
    If s = "" Then ToLongSafe = 0: Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch = "-" And num = "") Then
            num = num & ch
        End If
    Next

    If num = "" Or num = "-" Then
        ToLongSafe = 0
    Else
        ToLongSafe = CLng(num)
    End If
End Function
