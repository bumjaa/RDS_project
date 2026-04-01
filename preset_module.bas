Attribute VB_Name = "preset_module"

Sub clearSheet()

    Dim item1 As Variant
    Dim fixes As String
    Dim nonfixes As String
    Dim itm As Variant
    Dim check As String
    Dim ws As Worksheet
    Dim std As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ActiveSheet
    std = ws.Range("STD").Value

    item1 = standardFunction(std)
    fixes = "_CHECK"
    nonfixes = "_Remarks"

    On Error Resume Next
    For Each itm In item1
        check = itm & fixes
        Range(check).Value = ""
    Next

    For Each itm In item1
        Dim remarks As String
        remarks = itm & nonfixes
        Range(remarks).Value = ""
    Next
    On Error GoTo 0

    Range("G_Remarks").Value = ""

    Application.EnableEvents = True

    fixes = "_LOCATION"

    For Each itm In item1
        check = itm & fixes
        Range(check).Value = ""
    Next

    Application.ScreenUpdating = True

End Sub


Function ParseQuotedCSV(rawRow As String) As Variant
    Dim inQuote As Boolean
    Dim startPos As Long, i As Long
    Dim ch As String, field As String
    Dim values() As String, cnt As Long

    inQuote = False
    startPos = 1
    cnt = 0

    For i = 1 To Len(rawRow)
        ch = Mid$(rawRow, i, 1)
        If ch = """" Then
            inQuote = Not inQuote
        ElseIf ch = "," And Not inQuote Then
            field = Mid$(rawRow, startPos, i - startPos)
            field = Trim$(field)
            If Left$(field, 1) = """" And Right$(field, 1) = """" Then
                field = Mid$(field, 2, Len(field) - 2)
            End If
            ReDim Preserve values(0 To cnt)
            values(cnt) = field
            cnt = cnt + 1
            startPos = i + 1
        End If
    Next i

    field = Mid$(rawRow, startPos)
    field = Trim$(field)
    If Left$(field, 1) = """" And Right$(field, 1) = """" Then
        field = Mid$(field, 2, Len(field) - 2)
    End If
    ReDim Preserve values(0 To cnt)
    values(cnt) = field

    ParseQuotedCSV = values
End Function


Sub ProcessPresetFromDB(sheetName As String, std As String, preset As String, ins As String)

    Dim raw As String
    Dim rows As Variant, hdrFields As Variant
    Dim i As Long, j As Long
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrHandler
    raw = CallGasApi("base_url", "key=" & GetApiKey() & "&sheet=Test_Preset")
    On Error GoTo 0

    ' Parse: outer [] strip, split rows by ],[
    raw = Mid$(raw, 2, Len(raw) - 2)
    rows = Split(raw, "],[")

    ' Header field index mapping
    hdrFields = ParseQuotedCSV(Replace(Replace(rows(0), "[", ""), "]", ""))
    For j = 0 To UBound(hdrFields)
        idx(hdrFields(j)) = j
    Next j

    Call clearSheet
    Application.EnableEvents = False

    For i = 1 To UBound(rows)
        Dim cleanRow As String
        cleanRow = Replace(Replace(rows(i), "[", ""), "]", "")

        Dim fields As Variant
        fields = ParseQuotedCSV(cleanRow)

        If fields(idx("Std")) = std And fields(idx("Code")) = preset Then
            Dim ai As String, ni As String, ci As String, gi As String
            ai = fields(idx("applied_item"))
            ni = fields(idx("none_item"))
            ci = fields(idx("item_comment"))
            gi = fields(idx("g_remarks"))

            ' applied_item
            If Len(ai) > 0 Then
                Dim arrA As Variant, k As Long
                arrA = Split(ai, ",")
                For k = LBound(arrA) To UBound(arrA)
                    Dim nm As String
                    nm = Trim$(arrA(k))
                    On Error Resume Next
                    With Sheets(sheetName)
                        With .CheckBoxes(nm)
                            .Value = xlOn
                            If .OnAction <> "" Then
                                Application.Run .OnAction
                            End If
                        End With
                        .Range(nm & "_LOCATION").Value = ins
                    End With
                    On Error GoTo 0
                Next k
            End If

            ' none_item + item_comment
            If Len(ni) > 0 And Len(ci) > 0 Then
                Dim arrN As Variant, arrC As Variant
                arrN = Split(ni, ",")
                arrC = Split(ci, ",")
                If UBound(arrN) = UBound(arrC) Then
                    For k = LBound(arrN) To UBound(arrN)
                        Dim nm2 As String
                        nm2 = Trim$(arrN(k))
                        On Error Resume Next
                        Sheets(sheetName).Range(nm2 & "_Remarks").Value = Trim$(arrC(k))
                        On Error GoTo 0
                    Next k
                Else
                    MsgBox "none_item / item_comment count mismatch", vbExclamation
                End If
            End If

            ' g_remarks
            On Error Resume Next
            gi = Replace(gi, "\n", vbLf)
            With Sheets(sheetName).Range("G_Remarks")
                .Value = gi
                .WrapText = True
            End With
            On Error GoTo 0

            Exit For
        End If
    Next i

    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Preset load failed: " & Err.Description, vbExclamation
    Application.EnableEvents = True
End Sub
