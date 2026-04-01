Attribute VB_Name = "update_module"
Sub CheckVersionUpdate()
    Dim raw        As String
    Dim parsed     As Collection
    Dim headers    As Collection
    Dim dataRow    As Collection
    Dim i          As Long
    Dim verCol     As Long
    Dim extVersion As String
    Dim curVersion As String

    curVersion = GetVersion()

    On Error GoTo ErrHandler
    raw = CallGasApi("base_url", "key=" & GetApiKey() & "&sheet=Version")
    On Error GoTo 0

    ' 3) JSON Parse
    Set parsed = JsonConverter.ParseJson(raw)
    If parsed.Count < 2 Then
        Exit Sub
    End If

    Set headers = parsed(1)
    Set dataRow = parsed(parsed.Count)

    verCol = 0
    For i = 1 To headers.Count
        If headers(i) = "Version" Then
            verCol = i
            Exit For
        End If
    Next
    If verCol = 0 Then
        Exit Sub
    End If

    extVersion = CStr(dataRow(verCol))

    If curVersion <> extVersion Then
        MsgBox "Current: " & curVersion & " / Latest: " & extVersion, vbInformation
        Sheet1.Visible = xlSheetVeryHidden
    End If
    Exit Sub

ErrHandler:
    MsgBox "Version check failed: " & Err.Description, vbExclamation
End Sub



Sub ProcessVersionUpdate(curVersion As String)
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strConn As String, sql As String
    Dim curWS As Worksheet
    Dim lastRow As Long, newRow As Long
    Dim currentVersionFound As Boolean
    Dim extVersion As String
    Dim eventVal As Variant, targetVal As Variant

    Set curWS = ThisWorkbook.Sheets("Version")
    lastRow = curWS.Cells(curWS.Rows.Count, "A").End(xlUp).Row
    newRow = lastRow + 1

    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & CStr(GetCfg("paths.original_copy")) & ";" & _
              "Extended Properties=""Excel 12.0 Macro;HDR=Yes;IMEX=1"";"

    Set conn = New ADODB.Connection
    On Error GoTo ErrHandler
    conn.Open strConn

    sql = "SELECT Version, event, target FROM [Version$]"
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    currentVersionFound = False
    Do Until rs.EOF
        extVersion = rs.Fields("Version").Value
        If currentVersionFound Then
            eventVal = rs.Fields("event").Value
            targetVal = rs.Fields("target").Value

            Call ProcessEventTarget(CStr(eventVal), CStr(targetVal))

            curWS.Cells(newRow, "A").Value = extVersion
            curWS.Cells(newRow, "B").Value = eventVal
            curWS.Cells(newRow, "C").Value = targetVal
            newRow = newRow + 1
        ElseIf extVersion = curVersion Then
            currentVersionFound = True
        End If
        rs.MoveNext
    Loop

    rs.Close
    conn.Close

    Application.StatusBar = "Version update completed."
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
End Sub




Sub ProcessEventTarget(eventVal As String, targetVal As String)

    Select Case eventVal
        Case "sheetUP"
            If SheetExists(targetVal) Then
                Application.DisplayAlerts = False
                ThisWorkbook.Sheets(targetVal).Visible = True
                ThisWorkbook.Sheets(targetVal).Delete
                Application.DisplayAlerts = True
            End If
            CopySheetFromOriginal targetVal

        Case "sheetADD"
            CopySheetFromOriginal targetVal

        Case "codeUP"
            RemoveModule targetVal
            ImportModule targetVal

        Case "codeADD"
            ImportModule targetVal

        Case Else
            Debug.Print "Unknown event: " & eventVal
    End Select

End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Sub CopySheetFromOriginal(sheetName As String)
    Dim wbOrig As Workbook
    Dim wsOrig As Worksheet
    Dim wsCopied As Worksheet
    Dim filePath As String

    filePath = CStr(GetCfg("paths.original_copy"))

    Set wbOrig = Workbooks.Open(filePath, ReadOnly:=True)
    On Error Resume Next
    Set wsOrig = wbOrig.Sheets(sheetName)
    On Error GoTo 0

    If Not wsOrig Is Nothing Then
        wsOrig.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set wsCopied = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wsCopied.Visible = xlSheetVeryHidden
        Dim links As Variant, i As Long
        links = ThisWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
        If Not IsEmpty(links) Then
            For i = LBound(links) To UBound(links)
                ThisWorkbook.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
            Next i
        End If
    Else
        MsgBox "Sheet '" & sheetName & "' not found in original file.", vbExclamation
    End If

    wbOrig.Close SaveChanges:=False
End Sub

Sub RemoveModule(moduleName As String)
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0
    If Not vbComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If
End Sub

Sub ImportModule(moduleName As String)
    Dim modulePath As String
    modulePath = CStr(GetCfg("paths.bas_dir")) & "\" & moduleName & ".bas"
    ThisWorkbook.VBProject.VBComponents.Import modulePath
End Sub
