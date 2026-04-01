Attribute VB_Name = "personalDB_module"
Option Explicit

Sub CallLoadDBUserForm()
    LoadDB.Show
End Sub

Sub loadDBByOrderNo(ByVal orderNoParam As String)

    Dim sConn As String, strSQL As String
    Dim cN As Object, rs As Object
    Dim opModeStr As String, opModeCommentStr As String
    Dim totalConfigStr As String, systemConfigStr As String, connCablesStr As String
    Dim jsonData As Variant, itm As Variant
    Dim targetRange As Range, iRow As Long
    Dim totalConfigData As Variant, sysConfigData As Variant, connCablesData As Variant
    Dim requiredRows As Long, currentRows As Long, rowsToAdd As Long
    Dim ws As Worksheet
    Dim col As Variant, r As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set ws = ActiveSheet

    sConn = GetPersonalDBConn()

    On Error GoTo ErrHandler
    Set cN = CreateObject("ADODB.Connection")
    cN.Open sConn

    strSQL = "SELECT OPERATING_MODE, OPERATING_MODE_COMMENT, Total_Config, System_Config, Connection_Cables " & _
             "FROM Personal_DB WHERE Order_No = '" & Replace(orderNoParam, "'", "''") & "'"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, cN, 1, 1

    If Not rs.EOF Then
        opModeStr = rs.Fields("OPERATING_MODE").Value
        opModeCommentStr = rs.Fields("OPERATING_MODE_COMMENT").Value
        totalConfigStr = rs.Fields("Total_Config").Value
        systemConfigStr = rs.Fields("System_Config").Value
        connCablesStr = rs.Fields("Connection_Cables").Value
    Else
        MsgBox "No matching data found."
        GoTo Cleanup
    End If

    rs.Close
    cN.Close

    ' ===== OPERATING_MODE =====
    Set jsonData = JsonConverter.ParseJson(opModeStr)
    Set targetRange = ws.Range("OPERATING_MODE")
    If TypeName(jsonData) = "Collection" Then
        requiredRows = jsonData.Count
        currentRows = targetRange.Rows.Count
        If requiredRows > currentRows Then
            rowsToAdd = requiredRows - currentRows
            Call ExpandRange(ws, targetRange, "OPERATING_MODE", rowsToAdd + 1)
            Set targetRange = ws.Range("OPERATING_MODE")
        End If
        iRow = 1
        For Each itm In jsonData
            targetRange.Cells(iRow, 1).Value = itm("No")
            targetRange.Cells(iRow, 2).Value = itm("Name")
            targetRange.Cells(iRow, 4).Value = itm("Description")
            iRow = iRow + 1
        Next itm
    ElseIf TypeName(jsonData) = "Dictionary" Then
        targetRange.Cells(1, 1).Value = jsonData("No")
        targetRange.Cells(1, 2).Value = jsonData("Name")
        targetRange.Cells(1, 4).Value = jsonData("Description")
    End If

    ' ===== OPERATING_MODE_COMMENT =====
    Set targetRange = ws.Range("OPERATING_MODE_COMMENT")
    targetRange.Value = opModeCommentStr

    ' ===== Total_Config =====
    Call LoadConfigRange(ws, totalConfigStr, "Total_Config")

    ' ===== System_Config =====
    Call LoadConfigRange(ws, systemConfigStr, "System_Config")

    ' ===== Connection_Cables =====
    Set connCablesData = JsonConverter.ParseJson(connCablesStr)
    Set targetRange = ws.Range("Connection_Cables")
    If TypeName(connCablesData) = "Collection" Then
        requiredRows = connCablesData.Count
        currentRows = targetRange.Rows.Count - 2
        If requiredRows > currentRows Then
            rowsToAdd = requiredRows - currentRows
            Call ExpandRange(ws, targetRange, "Connection_Cables", rowsToAdd)
            Set targetRange = ws.Range("Connection_Cables")
        End If
        r = 3
        For Each itm In connCablesData
            For Each col In Array(1, 3, 5, 7, 9, 10)
                Dim keyStr As String
                keyStr = "Col" & col
                If itm.Exists(keyStr) Then
                    targetRange.Cells(r, col).Value = itm(keyStr)
                End If
            Next col
            r = r + 1
        Next itm
    ElseIf TypeName(connCablesData) = "Dictionary" Then
        For Each col In Array(1, 3, 5, 7, 9, 10)
            Dim keyStr2 As String
            keyStr2 = "Col" & col
            If connCablesData.Exists(keyStr2) Then
                targetRange.Cells(3, col).Value = connCablesData(keyStr2)
            End If
        Next col
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = True
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cN Is Nothing Then If cN.State = 1 Then cN.Close
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

' ── Shared helper: Load Total_Config / System_Config from JSON ──
Private Sub LoadConfigRange(ws As Worksheet, ByVal jsonStr As String, ByVal rangeName As String)
    Dim configData As Variant
    Dim targetRange As Range
    Dim colIndices As Variant
    Dim requiredRows As Long, currentRows As Long, rowsToAdd As Long
    Dim r As Long, idx As Long
    Dim itm As Variant, col As Variant
    Dim header As String

    Set configData = JsonConverter.ParseJson(jsonStr)
    Set targetRange = ws.Range(rangeName)
    colIndices = Array(1, 3, 5, 7, 9)

    If TypeName(configData) = "Collection" Then
        requiredRows = configData.Count
        currentRows = targetRange.Rows.Count - 2
        If requiredRows > currentRows Then
            rowsToAdd = requiredRows - currentRows
            Call ExpandRange(ws, targetRange, rangeName, rowsToAdd + 1)
            Set targetRange = ws.Range(rangeName)
        End If
        r = 3
        For Each itm In configData
            For Each col In colIndices
                header = targetRange.Cells(1, CLng(col)).Value
                If itm.Exists(header) Then
                    targetRange.Cells(r, CLng(col)).Value = itm(header)
                End If
            Next col
            r = r + 1
        Next itm
    ElseIf TypeName(configData) = "Dictionary" Then
        For Each col In colIndices
            header = targetRange.Cells(1, CLng(col)).Value
            If configData.Exists(header) Then
                targetRange.Cells(2, CLng(col)).Value = configData(header)
            End If
        Next col
    End If
End Sub


Sub SavetoDB()

    Dim sConn As String, strSQL As String
    Dim cN As Object, rs As Object
    Dim Order_No As String
    Dim opModeJSON As String, opModeComment As String
    Dim totalConfigJSON As String, systemConfigJSON As String, connectionCablesJSON As String
    Dim response As VbMsgBoxResult
    Dim Applicant As String, Model_Name As String, Product_Name As String

    Order_No = ActiveSheet.Range("Order_No").Value
    If Order_No = "" Or Left(Order_No, 4) <> "DTNC" Then
        MsgBox "Please check Order No."
        Exit Sub
    End If
    opModeComment = ActiveSheet.Range("OPERATING_MODE_COMMENT").Value
    Applicant = ActiveSheet.Range("Applicant").Value
    Model_Name = ActiveSheet.Range("Model_Name").Value
    Product_Name = ActiveSheet.Range("Product_Name").Value

    opModeJSON = BuildOperatingModeJSON()
    totalConfigJSON = BuildRangeJSON("Total_Config", Array(1, 3, 5, 7, 9), 3, True)
    systemConfigJSON = BuildRangeJSON("System_Config", Array(1, 3, 5, 7, 9), 3, True)
    connectionCablesJSON = BuildRangeJSON("Connection_Cables", Array(1, 3, 5, 7, 9, 10), 3, False)

    sConn = GetPersonalDBConn()

    Set cN = CreateObject("ADODB.Connection")
    cN.Open sConn

    strSQL = "SELECT COUNT(*) AS RecordCount FROM Personal_DB WHERE Order_No = '" & Replace(Order_No, "'", "''") & "'"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, cN, 1, 1

    If Not rs.EOF Then
        If rs.Fields("RecordCount").Value > 0 Then
            response = MsgBox("Order_No " & Order_No & " already exists. Update?", vbYesNo + vbQuestion)
            If response = vbYes Then
                strSQL = "UPDATE Personal_DB SET " & _
                         "Applicant = '" & Replace(Applicant, "'", "''") & "', " & _
                         "Model_Name = '" & Replace(Model_Name, "'", "''") & "', " & _
                         "Product_Name = '" & Replace(Product_Name, "'", "''") & "', " & _
                         "OPERATING_MODE = '" & Replace(opModeJSON, "'", "''") & "', " & _
                         "OPERATING_MODE_COMMENT = '" & Replace(opModeComment, "'", "''") & "', " & _
                         "Total_Config = '" & Replace(totalConfigJSON, "'", "''") & "', " & _
                         "System_Config = '" & Replace(systemConfigJSON, "'", "''") & "', " & _
                         "Connection_Cables = '" & Replace(connectionCablesJSON, "'", "''") & "' " & _
                         "WHERE Order_No = '" & Replace(Order_No, "'", "''") & "'"
                cN.Execute strSQL
                MsgBox "Data updated successfully."
            Else
                MsgBox "Update cancelled."
            End If
        Else
            strSQL = "INSERT INTO Personal_DB (Order_No, Applicant, Model_Name, Product_Name, OPERATING_MODE, OPERATING_MODE_COMMENT, Total_Config, System_Config, Connection_Cables) " & _
                     "VALUES ('" & Replace(Order_No, "'", "''") & "', " & _
                             "'" & Replace(Applicant, "'", "''") & "', " & _
                             "'" & Replace(Model_Name, "'", "''") & "', " & _
                             "'" & Replace(Product_Name, "'", "''") & "', " & _
                             "'" & Replace(opModeJSON, "'", "''") & "', " & _
                             "'" & Replace(opModeComment, "'", "''") & "', " & _
                             "'" & Replace(totalConfigJSON, "'", "''") & "', " & _
                             "'" & Replace(systemConfigJSON, "'", "''") & "', " & _
                             "'" & Replace(connectionCablesJSON, "'", "''") & "')"
            cN.Execute strSQL
            MsgBox "Data saved successfully."
        End If
    End If

    rs.Close
    Set rs = Nothing
    cN.Close
    Set cN = Nothing

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ── OPERATING_MODE JSON builder ──
Function BuildOperatingModeJSON() As String
    Dim rng As Range, r As Range
    Dim jsonItems As Collection, d As Object

    On Error GoTo ErrHandler
    Set rng = ActiveSheet.Range("OPERATING_MODE")
    Set jsonItems = New Collection

    For Each r In rng.Rows
        If Application.WorksheetFunction.CountA(r) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            d.Add "No", r.Cells(1, 1).Value
            d.Add "Name", r.Cells(1, 2).Value
            d.Add "Description", r.Cells(1, 4).Value
            jsonItems.Add d
        End If
    Next r

    BuildOperatingModeJSON = ConvertToJson(jsonItems)
    Exit Function

ErrHandler:
    BuildOperatingModeJSON = ""
End Function

' ── Generic range-to-JSON builder (replaces 3 near-identical functions) ──
' rangeName: Named range to read
' colIndices: Array of column indices to include
' startRow: First data row (relative to range)
' useHeaders: If True, use row 1 as key names; if False, use "Col{n}"
Function BuildRangeJSON(ByVal rangeName As String, _
                        ByVal colIndices As Variant, _
                        ByVal startRow As Long, _
                        ByVal useHeaders As Boolean) As String
    Dim rng As Range
    Dim jsonItems As Collection
    Dim d As Object
    Dim iRow As Long, idx As Long
    Dim keyVal As String

    On Error GoTo ErrHandler
    Set rng = ActiveSheet.Range(rangeName)
    Set jsonItems = New Collection

    For iRow = startRow To rng.Rows.Count
        If Application.WorksheetFunction.CountA(rng.Rows(iRow)) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            For idx = LBound(colIndices) To UBound(colIndices)
                If useHeaders Then
                    keyVal = CStr(rng.Cells(1, CLng(colIndices(idx))).Value)
                Else
                    keyVal = "Col" & colIndices(idx)
                End If
                d.Add keyVal, rng.Cells(iRow, CLng(colIndices(idx))).Value
            Next idx
            jsonItems.Add d
        End If
    Next iRow

    BuildRangeJSON = ConvertToJson(jsonItems)
    Exit Function

ErrHandler:
    BuildRangeJSON = ""
End Function
