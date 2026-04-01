Attribute VB_Name = "LineModule"

Sub ExpandMultipleRanges()

    Dim selectedCell As Range
    Dim ws As Worksheet
    Dim rangeNames As Variant
    Dim suffixes As Variant
    Dim rangeName As Variant
    Dim suffix As Variant
    Dim rg As Range
    Dim rangeFullName As String
    Dim v As Variant
    Dim cnt As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Volatile False
    Application.CutCopyMode = False
    UnprotectCurrentSheet
    

    ' 현재 워크시트 설정
    Set ws = ActiveSheet

    ' 처리할 이름 범위와 접미사 목록
    rangeNames = standardFunction(ActiveSheet.Range("STD").value)
    suffixes = Array("_ENV", "_COMMENTS", "_INSTRUMENTS")
    
    ' 접미사가 없는 범위 처리 (필요한 범위가 있다면 배열에 추가)
    Dim noSuffixRanges As Variant
    noSuffixRanges = Array("Total_Config", "System_Config", "Connection_Cables", "OPERATING_MODE", "Test_Voltage", "EUT_ports")

    ' 선택한 셀 가져오기
    Set selectedCell = Selection
    If selectedCell Is Nothing Then Exit Sub
    If TypeName(selectedCell) <> "Range" Then Exit Sub  ' 셀 선택이 아닐 때

    v = selectedCell.cells(1, 1).value  ' 다중 선택이면 첫 셀만 사용
    If IsNumeric(v) Then
        cnt = CLng(v)   ' 정수로 쓸 거면 CLng, 소수 유지하려면 CDbl
    Else
        cnt = 1
    End If
    ' 접미사가 없는 범위 처리
    For Each rangeName In noSuffixRanges
        rangeFullName = rangeName
        Set rg = IsCellInRange(ws, selectedCell, rangeFullName)
        If Not rg Is Nothing Then
            Call ExpandRange(ws, rg, rangeFullName, cnt)
            GoTo Err:
        End If
    Next rangeName

    ' 접미사가 있는 범위 처리
    For Each rangeName In rangeNames
        For Each suffix In suffixes
            ' 기본 범위 이름 생성 및 검사
            rangeFullName = rangeName & suffix
            Set rg = IsCellInRange(ws, selectedCell, rangeFullName)
            If Not rg Is Nothing Then
                Call ExpandRange(ws, rg, rangeFullName, 1)
                GoTo Err:
            End If
            
        Next suffix
    Next rangeName
    
    rangeNames = DataSuffix(ActiveSheet.Range("STD").value)
    For Each rangeName In rangeNames
        rangeFullName = rangeName & "_RESULT"
        Set rg = IsCellInRange(ws, selectedCell, rangeFullName)
        If Not rg Is Nothing Then
            Call ExpandRange(ws, rg, rangeFullName, 1)
            GoTo Err:
        End If
    Next rangeName

Err:
    ProtectCurrentSheet
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Volatile True

End Sub


Sub DeleteLastRowInRange()

    Dim selectedCell As Range
    Dim ws As Worksheet
    Dim rangeNames As Variant
    Dim suffixes As Variant
    Dim noSuffixRanges As Variant
    Dim minRows As Object
    Dim rangeName As Variant
    Dim suffix As Variant
    Dim rg As Range
    Dim rangeFullName As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Volatile False
    UnprotectCurrentSheet

    ' 현재 워크시트 설정
    Set ws = ActiveSheet

    ' 처리할 이름 범위와 접미사 목록
    rangeNames = standardFunction(ws.Range("STD").value)
    suffixes = Array("_ENV", "_COMMENTS", "_INSTRUMENTS")
    noSuffixRanges = Array("Total_Config", "System_Config", "Connection_Cables", "OPERATING_MODE", "Test_Voltage", "EUT_ports")

    ' 최소 행 개수 정의
    Set minRows = CreateObject("Scripting.Dictionary")
    minRows("_ENV") = 2
    minRows("_COMMENTS") = 1
    minRows("_INSTRUMENTS") = 3
    minRows("Total_Config") = 3
    minRows("System_Config") = 3
    minRows("Connection_Cables") = 4
    minRows("OPERATING_MODE") = 2
    minRows("Test_Voltage") = 3
    minRows("EUT_ports") = 3

    ' 선택한 셀 가져오기
    Set selectedCell = Selection
    If selectedCell Is Nothing Then Exit Sub

    ' 접미사가 없는 범위 처리
    For Each rangeName In noSuffixRanges
        rangeFullName = rangeName
        Set rg = IsCellInRange(ws, selectedCell, rangeFullName)
        If Not rg Is Nothing Then
            Call DeleteSelectedRow(ws, rg, rangeFullName, minRows(rangeName))
            GoTo Err
        End If
    Next rangeName

    ' 접미사가 있는 범위 처리
    For Each rangeName In rangeNames
        For Each suffix In suffixes
            ' 기본 범위 이름 생성 및 검사
            rangeFullName = rangeName & suffix
            Set rg = IsCellInRange(ws, selectedCell, rangeFullName)
            If Not rg Is Nothing Then
                Call DeleteSelectedRow(ws, rg, rangeFullName, minRows(suffix))
                GoTo Err
            End If
        Next suffix
    Next rangeName
    
    ' DATA 영역
    rangeNames = DataSuffix(ActiveSheet.Range("STD").value)
    For Each rangeName In rangeNames
        For Each suffix In suffixes
            ' 기본 범위 이름 생성 및 검사
            rangeFullName = rangeName & "_RESULT"
            Set rg = IsCellInRange(ws, selectedCell, rangeFullName)
            If Not rg Is Nothing Then
                Call DeleteSelectedRow(ws, rg, rangeFullName, 2)
                GoTo Err
            End If
        Next suffix
    Next rangeName
    
Err:
    ProtectCurrentSheet
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Volatile True
    
End Sub


' 범위 내 셀이 포함되어 있는지 검사하는 함수
Function IsCellInRange(ws As Worksheet, selectedCell As Range, rangeName As String) As Range

    Dim rg As Range
    On Error Resume Next
    Set rg = ws.Range(rangeName)
    On Error GoTo 0

    If Not rg Is Nothing Then
        If Not Intersect(selectedCell, rg) Is Nothing Then
            Set IsCellInRange = rg
        Else
            Set IsCellInRange = Nothing
        End If
    Else
        Set IsCellInRange = Nothing
    End If

End Function

' 행을 추가하고 범위를 확장하는 함수
Sub ExpandRange(ws As Worksheet, rg As Range, rangeName As String, rowsToAdd As Long)
    
    Dim scr As Boolean, calc As XlCalculation, evt As Boolean
    Dim insertStart As Long, lastRowBefore As Long
    Dim insRows As Range, newRg As Range
    Dim nm As Name
    
    If rowsToAdd <= 0 Then rowsToAdd = 1
    
    ' --- 성능 옵션 저장/비활성화 ---
    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
    evt = Application.EnableEvents:    Application.EnableEvents = False
    
    ' 현재 rg 바로 아래부터 rowsToAdd 만큼 "한 번에" 행 추가
    insertStart = rg.Row + rg.rows.Count
    lastRowBefore = insertStart - 1
    Set insRows = ws.rows(insertStart).Resize(rowsToAdd)
    insRows.Insert Shift:=xlShiftDown
    
    ' 이름 재정의는 한 번만 (시트 범위명 기준)
    Set newRg = ws.Range(rg.Resize(rg.rows.Count + rowsToAdd, rg.Columns.Count).Address)
    
    On Error Resume Next
    Set nm = ws.Names(rangeName)              ' 시트 범위명
    On Error GoTo 0
    
    If nm Is Nothing Then
        ws.Names.Add Name:=rangeName, RefersTo:="=" & newRg.Address(External:=True)
    Else
        nm.RefersTo = "=" & newRg.Address(External:=True)
    End If
    
    ' 서식만 일괄 복사 (값/수식은 복사하지 않음)
    ws.rows(lastRowBefore).Copy
    ws.rows(insertStart).Resize(rowsToAdd).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' 호출 측에서도 확장된 범위를 쓸 수 있게 rg 갱신(ByRef)
    Set rg = newRg
    rg.cells(rg.rows.Count, 1).Select

SafeExit:
    ' --- 성능 옵션 복구 ---
    Application.EnableEvents = evt
    Application.Calculation = calc
    Application.ScreenUpdating = scr
    
End Sub


Sub DeleteSelectedRow(ws As Worksheet, rg As Range, rangeName As String, minRowCount As Integer)
    
    Dim selectedRow As Long
    
    If Not Intersect(ActiveCell, rg) Is Nothing Then
        selectedRow = ActiveCell.Row
        If selectedRow = rg.Row Or selectedRow = rg.Row + 1 Then

        ElseIf rg.rows.Count > minRowCount Then
            ws.rows(selectedRow).Delete
        End If
    End If
    rg.cells(rg.rows.Count, 1).Select
        
End Sub
