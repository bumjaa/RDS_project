Attribute VB_Name = "uiux_module"
'------------------------------------------------------------
' 범용 포맷팅: 지정한 범위의 글자색·테두리색을 clr로 변경
'------------------------------------------------------------
Public Sub Format_BaseRange(rng As Range, clr As Long)
    Dim cel As Range
    Dim bTypes As Variant, b As Variant
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 글자색
    rng.Font.Color = clr

    ' 외곽·대각선 테두리
    bTypes = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                   xlDiagonalDown, xlDiagonalUp)
    For Each cel In rng.cells
        For Each b In bTypes
            With cel.Borders(b)
                If .LineStyle <> xlNone Then .Color = clr
            End With
        Next b
    Next cel

    ' 내부 수평·수직선
    With rng.Borders(xlInsideHorizontal)
        If .LineStyle <> xlNone Then .Color = clr
    End With
    With rng.Borders(xlInsideVertical)
        If .LineStyle <> xlNone Then .Color = clr
    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------
' 체크박스 클릭 시 호출되는 매크로
' (폼 컨트롤 체크박스에 이 매크로를 연결하세요)
'------------------------------------------------------------
Public Sub CheckBox_Click()
    Dim cbName As String
    Dim prefix As String
    Dim clr As Long
    Dim targetRng As Range

    ' 호출한 컨트롤 이름을 가져와 prefix 로 사용
    cbName = Application.Caller
    prefix = cbName  ' “CE”, “RE” 등

    ' 폼 컨트롤은 Value = xlOn (1) or xlOff (-4146)
    If ActiveSheet.CheckBoxes(cbName).value = xlOn Then
        clr = RGB(0, 0, 0)           ' 체크된 경우 검정
    Else
        clr = RGB(178, 178, 178)     ' 체크 해제된 경우 회색
    End If

    ' prefix_BASE 범위 가져와 포맷팅
    On Error Resume Next
    Set targetRng = ActiveSheet.Range(prefix & "_BASE")
    On Error GoTo 0

    If Not targetRng Is Nothing Then
        Format_BaseRange targetRng, clr
    End If
End Sub


