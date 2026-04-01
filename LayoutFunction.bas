Attribute VB_Name = "LayoutFunction"

Sub CallCreateAndGroupLinesFromString(paramStr As String)
    Dim paramStrArray() As String
    Dim numArgs() As Variant
    Dim i As Integer, argCount As Integer
    
    paramStrArray = Split(paramStr, ",")
    argCount = UBound(paramStrArray) - LBound(paramStrArray) + 1
    
    Select Case argCount
        Case 4, 6, 8, 10, 12
            ReDim numArgs(0 To argCount - 1)
            For i = 0 To argCount - 1
                numArgs(i) = CLng(Trim(paramStrArray(i)))
            Next i
            
            Select Case argCount
                Case 4
                    CreateAndGroupLines numArgs(0), numArgs(1), numArgs(2), numArgs(3)
                Case 6
                    CreateAndGroupLines numArgs(0), numArgs(1), numArgs(2), numArgs(3), numArgs(4), numArgs(5)
                Case 8
                    CreateAndGroupLines numArgs(0), numArgs(1), numArgs(2), numArgs(3), numArgs(4), numArgs(5), numArgs(6), numArgs(7)
                Case 10
                    CreateAndGroupLines numArgs(0), numArgs(1), numArgs(2), numArgs(3), numArgs(4), numArgs(5), numArgs(6), numArgs(7), numArgs(8), numArgs(9)
                Case 12
                    CreateAndGroupLines numArgs(0), numArgs(1), numArgs(2), numArgs(3), numArgs(4), numArgs(5), numArgs(6), numArgs(7), numArgs(8), numArgs(9), numArgs(10), numArgs(11)
            End Select
        Case Else
            
    End Select
End Sub


Sub CallEquipmentDrawFromString(p1 As String)
    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer
    
    params = Split(p1, ",")

    EqName = params(0)
    x = CInt(params(1))
    y = CInt(params(2))
    w = CInt(params(3))
    h = CInt(params(4))
    
    EquipmentDraw EqName, x, y, w, h
    
End Sub

Sub CallTextboxDrawFromString(p1 As String)
    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer
    
    params = Split(p1, ",")

    EqName = params(0)
    x = CInt(params(1))
    y = CInt(params(2))
    
    TextboxDraw EqName, x, y
    
End Sub


Sub CallPasteYLinesFromString(p1 As String)
    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer

    params = Split(p1, ",")
    
    x = CInt(params(0))
    y = CInt(params(1))
    w = CInt(params(2))
    h = CInt(params(3))
    
    PasteYLines x, y, w, h

End Sub

Sub CallPasteCLinesFromString(p1 As String)
    
    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer

    params = Split(p1, ",")
    
    x = CInt(params(0))
    y = CInt(params(1))
    w = CInt(params(2))
    h = CInt(params(3))
    
    PasteCLines x, y, w, h
    
End Sub

Sub CallPasteILinesFromString(p1 As String)
    
    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer

    params = Split(p1, ",")
    
    x = CInt(params(0))
    y = CInt(params(1))
    w = CInt(params(2))
    h = CInt(params(3))
    
    PasteILines x, y, w, h
    
End Sub

Sub CallPasteHedsetFromString(p1 As String)

    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer

    params = Split(p1, ",")
    
    x = CInt(params(0))
    y = CInt(params(1))
    w = CInt(params(2))
    h = CInt(params(3))
    
    PasteHedset x, y, w, h
    
End Sub


Sub CallPasteUSBFromString(p1 As String)

    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer

    params = Split(p1, ",")
    
    x = CInt(params(0))
    y = CInt(params(1))
    w = CInt(params(2))
    h = CInt(params(3))
    
    PasteUSB x, y, w, h
    
End Sub

Sub CallPasteKeyboardFromString(p1 As String)
    
    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer

    params = Split(p1, ",")
    
    x = CInt(params(0))
    y = CInt(params(1))
    w = CInt(params(2))
    h = CInt(params(3))
    
    PasteKeyboard x, y, w, h
    
End Sub

Sub CallPasteMouseFromString(p1 As String)
    
    Dim params() As String
    Dim EqName As String
    Dim x As Integer, y As Integer, w As Integer, h As Integer

    params = Split(p1, ",")
    
    x = CInt(params(0))
    y = CInt(params(1))
    w = CInt(params(2))
    h = CInt(params(3))
    
    PasteMouse x, y, w, h
    
End Sub
