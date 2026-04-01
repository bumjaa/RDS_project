Attribute VB_Name = "Layout_module"
Sub saveLayout()
    Dim defaultFolder As String
    Dim defaultFile As String
    Dim filePath As Variant
    Dim i As Long
    Dim fso As Object
    Dim rng As Range
    Dim tempSheet As Worksheet
    Dim tempChartObj As ChartObject

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' 기본 경로: 현재 워크북 경로의 \CP 폴더
    defaultFolder = ThisWorkbook.path & "\CP"
    
    ' FileSystemObject로 폴더 존재 여부 확인 후, 없으면 생성
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(defaultFolder) Then
        fso.CreateFolder defaultFolder
    End If
    
    ' 기본 파일명: 배치도1.jpg, 존재하면 배치도2.jpg, ...
    i = 1
    Do While fso.FileExists(defaultFolder & "\배치도" & i & ".jpg")
        i = i + 1
    Loop
    defaultFile = "배치도" & i & ".jpg"
    
    ' 파일 저장 대화상자를 통해 저장 경로와 파일명 선택 (기본값 제공)
    filePath = Application.GetSaveAsFilename( _
                    InitialFileName:=defaultFolder & "\" & defaultFile, _
                    FileFilter:="JPEG Files (*.jpg), *.jpg", _
                    Title:="저장할 파일 경로와 파일명을 선택하세요")
    
    ' 사용자가 저장을 취소한 경우
    If filePath = False Then
        MsgBox "저장이 취소되었습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 이름이 "save_range"인 영역을 가져옴
    On Error Resume Next
    Set rng = Range("save_range")
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "이름이 'save_range'인 영역을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 영역을 그림으로 복사 (화면상의 모습으로)
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    DoEvents  ' 클립보드 준비를 위해 잠시 대기
    
    ' 임시 시트를 추가하고, 그 위에 임시 차트를 생성하여 영역 크기에 맞춤
    Set tempSheet = ActiveSheet
    Set tempChartObj = tempSheet.ChartObjects.Add(Left:=0, Top:=0, Width:=rng.Width, Height:=rng.Height)
    tempChartObj.Border.LineStyle = xlNone
    
    With tempChartObj.Chart
        ' 기존 기본 차트 요소(예: 기본 시리즈)가 있다면 삭제하여 깨끗한 상태로 만듦
        On Error Resume Next
        Do Until .SeriesCollection.Count = 0
            .SeriesCollection(1).Delete
        Loop
        On Error GoTo 0
        
        ' 임시 차트를 활성화한 후 ActiveChart를 통해 붙여넣기
        tempChartObj.Activate
        ActiveChart.Paste
        
        ' 차트 개체의 크기를 영역과 동일하게 설정
        .Parent.Width = rng.Width
        .Parent.Height = rng.Height
        
        ' 그림 파일(JPG)로 내보내기
        .Export fileName:=filePath, FilterName:="jpg"
    End With
    
    ' 임시 시트 삭제
    Application.DisplayAlerts = False
    tempChartObj.Delete
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "저장되었습니다: " & filePath, vbInformation
End Sub


'──────────────────────────────────────────────────────────
' saveLayout 실행 전에 이 함수를 먼저 호출하면
' save_range 내부의 도형들이 중앙 정렬된 상태로 배치됩니다.
'──────────────────────────────────────────────────────────
Public Sub CenterAlignSaveRange()
    Dim ws             As Worksheet
    Dim rng            As Range
    Dim sh             As Shape
    Dim minLeft        As Double, maxRight As Double
    Dim minTop         As Double, maxBottom As Double
    Dim centerRangeX   As Double, centerRangeY As Double
    Dim centerShapesX  As Double, centerShapesY As Double
    Dim offsetX        As Double, offsetY As Double

    Set ws = ActiveSheet
    On Error Resume Next
        Set rng = ws.Range("save_range")
    On Error GoTo 0
    If rng Is Nothing Then
        Exit Sub
    End If

    ' 초기값 세팅
    minLeft = 1E+99: maxRight = 0
    minTop = 1E+99: maxBottom = 0

    ' 1) 경계값 계산 (LineBasic 제외)
    For Each sh In ws.Shapes
        If sh.Name <> "LineBasic" Then
            If Not (sh.Left + sh.Width < rng.Left Or _
                    sh.Left > rng.Left + rng.Width Or _
                    sh.Top + sh.Height < rng.Top Or _
                    sh.Top > rng.Top + rng.Height) Then

                minLeft = Application.Min(minLeft, sh.Left)
                maxRight = Application.Max(maxRight, sh.Left + sh.Width)
                minTop = Application.Min(minTop, sh.Top)
                maxBottom = Application.Max(maxBottom, sh.Top + sh.Height)
            End If
        End If
    Next sh

    ' 정렬 대상 도형이 없는 경우
    If maxRight = 0 Then
        Exit Sub
    End If

    ' 2) 중심 좌표 계산
    centerRangeX = rng.Left + rng.Width / 2
    centerRangeY = rng.Top + rng.Height / 2
    centerShapesX = (minLeft + maxRight) / 2
    centerShapesY = (minTop + maxBottom) / 2

    offsetX = centerRangeX - centerShapesX
    offsetY = centerRangeY - centerShapesY

    ' 3) 도형 이동
    For Each sh In ws.Shapes
        If sh.Name <> "LineBasic" Then
            If Not (sh.Left + sh.Width < rng.Left Or _
                    sh.Left > rng.Left + rng.Width Or _
                    sh.Top + sh.Height < rng.Top Or _
                    sh.Top > rng.Top + rng.Height) Then

                sh.Left = sh.Left + offsetX
                sh.Top = sh.Top + offsetY
            End If
        End If
    Next sh

End Sub


Sub LayOut_Helper(selectedSheetName As String)
    Dim wsSelected    As Worksheet
    Dim rngTotal      As Range, rngLayout As Range
    Dim shapeCount    As Integer
    Dim layoutLeft    As Integer, layoutTop As Integer, layoutWidth As Integer
    Dim shapesPerRow  As Integer
    Dim i             As Integer, rowIdx As Integer, colIdx As Integer
    Dim currentX      As Integer, currentY As Integer
    Dim w             As Integer, h As Integer, gap As Integer
    Dim equipName     As String
    Dim dataRow       As Integer

    On Error Resume Next
        Set wsSelected = ThisWorkbook.Worksheets(selectedSheetName)
        Set rngTotal = wsSelected.Range("Total_Config")
        Set rngLayout = ActiveSheet.Range("Layout_board")
    On Error GoTo 0

    If wsSelected Is Nothing Or rngTotal Is Nothing Or rngLayout Is Nothing Then
        MsgBox "영역을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' Layout_board 위치/크기 (Integer)
    layoutLeft = CInt(rngLayout.Left)
    layoutTop = CInt(rngLayout.Top)
    layoutWidth = CInt(rngLayout.Width)

    ' 도형 크기/간격
    w = 60
    h = 80
    gap = 10

    ' 그릴 도형 수 (제목행 제외)
    shapeCount = Application.WorksheetFunction.CountA(rngTotal.Columns(1)) - 1
    If shapeCount < 1 Then Exit Sub

    ' 한 행에 들어갈 최대 도형 수
    shapesPerRow = (layoutWidth + gap) \ (w + gap)
    If shapesPerRow < 1 Then shapesPerRow = 1

    ' 도형 그리기
    For i = 0 To shapeCount - 1
        rowIdx = i \ shapesPerRow   ' 0부터 시작하는 행 인덱스
        colIdx = i Mod shapesPerRow  ' 0부터 시작하는 열 인덱스

        currentX = layoutLeft + colIdx * (w + gap)
        currentY = layoutTop + rowIdx * 30    ' 매 행마다 10씩 아래로

        ' 이름 결정
        If i = 0 Then
            equipName = "피시험기자재"
        Else
            dataRow = i + 2     ' 1행(제목) + i번째 데이터
            equipName = rngTotal.cells(dataRow, 1).value
        End If

        ' EquipmentDraw는 모두 Integer 파라미터여야 합니다
        EquipmentDraw equipName, currentX, currentY, w, h
    Next i
End Sub


Sub EquipmentDraw(ByVal EqName As String, x As Integer, y As Integer, w As Integer, h As Integer)

    Application.EnableEvents = False
    
    Dim shtx As Worksheet
    Dim shpx As Shape
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    Set shpx = shtx.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    With shpx
    
        .Fill.ForeColor.SchemeColor = 1
        .Line.ForeColor.SchemeColor = 0
        With .TextFrame.Characters
            .Text = EqName
            With .Font
                .Name = "맑은 고딕"
                .Size = 10
                .ColorIndex = 1
                .Bold = False
            End With
        End With
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    
    End With
    
    Range("eqName").value = ""
    Application.EnableEvents = True
    
End Sub

Sub TextboxDraw(ByVal txtName As String, x As Integer, y As Integer)

    Application.EnableEvents = False
    
    Dim shtx As Worksheet
    Dim shpx As Shape
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    Set shpx = shtx.Shapes.AddLabel(msoTextOrientationHorizontal, x, y, 80, 20)
    With shpx
        .TextFrame.Characters.Text = txtName
        .TextFrame.Characters.Font.Size = 10
    End With
    
    Range("txtName") = ""
    
    Application.EnableEvents = True
    
End Sub


Sub PasteYLines(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y

    Set shpRange = ActiveSheet.Shapes.Range(Array("MainsLine")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
        .Left = x: .Top = y: .Width = w: .Height = h
    End With
    
End Sub

Sub PasteCLines(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y

    Set shpRange = ActiveSheet.Shapes.Range(Array("ConnectionLine")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
        .Left = x: .Top = y: .Width = w: .Height = h
    End With
    
End Sub

Sub PasteILines(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y

    Set shpRange = ActiveSheet.Shapes.Range(Array("ConnectionRight")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
        .Left = x: .Top = y: .Width = w: .Height = h
    End With
    
End Sub

Sub PasteUSB(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    
    Set shpRange = ActiveSheet.Shapes.Range(Array("USBp")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
         .Left = x: .Top = y: .Width = w: .Height = h
         .ZOrder msoSendToBack
         .Flip msoFlipHorizontal
    End With
    
End Sub

Sub PasteKeyboard(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    
    Set shpRange = ActiveSheet.Shapes.Range(Array("Keyboard")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
         .Left = x: .Top = y: .Width = w: .Height = h
    End With
    
End Sub

Sub PasteMouse(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    
    Set shpRange = ActiveSheet.Shapes.Range(Array("Mouse")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
         .Left = x: .Top = y: .Width = w: .Height = h
    End With
    
End Sub

Sub PasteHedset(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    
    Set shpRange = ActiveSheet.Shapes.Range(Array("Headset")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
         .Left = x: .Top = y: .Width = w: .Height = h
    End With
    
End Sub


Sub PasteWireless(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    
    Set shpRange = ActiveSheet.Shapes.Range(Array("Wireless")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
         .Left = x: .Top = y: .Width = w: .Height = h
    End With
    
End Sub

Sub PasteFrameGround(x As Integer, y As Integer, w As Integer, h As Integer)

    Dim newShape As Shape
    Dim shpRange As ShapeRange
    Dim shtx As Worksheet
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    
    Set shpRange = ActiveSheet.Shapes.Range(Array("FrameGround")).Duplicate
    Set newShape = shpRange(1)
    
    With newShape
        .Name = .Name & "temp"
         .Left = x: .Top = y: .Width = w: .Height = h
    End With
    
End Sub

Sub CreateAndGroupLines(ParamArray coords() As Variant)
    
    Dim countCoords As Long, numPoints As Long, i As Long
    countCoords = UBound(coords) - LBound(coords) + 1
    
    Dim shapeCount As Long

    If countCoords Mod 2 <> 0 Then
        Exit Sub
    End If
    
    numPoints = countCoords / 2
    If numPoints < 2 Then
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim layRG As Range
    
    Set layRG = ws.Range("Layout_board")
    x = layRG.Left + x
    y = layRG.Top + y
    
    Dim arrNames() As String
    ReDim arrNames(1 To numPoints - 1)
    shapeCount = UBound(arrNames) - LBound(arrNames) + 1
    For i = 1 To numPoints - 1
        Dim startX As Variant, startY As Variant, endX As Variant, endY As Variant
        startX = coords((i - 1) * 2) + x
        startY = coords((i - 1) * 2 + 1) + y
        endX = coords(i * 2) + x
        endY = coords(i * 2 + 1) + y
        
        Dim newLine As Shape
        Set newLine = ws.Shapes.AddLine(startX, startY, endX, endY)
        With newLine.Line
            .Weight = 1
            .ForeColor.RGB = RGB(0, 0, 0)
        End With
        arrNames(i) = newLine.Name
    Next i
    
    If shapeCount > 1 Then
        Dim grp As Shape
        Set grp = ws.Shapes.Range(arrNames).Group
        grp.Name = "GroupedLines"
    End If
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' button click
Sub clicked_mainsLine()

    On Error GoTo ErrHandler
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top
    
        If Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
            With Selection
                Call PasteYLines(.Left - x, .Top - 40 - y, 30, 40)
            End With
        End If
    Exit Sub

ErrHandler:
    Exit Sub
    
End Sub

Sub clicked_connectionLine()
    
    Dim l1 As Integer, l2 As Integer
    On Error GoTo ErrHandler
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top
    
    Set shpRange = Selection.ShapeRange
    If shpRange.Count = 2 Then
        Set shpRange = Selection.ShapeRange
            
        l1 = shpRange(1).Left + shpRange(1).Width - 10
        l2 = shpRange(2).Left - l1
        
        Call PasteCLines(l1 - x, shpRange(1).Top - 30 - y, l2 + 10, 30)

    ElseIf Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
        With Selection
            Call PasteCLines(.Left + 60 - x, .Top - 30 - y, 60, 30)
        End With
            End If
    Exit Sub

ErrHandler:
    Exit Sub

End Sub

Sub clicked_connectionRight()
    
    Dim l1 As Integer, l2 As Integer
    On Error GoTo ErrHandler
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top
    
    Set shpRange = Selection.ShapeRange
    If shpRange.Count = 2 Then
        Set shpRange = Selection.ShapeRange
            
        l1 = shpRange(1).Left + shpRange(1).Width
        l2 = shpRange(2).Left - l1
        
        Call PasteILines(l1 - x, shpRange(1).Top + 30 - y, l2, 0)

    ElseIf Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
        With Selection
            Call PasteILines(.Left + .Width - x, .Top + 30 - y, 40, 0)
        End With
            End If
    Exit Sub

ErrHandler:
    Exit Sub

End Sub

Sub clicked_USBp()
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top
    
    On Error GoTo ErrHandler
        If Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
            With Selection
                Call PasteUSB(.Left + .Width - 3 - x, .Top + 20 - y, 10, 15)
            End With
        End If
    Exit Sub

ErrHandler:
    Exit Sub

End Sub

Sub clicked_Keyboard()
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top

    On Error GoTo ErrHandler
        If Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
            With Selection
                Call PasteKeyboard(.Left + 10 - x, .Top + .Height + 20 - y, 40, 20)
                Call CreateAndGroupLines(.Left + 30 - x, .Top + .Height - y, .Left + 30 - x, .Top + .Height + 22 - y)
            End With
        End If
    Exit Sub

ErrHandler:
    Exit Sub

End Sub

Sub clicked_Mouse()

    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top
    
    On Error GoTo ErrHandler
        If Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
            With Selection
                Call PasteMouse(.Left + .Width - 20 - x, .Top + .Height + 20 - y, 20, 20)
                Call CreateAndGroupLines(.Left + .Width - 11 - x, .Top + .Height - y, .Left + .Width - 11 - x, .Top + .Height + 21 - y)
            End With
        End If
    Exit Sub

ErrHandler:
    Exit Sub

End Sub

Sub clicked_Headset()

    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top
    
    On Error GoTo ErrHandler
        If Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
            With Selection
                Call PasteHedset(.Left + .Width - 20 - x, .Top + .Height + 20 - y, 20, 20)
                Call CreateAndGroupLines(.Left + .Width - 11 - x, .Top + .Height - y, .Left + .Width - 11 - x, .Top + .Height + 21 - y)
            End With
        End If
    Exit Sub

ErrHandler:
    Exit Sub

End Sub

Sub clicked_Wireless()
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top

    On Error GoTo ErrHandler
        If Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
            With Selection
                Call PasteWireless(.Left + .Width + 20 - x, .Top + .Height / 2 - 10 - y, 20, 20)
            End With
        End If
    Exit Sub

ErrHandler:
    Exit Sub

End Sub


Sub clicked_FrameGround()
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top
    On Error GoTo ErrHandler
        If Selection.Top > Range("Layout_board").Top And Selection.Top < Range("Layout_board").Height + Range("Layout_board").Top Then
            With Selection
                Call PasteFrameGround(.Left + .Width - 20 - x, .Top + .Height - y, 20, 20)
            End With
        End If
    Exit Sub

ErrHandler:
    Exit Sub

End Sub

Sub clicked_ModeDevision()
    Dim layRG As Range
    
    Set shtx = ActiveSheet
    Set layRG = shtx.Range("Layout_board")
    x = layRG.Left
    y = layRG.Top
    midX = x + layRG.Width / 2
    
    Set shp = shtx.Shapes.AddLine( _
        BeginX:=midX, BeginY:=y, _
        endX:=midX, endY:=y + layRG.Height _
    )
    
    With shp.Line
        .DashStyle = msoLineDash      ' 대시 타입
        .Weight = 1                   ' 선 두께
        .ForeColor.RGB = RGB(0, 0, 0) ' 검정색
    End With
    
    

ErrHandler:
    Exit Sub
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' clear 함수
Sub LayOut_Clear()
    
    Dim shtx As Worksheet
    Dim rng As String
    
    Set shtx = ActiveSheet
    rng = "C5:J24"
        Call Delete_Picture(shtx, rng)
    rng = "K5:L20"
        Call Delete_Picture(shtx, rng)
    
End Sub

Sub Delete_Picture(ByVal shtx As Worksheet, rg As String)
   
    Dim shpC As Shape, rngShp As Range, rngAll As Range
    
    On Error Resume Next
    Application.ScreenUpdating = False
    
    shtx.Activate
    Set rngAll = Range(rg)
   
    For Each shpC In shtx.Shapes
        Set rngShp = shpC.TopLeftCell
        If Not Intersect(rngAll, rngShp) Is Nothing Then
            shpC.Delete
        End If
    Next shpC
    
    Set shtx = Nothing
    Set rngAll = Nothing
    Set rngShp = Nothing
    
    Application.ScreenUpdating = True
    
    DoEvents
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''preset 호출

Sub ExecutePresetCommands(preset As String)
    Const API_URL As String = _
      "https://script.google.com/macros/s/AKfycbzXJTQcdnj3dP-HrgPC69SVxyXWuReIM06PEY9wuaBuOprjYIU0ISSIHWvjcrW_1IgM/exec"
    Const API_KEY As String = "dtncalfdnjzl!453"
    
    Dim http As Object, json As String
    Dim rows As Variant, rawRow As String
    Dim i As Long
    Dim url As String
    
    ' 1) 전체 Preset_Param 시트 가져오기
    url = API_URL & "?key=" & API_KEY & "&sheet=Preset_Param"
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    If http.Status <> 200 Then
        MsgBox "명령 가져오기 실패: " & http.Status & " " & http.statusText, vbExclamation
        Exit Sub
    End If
    json = http.responseText
    
    ' 2) JSON 간이 파싱
    json = Mid$(json, 2, Len(json) - 2)
    rows = Split(json, "],[")   ' 첫 행은 헤더
    
    ' 3) 데이터 행 순회
    For i = 1 To UBound(rows)
        ' 3-1) 대괄호·쌍따옴표 제거
        rawRow = Replace(Replace(Replace(rows(i), "[", ""), "]", ""), """", "")
        
        ' 3-2) 맨 앞 쉼표 위치로 Preset만 추출
        Dim pos1 As Long
        pos1 = InStr(rawRow, ",")
        Dim rowPreset As String
        rowPreset = Left(rawRow, pos1 - 1)
        
        ' 3-3) preset이 일치할 때만 실행
        If rowPreset = preset Then
            ' 3-4) 두 번째·세 번째 쉼표 위치 찾아 함수명·Param1 분리
            Dim pos2 As Long, pos3 As Long
            pos2 = InStr(pos1 + 1, rawRow, ",")
            pos3 = InStr(pos2 + 1, rawRow, ",")
            
            Dim funcName As String
            funcName = Mid(rawRow, pos2 + 1, pos3 - pos2 - 1)
            
            Dim p1 As String
            p1 = Mid(rawRow, pos3 + 1)  ' 3번째 쉼표 뒤 전체
            
            ' 3-5) 호출
            'Debug.Print "Calling "; funcName; " with p1="; p1
            Application.Run funcName, p1
        End If
    Next i
End Sub


Sub ShowSheetSelector()

    List.Show
    
End Sub
