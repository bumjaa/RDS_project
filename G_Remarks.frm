VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} G_Remarks 
   Caption         =   "비고"
   ClientHeight    =   8310.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17175
   OleObjectBlob   =   "G_Remarks.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "G_Remarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aOptionButtons() As New optClass

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    Dim uniqueItems As Collection
    Dim prefixDict As Object
    Dim optionButton As Object
    Dim key As Variant
    Dim optionIndex As Integer
    Dim iCounter As Integer
    Dim gRemarksContent As String
    Dim splitValues As Variant
    Dim i As Integer

    ' Worksheet 및 데이터 범위 설정
    Set ws = ActiveSheet ' 적절한 시트 이름으로 변경하세요
    Set dataRange = ws.Range("G_Remarks_contents")
    
    If Application.WorksheetFunction.CountA(dataRange) = 0 Then
        Exit Sub
    End If
    
    ' 중복 제거를 위한 Collection 및 Prefix 저장용 Dictionary 초기화
    Set uniqueItems = New Collection
    Set prefixDict = CreateObject("Scripting.Dictionary")

    ' G_Remarks_contents 영역에서 값을 추출하고 중복 제거
    On Error Resume Next ' Collection에 중복값 추가 시 오류 방지
    For Each cell In dataRange
        Dim splitValuesTemp As Variant
        splitValuesTemp = Split(cell.value, ",")
        Dim j As Integer
        For j = LBound(splitValuesTemp) To UBound(splitValuesTemp)
            uniqueItems.Add Trim(splitValuesTemp(j)), CStr(Trim(splitValuesTemp(j)))
        Next j
    Next cell
    On Error GoTo 0

    ' Prefix 값 매핑 (셀 이름에서 시트 이름과 "_Remarks" 제거)
    For Each cell In dataRange
        Dim cellName As String
        Dim processedName As String

        ' 셀의 이름 관리자 이름 가져오기
        On Error Resume Next
        cellName = cell.Name.Name
        On Error GoTo 0

        If cellName = "" Then
            processedName = "Unknown" ' 이름이 없을 경우 기본값
        Else
            ' 시트 이름 및 "_Remarks" 제거
            processedName = Replace(cellName, "'" & ws.Name & "'!", "")
            processedName = Replace(processedName, "_Remarks", "")
        End If

        Dim splitRemarks As Variant
        splitRemarks = Split(cell.value, ",")
        Dim k As Integer
        For k = LBound(splitRemarks) To UBound(splitRemarks)
            Dim remark As String
            remark = Trim(splitRemarks(k))
            If Not prefixDict.Exists(remark) Then
                prefixDict.Add remark, processedName
            Else
                prefixDict(remark) = prefixDict(remark) & ", " & processedName
            End If
        Next k
    Next cell

    ' UserForm에 OptionButton 생성
    optionIndex = 1
    ReDim aOptionButtons(1 To uniqueItems.Count)
    iCounter = 0

    For Each key In uniqueItems
        Set optionButton = Me.Controls.Add("Forms.OptionButton.1", "OptionButton" & optionIndex)
        optionButton.Caption = key & " : " & prefixDict(key)
        optionButton.Top = 20 + optionIndex * 20 ' 각 옵션 단추의 간격
        optionButton.Left = 720 ' 기준 Left 위치
        optionButton.Width = 200

        ' Class1 연결
        iCounter = iCounter + 1
        Set aOptionButtons(iCounter) = New optClass
        Set aOptionButtons(iCounter).ctlOptionButton = optionButton

        optionIndex = optionIndex + 1
    Next key

End Sub

Public Sub optSearch()
    Dim selectedOption As String
    Dim prefixArray  As Variant
    Dim prefix       As Variant
    Dim SSID         As String, gid_GRem As String
    Dim lan          As String
    Dim q            As String, url As String
    Dim http         As Object, respRaw As String, respJson As String
    Dim p1           As Long, p2 As Long
    Dim parsed       As Object, rows As Object, rowObj As Object
    Dim cells        As Object
    Dim i            As Long
    Dim uniqueDict   As Object
    
    ' 1) 선택된 OptionButton 확인
    If TypeName(ActiveControl) <> "OptionButton" Then
        MsgBox "옵션 단추를 선택해주세요.", vbExclamation
        Exit Sub
    End If
    selectedOption = ActiveControl.Caption
    
    ' 2) Prefix 분리
    prefixArray = Split(Mid(selectedOption, InStr(selectedOption, ":") + 1), ",")
    
    ' 3) 언어 결정
    If Left(ActiveSheet.Range("STD").value, 2) = "KS" Or Left(ActiveSheet.Range("STD").value, 2) = "KN" Then
        lan = "KO"
    Else
        lan = "EN"
    End If
    
    stds = Replace(ActiveSheet.Range("STD").value, ",", "")
    
    ' 4) Google Sheets 정보
    SSID = "16ohT-jZlrzg9awKMRgztCtAL1POuRua-JfDlkG6YHKU"      ' 본인 ID
    gid_GRem = "1165139260" ' ← G_Remarks 시트의 gid
    
    Set uniqueDict = CreateObject("Scripting.Dictionary")
    
    For Each prefix In prefixArray
        prefix = Trim(prefix)
        
        ' 5) 쿼리 문자열: A열=ITEM, B열=Remarks_<lan>
        If lan = "KO" Then
            q = "select C where B = '" & Replace(prefix, "'", "\'") & "' and (A = 'COMMON' or A = '" & Replace(stds, " '", "\'") & "')"
        Else
            q = "select D where B = '" & Replace(prefix, "'", "\'") & "' and (A = 'COMMON' or A = '" & Replace(stds, " '", "\'") & "')"
        End If
        
        url = "https://docs.google.com/spreadsheets/d/" & SSID & "/gviz/tq" & _
              "?gid=" & gid_GRem & _
              "&tqx=out:json" & _
              "&tq=" & URLEncode(q)
        
        ' 6) HTTP GET
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.Open "GET", url, False
        http.send
        If http.Status <> 200 Then
            MsgBox "G_Remarks 조회 실패: " & http.Status, vbExclamation
            Exit Sub
        End If
        respRaw = http.responseText
        
        ' 7) 래퍼 제거
        p1 = InStr(respRaw, "{"): p2 = InStrRev(respRaw, "}")
        If p1 = 0 Or p2 < p1 Then
            MsgBox "JSON 추출 실패", vbExclamation
            Exit Sub
        End If
        respJson = Mid$(respRaw, p1, p2 - p1 + 1)
        
        ' 8) JSON 파싱
        Set parsed = JsonConverter.ParseJson(respJson)
        Set rows = parsed("table")("rows")   ' 1-based
        
        ' 9) 결과 수집
        For i = 1 To rows.Count
            Set rowObj = rows(i)
            Set cells = rowObj.item("c")       ' 1-based Collection
            ' cells(1)("v") 이면 첫 컬럼, 여기선 B열(=첫 컬럼)만 select 했으므로 cells(1)
            If Not IsNull(cells(1).item("v")) Then
                uniqueDict(cells(1).item("v")) = True
            End If
        Next i
    Next prefix
    
    ' 10) ListBox1 채우기
    Me.ListBox1.Clear
    Dim key As Variant
    For Each key In uniqueDict.keys
        Me.ListBox1.AddItem key
    Next key
    
End Sub



Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedItem As String
    Dim selectedOption As String
    Dim prefix As String
    Dim fullContent As String

    ' ListBox1에서 선택된 항목 확인
    If ListBox1.ListIndex <> -1 Then
        selectedItem = ListBox1.List(ListBox1.ListIndex)

        ' 현재 선택된 OptionButton의 Caption 확인
        selectedOption = GetSelectedOptionButtonCaption()
        If selectedOption = "" Then
            MsgBox "현재 선택된 옵션 버튼이 없습니다.", vbExclamation
            Exit Sub
        End If

        ' OptionButton의 주제와 선택된 ListBox 항목 결합
        prefix = Split(selectedOption, ":")(0) ' 주제 부분 추출 (예: "주1)")
        fullContent = prefix & selectedItem

        ' ListBox2에 항목 추가 (중복 확인)
        If Not IsItemInListBox(ListBox2, fullContent) Then
            ListBox2.AddItem fullContent
        End If
    Else
        MsgBox "항목을 선택해주세요.", vbExclamation
    End If
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' ListBox2에서 선택된 항목 삭제
    If ListBox2.ListIndex <> -1 Then
        ListBox2.RemoveItem ListBox2.ListIndex
    Else
        MsgBox "항목을 선택해주세요.", vbExclamation
    End If
End Sub



Private Function GetSelectedOptionButtonCaption() As String
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" And ctrl.value = True Then
            GetSelectedOptionButtonCaption = ctrl.Caption
            Exit Function
        End If
    Next ctrl
    GetSelectedOptionButtonCaption = ""
End Function

' ListBox 중복 확인 함수
Private Function IsItemInListBox(lst As MSForms.listBox, item As String) As Boolean
    Dim i As Integer
    For i = 0 To lst.listCount - 1
        If lst.List(i) = item Then
            IsItemInListBox = True
            Exit Function
        End If
    Next i
    IsItemInListBox = False
End Function


Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim listContent As String
    Dim i As Long
    Dim gRms As Range

    ' ListBox2 -> 줄바꿈으로 연결 (마지막에 불필요한 vbLf 안 붙이기)
    For i = 0 To Me.ListBox2.listCount - 1
        If i > 0 Then listContent = listContent & vbLf
        listContent = listContent & Me.ListBox2.List(i)
    Next i

    ' 앞/뒤의 줄바꿈만 제거(중간 줄바꿈은 보존)
    listContent = TrimOuterLineBreaks(listContent)

    ' G_Remarks 범위 얻기(정의 안돼 있으면 안내)
    On Error Resume Next
    Set ws = ActiveSheet
    Set gRms = ws.Range("G_Remarks")
    On Error GoTo 0
    If gRms Is Nothing Then
        MsgBox "G_Remarks 영역이 정의되지 않았습니다.", vbExclamation
        Exit Sub
    End If

    If listContent <> "" Then
        ' 기존 로직 유지: 기존 값이 있으면 줄바꿈 후 이어붙이기
        If Len(gRms.Value2) > 0 Then
            gRms.value = gRms.value & vbLf & listContent
        Else
            gRms.value = listContent
        End If
        gRms.WrapText = True
    Else
        MsgBox "추가할 항목이 없습니다.", vbInformation
    End If

    Unload Me
End Sub

' 앞/뒤의 CR/LF만 제거(중간은 건드리지 않음)
Private Function TrimOuterLineBreaks(ByVal s As String) As String
    ' 앞쪽 제거
    Do While Len(s) > 0 And (Left$(s, 1) = vbCr Or Left$(s, 1) = vbLf)
        s = Mid$(s, 2)
    Loop
    ' 뒤쪽 제거
    Do While Len(s) > 0 And (Right$(s, 1) = vbCr Or Right$(s, 1) = vbLf)
        s = Left$(s, Len(s) - 1)
    Loop
    TrimOuterLineBreaks = s
End Function


Private Sub ListBox1_Click()
    Dim i As Integer
    ' ListBox2의 선택 상태를 초기화
    For i = 0 To Me.ListBox2.listCount - 1
        Me.ListBox2.Selected(i) = False
    Next i
End Sub

' ListBox2에서 선택했을 때 ListBox1의 선택을 취소
Private Sub ListBox2_Click()
    Dim i As Integer
    ' ListBox1의 선택 상태를 초기화
    For i = 0 To Me.ListBox1.listCount - 1
        Me.ListBox1.Selected(i) = False
    Next i
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then ' ESC 키
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub ListBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then ' ESC 키
        Unload Me
        Exit Sub
    End If
End Sub
