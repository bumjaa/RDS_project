VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} T_Comments 
   Caption         =   "시험자 의견"
   ClientHeight    =   8310.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17160
   OleObjectBlob   =   "T_Comments.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "T_Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public currentPrefix As String ' prefix 저장 속성
Private optHandlers As Collection ' 라디오 버튼 핸들러 컬렉션

Private Const SSID_COMMENTS As String = "16ohT-jZlrzg9awKMRgztCtAL1POuRua-JfDlkG6YHKU"  ' 여러분 스프레드시트 ID
Private Const GID_COMMENTS As String = "1187068989"  ' Comments 시트의 gid (URL #gid=xxxx)

Public Sub LoadData(ByVal prefix As String)
    Dim lan      As String
    Dim q        As String, url As String
    Dim http     As Object, raw As String, jsn As String
    Dim parsed   As Object, rows As Object, rowObj As Object
    Dim cells    As Object
    Dim dictMode As Object
    Dim i As Long, topPos As Double
    Dim modeKey  As Variant
    
    currentPrefix = prefix
    
    ' 1) 언어 결정
    If Left(ActiveSheet.Range("STD").value, 2) = "KS" Or Left(ActiveSheet.Range("STD").value, 2) = "KN" Then
        lan = "KO"
    Else
        lan = "EN"
    End If
    
    ' 2) Clear 기존 UI
    Me.ListBox1.Clear
    ' 동적 OptionButtons 초기화
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" Then Me.Controls.Remove ctrl.Name
    Next
    Set dictMode = CreateObject("Scripting.Dictionary")
    
    ' 3) Remarks_<lan> 가져오기
    If lan = "KO" Then
        q = "select D where B = '" & Replace(prefix, "'", "\'") & "' and C <> ''"
    Else
        q = "select E where B = '" & Replace(prefix, "'", "\'") & "' and C <> ''"
    End If
    url = "https://docs.google.com/spreadsheets/d/" & SSID_COMMENTS & "/gviz/tq" & _
          "?gid=" & GID_COMMENTS & _
          "&tqx=out:json" & _
          "&tq=" & URLEncode(q)
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False: http.send
    If http.Status = 200 Then
        raw = http.responseText
        ' 래퍼 제거
        Dim p1 As Long, p2 As Long
        p1 = InStr(raw, "{"): p2 = InStrRev(raw, "}")
        If p1 > 0 And p2 > p1 Then jsn = Mid$(raw, p1, p2 - p1 + 1)
        Set parsed = JsonConverter.ParseJson(jsn)
        Set rows = parsed("table")("rows")
        ' ListBox1 채우기
        For i = 1 To rows.Count
            Set rowObj = rows(i)
            Set cells = rowObj.item("c")
            Me.ListBox1.AddItem cells(1).item("v")
        Next
    Else
        MsgBox "Remarks 조회 실패: " & http.Status, vbExclamation
    End If
    
    ' 4) MODE 가져오기 (Dictionary 로 중복 제거)
    q = "select C where B = '" & Replace(prefix, "'", "\'") & "' and C <> ''"
    url = "https://docs.google.com/spreadsheets/d/" & SSID_COMMENTS & "/gviz/tq" & _
          "?gid=" & GID_COMMENTS & _
          "&tqx=out:json" & _
          "&tq=" & URLEncode(q)
    http.Open "GET", url, False: http.send
    If http.Status = 200 Then
        raw = http.responseText
        p1 = InStr(raw, "{"): p2 = InStrRev(raw, "}")
        If p1 > 0 And p2 > p1 Then jsn = Mid$(raw, p1, p2 - p1 + 1)
        Set parsed = JsonConverter.ParseJson(jsn)
        Set rows = parsed("table")("rows")
        For i = 1 To rows.Count
            Set rowObj = rows(i)
            Set cells = rowObj.item("c")
            dictMode(cells(1).item("v")) = True
        Next
    Else
        MsgBox "MODE 조회 실패: " & http.Status, vbExclamation
    End If
    
    ' 5) 동적 OptionButton 생성
    topPos = 20
    Dim idx As Long: idx = 1
    Dim hb As New Collection
    Dim optBtn As MSForms.optionButton
    For Each modeKey In dictMode.keys
        Set optBtn = Me.Controls.Add("Forms.OptionButton.1", "optMode" & idx)
        optBtn.Caption = modeKey
        optBtn.Left = 700
        optBtn.Top = topPos
        topPos = topPos + 20
        
        ' 클래스 핸들러 연결 (clsOptionButtonHandler)
        Dim h As clsOptionButtonHandler
        Set h = New clsOptionButtonHandler
        Set h.optButton = optBtn
        hb.Add h
        
        idx = idx + 1
    Next modeKey
    
    Set optHandlers = hb
End Sub


Public Sub OptionButton_Click()
    Dim selectedMode As String
    Dim lan  As String
    Dim q    As String, url As String
    Dim http As Object, raw As String, jsn As String
    Dim parsed As Object, rows As Object, rowObj As Object
    Dim cells  As Object
    Dim i      As Long
    
    ' 1) 선택된 MODE 추출
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.value = True Then selectedMode = ctrl.Caption: Exit For
        End If
    Next
    If selectedMode = "" Then Exit Sub
    
    ' 2) 언어 결정
    If Left(ActiveSheet.Range("STD").value, 2) = "KS" Then lan = "KO" Else lan = "EN"
    
    ' 3) ListBox1 클리어
    Me.ListBox1.Clear
    
    ' 4) server 에서 필터링된 Remarks_<lan>만 가져오기
    q = "select D where B = '" & Replace(currentPrefix, "'", "\'") & "' " & _
        "and C = '" & Replace(selectedMode, "'", "\'") & "'"
    url = "https://docs.google.com/spreadsheets/d/" & SSID_COMMENTS & "/gviz/tq" & _
          "?gid=" & GID_COMMENTS & _
          "&tqx=out:json" & _
          "&tq=" & URLEncode(q)
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False: http.send
    If http.Status = 200 Then
        raw = http.responseText
        Dim p1 As Long, p2 As Long
        p1 = InStr(raw, "{"): p2 = InStrRev(raw, "}")
        If p1 > 0 And p2 > p1 Then jsn = Mid$(raw, p1, p2 - p1 + 1)
        Set parsed = JsonConverter.ParseJson(jsn)
        Set rows = parsed("table")("rows")
        For i = 1 To rows.Count
            Set rowObj = rows(i)
            Set cells = rowObj.item("c")
            Me.ListBox1.AddItem cells(1).item("v")
        Next
    Else
        MsgBox "조회 실패: " & http.Status, vbExclamation
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If ListBox1.ListIndex <> -1 Then ' 선택된 항목이 있을 경우
        Dim selectedItem As String
        selectedItem = ListBox1.value
        ' 중복 확인 후 추가
        If Not IsInList(ListBox2, selectedItem) Then
            ListBox2.AddItem selectedItem
        End If
    End If
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If ListBox2.ListIndex <> -1 Then ' 선택된 항목이 있을 경우
        ListBox2.RemoveItem ListBox2.ListIndex
    End If
End Sub

Private Sub ListBox1_Change()
    If ListBox1.ListIndex <> -1 Then
        ListBox2.ListIndex = -1 ' ListBox2 선택 취소
    End If
End Sub

Private Sub ListBox2_Change()
    If ListBox2.ListIndex <> -1 Then
        ListBox1.ListIndex = -1 ' ListBox1 선택 취소
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim combinedText As String
    Dim i As Integer

    ' ListBox2의 값 가져오기 (줄바꿈 처리)
    For i = 0 To ListBox2.listCount - 1
        If combinedText <> "" Then
            combinedText = combinedText & vbLf ' 각 항목 사이에 줄바꿈 추가
        End If
        combinedText = combinedText & ListBox2.List(i)
    Next i

    ' 현재 선택된 셀에 값 추가
    If combinedText <> "" Then
        If Not ActiveCell Is Nothing Then
            If ActiveCell.value <> "" Then
                ActiveCell.value = combinedText
            Else
                ActiveCell.value = combinedText
            End If
        Else
            MsgBox "셀을 선택해주세요.", vbExclamation
        End If
    Else
        MsgBox "추가할 항목이 없습니다.", vbInformation
    End If

    ' UserForm 닫기
    Unload Me
End Sub


Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Function IsInList(lst As MSForms.listBox, value As String) As Boolean
    Dim i As Integer
    For i = 0 To lst.listCount - 1
        If lst.List(i) = value Then
            IsInList = True
            Exit Function
        End If
    Next i
    IsInList = False
End Function
