VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Instruments
   Caption         =   "Instruments"
   ClientHeight    =   8790.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16605
   OleObjectBlob   =   "Instruments.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "Instruments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public currentPrefix As String
Private m_InstrumentsData As Collection


Private Function GetInstrumentsData(Optional ByVal searchText As String) As Collection
    Dim url    As String
    Dim raw    As String
    Dim parsed As Object

    url = GetApiUrl("instruments_url") _
        & "?sheet=Instruments" _
        & "&search=" & URLEncodeUTF8(searchText)

    raw = HttpGet(url)
    Set parsed = JsonConverter.ParseJson(raw)
    Set GetInstrumentsData = parsed

End Function

Private Sub UserForm_Initialize()
    With Me.ListBox1
        .ColumnCount = 8
        .ColumnWidths = "50,200,120,100,105,70,70,60"
        .ColumnHeads = False
    End With

    With Me.ListBox2
        .ColumnCount = 8
        .ColumnWidths = "50,200,120,100,105,70,70,60"
        .ColumnHeads = False
    End With

    Me.TextBox1.Visible = True

    On Error Resume Next
    Set m_InstrumentsData = GetInstrumentsData()
    On Error GoTo 0

    If Not m_InstrumentsData Is Nothing Then
        LoadListBoxFromCollection Me.ListBox1, m_InstrumentsData
    End If

CleanExit:
    Exit Sub

ErrHandler:
    MsgBox "Failed to load Instruments data:" & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
End Sub


Private Sub UserForm_Activate()
    If currentPrefix <> "" Then
        'LoadData currentPrefix
    End If
End Sub

Private Sub LoadListBoxFromCollection(lb As MSForms.ListBox, rs As Collection)
  Dim i     As Long
  Dim rec   As Object
  Dim keys  As Variant

    keys = Array("Control_No", "Instrument_Name", "Model_Name", "Manufacturer", "Serial_No", "Previous_Cal", "Recent_Cal", "Cal_Periodic")

    lb.Clear
    If rs.Count = 0 Then Exit Sub

    For i = 1 To rs.Count
        Set rec = rs(i)
        lb.AddItem CStr(rec(keys(0)))
        lb.List(lb.ListCount - 1, 1) = CStr(rec(keys(1)))
        lb.List(lb.ListCount - 1, 2) = CStr(rec(keys(2)))
        lb.List(lb.ListCount - 1, 3) = CStr(rec(keys(3)))
        lb.List(lb.ListCount - 1, 4) = CStr(rec(keys(4)))
        lb.List(lb.ListCount - 1, 5) = CStr(rec(keys(5)))
        lb.List(lb.ListCount - 1, 6) = CStr(rec(keys(6)))
        lb.List(lb.ListCount - 1, 7) = CStr(rec(keys(7)))
    Next i

End Sub

Private Sub LoadListBox(ByRef listBox As MSForms.ListBox, ByVal rs As Object)
    Dim i As Integer

    listBox.Clear

    If Not rs.EOF Then
        Do While Not rs.EOF
            listBox.AddItem
            For i = 0 To listBox.ColumnCount - 1
                listBox.List(listBox.ListCount - 1, i) = rs.Fields(i).Value
            Next i
            rs.MoveNext
        Loop
    End If
End Sub


Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With TextBox1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TextBox1_Change()
    Dim f As String
    Dim filtered As Collection

    f = Trim(Me.TextBox1.Text)
    Set filtered = New Collection

    Dim rec As Object, key As Variant
    Dim keys As Variant
    keys = Array("Control_No", "Instrument_Name", "Model_Name", "Manufacturer", "Serial_No")

    If Not m_InstrumentsData Is Nothing Then
      For Each rec In m_InstrumentsData
        For Each key In keys
          If InStr(1, CStr(rec(key)), f, vbTextCompare) > 0 Then
            filtered.Add rec
            Exit For
          End If
        Next
      Next
    End If

    LoadListBoxFromCollection Me.ListBox1, filtered
End Sub



Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer

    If Me.ListBox1.ListIndex >= 0 Then
        Me.ListBox2.AddItem
        For i = 0 To Me.ListBox1.ColumnCount - 1
            Me.ListBox2.List(Me.ListBox2.ListCount - 1, i) = Me.ListBox1.List(Me.ListBox1.ListIndex, i)
        Next i
    Else
        MsgBox "Please select an item.", vbExclamation
    End If
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedIndex As Integer

    selectedIndex = Me.ListBox2.ListIndex
    If selectedIndex >= 0 Then
        Me.ListBox2.RemoveItem selectedIndex
    Else
        MsgBox "Please select an item to remove.", vbExclamation
    End If
End Sub

Public Sub LoadData(ByVal prefix As String)
    Dim ws As Worksheet
    Dim instrumentsRange As Range
    Dim data As Variant
    Dim i As Long

    currentPrefix = prefix

    Set ws = ActiveSheet
    On Error Resume Next
    Set instrumentsRange = ws.Range(prefix & "_INSTRUMENTS")
    On Error GoTo 0

    Me.ListBox2.Clear

    If Not instrumentsRange Is Nothing And instrumentsRange.Rows.Count > 1 Then
        data = instrumentsRange.Offset(1, 0).Resize(instrumentsRange.Rows.Count - 1, instrumentsRange.Columns.Count).Value
        For i = 1 To UBound(data, 1)
            If Trim(data(i, 1)) <> "" Then
                Me.ListBox2.AddItem
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 1) = data(i, 1)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 2) = data(i, 3)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 3) = data(i, 5)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 4) = data(i, 6)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 5) = data(i, 7)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 6) = data(i, 8)
            End If
        Next i
    Else
        MsgBox prefix & " instruments data not found.", vbExclamation
    End If
End Sub

Private Sub ListBox1_Click()
    Dim i As Integer
    For i = 0 To Me.ListBox2.ListCount - 1
        Me.ListBox2.Selected(i) = False
    Next i
End Sub

Private Sub ListBox2_Click()
    Dim i As Integer
    For i = 0 To Me.ListBox1.ListCount - 1
        Me.ListBox1.Selected(i) = False
    Next i
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub ListBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim instrRange As Range
    Dim lastDataRow As Long, newRows As Long, requiredRows As Long, rowsToAdd As Long
    Dim startRow As Long, i As Long, j As Long
    Dim hasData As Boolean

    Dim envRange As Range, envData As Variant, envDates As Collection, minDate As Date
    Dim rcal As Date, pcal As Date, period As Double, nextCal As Date, calDate As Date

    Set ws = ActiveSheet
    Set instrRange = ws.Range(currentPrefix & "_INSTRUMENTS")

    ' 1. ENV date minimum
    On Error Resume Next
    Set envRange = ws.Range(currentPrefix & "_ENV").Offset(1, 3)
    envData = envRange.Resize(instrRange.Rows.Count - 1, 1).Value
    On Error GoTo 0

    Set envDates = New Collection
    If Not IsEmpty(envData) Then
        If IsArray(envData) Then
            For i = 1 To UBound(envData, 1)
                If IsDate(envData(i, 1)) Then envDates.Add DateValue(envData(i, 1))
            Next i
        Else
            If IsDate(envData) Then envDates.Add DateValue(envData)
        End If
    End If

    If envDates.Count > 0 Then
        minDate = envDates(1)
        For i = 2 To envDates.Count
            If envDates(i) < minDate Then minDate = envDates(i)
        Next i
    Else
        minDate = Date
    End If

    ' 2. Find last data row
    lastDataRow = 1
    For i = 2 To instrRange.Rows.Count
        If Trim(CStr(instrRange.Cells(i, 1).Value)) <> "" Then
            lastDataRow = i
        End If
    Next i

    newRows = Me.ListBox2.ListCount
    requiredRows = lastDataRow + newRows

    If instrRange.Rows.Count < requiredRows Then
        rowsToAdd = requiredRows - instrRange.Rows.Count
        ExpandRange ws, instrRange, currentPrefix & "_INSTRUMENTS", rowsToAdd
        Set instrRange = ws.Range(currentPrefix & "_INSTRUMENTS")
    End If

    startRow = instrRange.Row + lastDataRow

    Application.EnableEvents = False
    For i = 0 To newRows - 1
        Dim vP As Variant, vR As Variant, vPeriod As Variant
        Dim baseCal As Variant

        With ws
            .Cells(startRow + i, instrRange.Column + 0).Value = Me.ListBox2.List(i, 1)
            .Cells(startRow + i, instrRange.Column + 2).Value = Me.ListBox2.List(i, 2)
            .Cells(startRow + i, instrRange.Column + 4).Value = Me.ListBox2.List(i, 3)
            .Cells(startRow + i, instrRange.Column + 5).Value = Me.ListBox2.List(i, 4)
        End With

        vP = Empty: vR = Empty: vPeriod = Empty
        If IsDate(Me.ListBox2.List(i, 5)) Then vP = DateValue(CDate(Me.ListBox2.List(i, 5)))
        If IsDate(Me.ListBox2.List(i, 6)) Then vR = DateValue(CDate(Me.ListBox2.List(i, 6)))
        If Len(Trim(CStr(Me.ListBox2.List(i, 7)))) > 0 And IsNumeric(Me.ListBox2.List(i, 7)) Then _
            vPeriod = CDbl(Me.ListBox2.List(i, 7))

        baseCal = Empty
        If Not IsEmpty(vR) And vR <= minDate Then
            baseCal = vR
        ElseIf Not IsEmpty(vP) And vP <= minDate Then
            baseCal = vP
        ElseIf Not IsEmpty(vR) Then
            baseCal = vR
        ElseIf Not IsEmpty(vP) Then
            baseCal = vP
        End If

        With ws
            If IsEmpty(baseCal) Then
                .Cells(startRow + i, instrRange.Column + 6).Value = "N/A"
                .Cells(startRow + i, instrRange.Column + 7).Value = "N/A"
            Else
                .Cells(startRow + i, instrRange.Column + 6).Value = baseCal
                If IsEmpty(vPeriod) Then
                    .Cells(startRow + i, instrRange.Column + 7).Value = "N/A"
                Else
                    nextCal = DateAdd("yyyy", CLng(vPeriod), baseCal) - 1
                    .Cells(startRow + i, instrRange.Column + 7).Value = nextCal
                End If
            End If

            If IsEmpty(vPeriod) Then
                .Cells(startRow + i, instrRange.Column + 8).Value = "N/A"
            Else
                .Cells(startRow + i, instrRange.Column + 8).Value = vPeriod
            End If

            .Cells(startRow + i, instrRange.Column + 9).Value = 1
        End With
    Next i

    CheckValidation currentPrefix
    Application.EnableEvents = True
    Unload Me
End Sub
