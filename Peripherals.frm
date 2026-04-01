VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Peripherals
   Caption         =   "Peripherals"
   ClientHeight    =   8205.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15435
   OleObjectBlob   =   "Peripherals.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "Peripherals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public currentPrefix As String
Private m_InstrumentsData As Collection

Private Sub UserForm_Activate()
    If currentPrefix <> "" Then
        LoadData currentPrefix
    End If
End Sub


Private Function GetInstrumentsData(Optional ByVal searchText As String) As Collection
    Dim url    As String
    Dim raw    As String
    Dim parsed As Object

    url = GetApiUrl("instruments_url") _
        & "?sheet=Peripherals" _
        & "&search=" & URLEncodeUTF8(searchText)

    raw = HttpGet(url)
    Set parsed = JsonConverter.ParseJson(raw)
    Set GetInstrumentsData = parsed

End Function

Private Sub UserForm_Initialize()
    Dim query As String
    Dim rs As Object

    With Me.ListBox1
        .ColumnCount = 6
        .ColumnWidths = "80,80,120,80,80,160"
        .ColumnHeads = False
    End With

    With Me.ListBox2
        .ColumnCount = 6
        .ColumnWidths = "80,80,120,80,80,160"
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
    MsgBox "Failed to load Peripherals data:" & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub LoadListBoxFromCollection(lb As MSForms.ListBox, rs As Collection)
  Dim i     As Long
  Dim rec   As Object
  Dim keys  As Variant

  keys = Array("Instrument_Name", "Model_Name", "Manufacturer", "Serial_No", "Remarks", "Port")

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
  Next i
End Sub

Private Sub TextBox1_Change()
    Dim f As String
    Dim filtered As Collection

    f = Trim(Me.TextBox1.Text)
    Set filtered = New Collection

    Dim rec As Object, key As Variant
    Dim keys As Variant
    keys = Array("Instrument_Name", "Model_Name", "Manufacturer", "Serial_No")

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
    Set instrumentsRange = ws.Range("Total_Config")
    On Error GoTo 0

    Me.ListBox2.Clear

    If Not instrumentsRange Is Nothing And instrumentsRange.Rows.Count > 1 Then
        data = instrumentsRange.Offset(1, 0).Resize(instrumentsRange.Rows.Count - 1, instrumentsRange.Columns.Count).Value
        For i = 1 To UBound(data, 1)
            If Trim(data(i, 1)) <> "" Then
                Me.ListBox2.AddItem
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 0) = data(i, 1)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 1) = data(i, 3)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 2) = data(i, 5)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 3) = data(i, 7)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 4) = data(i, 9)
            End If
        Next i
    Else
        MsgBox prefix & " peripherals data not found.", vbExclamation
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
    Dim instrumentsRange As Range
    Dim i As Long, j As Integer
    Dim requiredRows As Long
    Dim listCount As Long

    Set ws = ActiveSheet
    On Error Resume Next
    Set instrumentsRange = ws.Range("Total_Config")
    On Error GoTo 0

    listCount = Me.ListBox2.ListCount
    If listCount = 0 Then
        MsgBox "No items in ListBox2.", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    requiredRows = listCount - instrumentsRange.Rows.Count + 1
    If requiredRows > 0 Then
        ExpandRange ws, instrumentsRange, "Total_Config", requiredRows
    End If

    Set instrumentsRange = ws.Range("Total_Config")

    For i = 1 To listCount
        instrumentsRange.Cells(i + 1, 1).Value = Me.ListBox2.List(i - 1, 0)
        instrumentsRange.Cells(i + 1, 3).Value = Me.ListBox2.List(i - 1, 1)
        instrumentsRange.Cells(i + 1, 5).Value = Me.ListBox2.List(i - 1, 2)
        instrumentsRange.Cells(i + 1, 7).Value = Me.ListBox2.List(i - 1, 3)
    Next i

    Application.EnableEvents = True

    Unload Me

End Sub
