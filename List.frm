VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} List 
   Caption         =   "List"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5010
   OleObjectBlob   =   "List.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim selectedSheetName As String
    If ListBox1.ListIndex <> -1 Then
        selectedSheetName = ListBox1.value
        Call LayOut_Helper(selectedSheetName)
        Unload Me
    Else
        MsgBox "시트를 선택하세요.", vbExclamation
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim ws As Worksheet
    ListBox1.Clear
    For Each ws In ThisWorkbook.Worksheets
        If (ws.Name <> "Main" And ws.Name <> "Layout") And ws.Visible <> xlSheetVeryHidden Then
            ListBox1.AddItem ws.Name
        End If
    Next ws
    
End Sub
