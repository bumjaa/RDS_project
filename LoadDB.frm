VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadDB
   Caption         =   "Loading DataBase"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   OleObjectBlob   =   "LoadDB.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "LoadDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim selectedOrderNo As String
    If ListBox1.ListIndex <> -1 Then
        selectedOrderNo = ListBox1.List(ListBox1.ListIndex, 0)
        Call loadDBByOrderNo(selectedOrderNo)
    Else
        MsgBox "Please select an item."
    End If
    Unload Me

End Sub

Private Sub UserForm_Initialize()
    Call SearchDB("")
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call SearchDB(TextBox1.Text)
        KeyCode = 0
    End If
End Sub

Private Sub SearchDB(ByVal searchText As String)
    Dim sConn As String, strSQL As String
    Dim cN As Object, rs As Object
    Dim dataArr As Variant, result() As Variant
    Dim rowCount As Long, i As Long, j As Long

    sConn = GetPersonalDBConn()

    Set cN = CreateObject("ADODB.Connection")
    cN.Open sConn

    If searchText = "" Then
        strSQL = "SELECT Order_No, Applicant, Model_Name, Product_Name FROM Personal_DB"
    Else
        strSQL = "SELECT Order_No, Applicant, Model_Name, Product_Name FROM Personal_DB " & _
                 "WHERE Order_No LIKE '%" & Replace(searchText, "'", "''") & "%' " & _
                 "OR Applicant LIKE '%" & Replace(searchText, "'", "''") & "%' " & _
                 "OR Model_Name LIKE '%" & Replace(searchText, "'", "''") & "%' " & _
                 "OR Product_Name LIKE '%" & Replace(searchText, "'", "''") & "%'"
    End If

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, cN, 1, 1

    With Me.ListBox1
        .Clear
        .ColumnCount = 4
    End With

    If Not rs.EOF Then
        dataArr = rs.GetRows
        rowCount = UBound(dataArr, 2) - LBound(dataArr, 2) + 1

        ReDim result(0 To rowCount - 1, 0 To 3)
        For i = 0 To rowCount - 1
            For j = 0 To 3
                result(i, j) = dataArr(j, i)
            Next j
        Next i

        Me.ListBox1.List = result
    End If

    rs.Close
    cN.Close
End Sub
