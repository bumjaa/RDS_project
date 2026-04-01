VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptionButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents optButton As MSForms.optionButton
Attribute optButton.VB_VarHelpID = -1

Private Sub optButton_Click()
    ' 클릭 시 동작 - T_Comments.UserForm의 OptionButton_Click 호출
    T_Comments.OptionButton_Click
End Sub
