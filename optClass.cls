VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "optClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents ctlOptionButton As MSForms.optionButton
Attribute ctlOptionButton.VB_VarHelpID = -1

Private Sub ctlOptionButton_Click()
    Dim selectedCaption As String
    selectedCaption = ctlOptionButton.Caption
    Call G_Remarks.optSearch
    
End Sub
