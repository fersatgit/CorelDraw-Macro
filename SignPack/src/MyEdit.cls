VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Edit As TextBox
Attribute Edit.VB_VarHelpID = -1
Private Sub Edit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Select Case KeyCode
  Case 8, 46, 48 To 58, 96 To 105
  Case Else
  KeyCode = 0
  End Select
End Sub
