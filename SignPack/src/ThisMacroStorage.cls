VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Start()
  If ActiveSelectionRange.Count = 0 Then
    MsgBox "Ничего не выбрано"
  Else
    On Error GoTo Finish
    Optimization = True
    ActiveDocument.BeginCommandGroup "Pack"
    PackForm.Show
Finish:
    ActiveDocument.EndCommandGroup
    Optimization = False
    Refresh
  End If
End Sub

