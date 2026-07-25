VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Marks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Start()
  If ActiveSelectionRange.Count < 1 Then
    MsgBox "Выделите один или несколько объектов"
  Else
    MarksForm.Show
  End If
End Sub
