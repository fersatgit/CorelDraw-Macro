VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorProbeForm 
   Caption         =   "Цветопроба"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   OleObjectBlob   =   "ColorProbeForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColorProbeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Function EditChange(ByRef Edit) As Long
  If Edit.Value = Empty Then
    EditChange = 0
  Else
    If Edit.Value > 100 Then
      Edit.Value = 100
    End If
    EditChange = Edit.Value
  End If
  UserForm_Initialize
End Function

Private Sub Steps_edit_KeyPress(ByVal Key As MSForms.ReturnInteger)
  Key = IIf((Key > 46) And (Key < 58), Key, 0)
End Sub

Private Sub C1_edit_Change()
  EditChange C1_edit
End Sub

Private Sub M1_edit_Change()
  EditChange M1_edit
End Sub

Private Sub Y1_edit_Change()
  EditChange Y1_edit
End Sub

Private Sub K1_edit_Change()
  EditChange K1_edit
End Sub

Private Sub C2_edit_Change()
  EditChange C2_edit
End Sub

Private Sub M2_edit_Change()
  EditChange M2_edit
End Sub

Private Sub Y2_edit_Change()
  EditChange Y2_edit
End Sub

Private Sub K2_edit_Change()
  EditChange K2_edit
End Sub

Private Sub Color1_Label_Click()
  With New Color
    If .UserAssignEx Then
      .ConvertToCMYK
      C1_edit.Value = .CMYKCyan
      M1_edit.Value = .CMYKMagenta
      Y1_edit.Value = .CMYKYellow
      K1_edit.Value = .CMYKBlack
      .ConvertToRGB
      Color1_Label.BackColor = RGB(.RGBRed, .RGBGreen, .RGBBlue)
    End If
  End With
End Sub

Private Sub Color2_Label_Click()
  With New Color
    If .UserAssignEx Then
      .ConvertToCMYK
      C2_edit.Value = .CMYKCyan
      M2_edit.Value = .CMYKMagenta
      Y2_edit.Value = .CMYKYellow
      K2_edit.Value = .CMYKBlack
      .ConvertToRGB
      Color2_Label.BackColor = RGB(.RGBRed, .RGBGreen, .RGBBlue)
    End If
  End With
End Sub

Private Sub ColorPicker1_Click()
  With ActiveDocument
    Dim X As Double
    Dim Y As Double
    Dim State As Long
    .GetUserClick X, Y, State, 0, False, cdrCursorEyeDrop
    With .SampleColorAtPoint(X, Y, cdrColorCMYK)
      C1_edit.Value = .CMYKCyan
      M1_edit.Value = .CMYKMagenta
      Y1_edit.Value = .CMYKYellow
      K1_edit.Value = .CMYKBlack
      .ConvertToRGB
      Color1_Label.BackColor = RGB(.RGBRed, .RGBGreen, .RGBBlue)
    End With
  End With
End Sub

Private Sub ColorPicker2_Click()
  With ActiveDocument
    Dim X As Double
    Dim Y As Double
    Dim State As Long
    .GetUserClick X, Y, State, 0, False, cdrCursorEyeDrop
    With .SampleColorAtPoint(X, Y, cdrColorCMYK)
      C2_edit.Value = .CMYKCyan
      M2_edit.Value = .CMYKMagenta
      Y2_edit.Value = .CMYKYellow
      K2_edit.Value = .CMYKBlack
      .ConvertToRGB
      Color2_Label.BackColor = RGB(.RGBRed, .RGBGreen, .RGBBlue)
    End With
  End With
End Sub

Private Sub Ok_buton_Click()
  Dim i As Long
  Dim XCoord As Double
  Dim YCoord As Double
  Dim C As Double
  Dim M As Double
  Dim Y As Double
  Dim K As Double
  Dim CStep As Double
  Dim MStep As Double
  Dim YStep As Double
  Dim KStep As Double
  Dim Steps As Double
  C = C1_edit.Value
  M = M1_edit.Value
  Y = Y1_edit.Value
  K = K1_edit.Value
  Steps = Steps_edit.Value + 1
  CStep = (C2_edit.Value - C) / Steps
  MStep = (M2_edit.Value - M) / Steps
  YStep = (Y2_edit.Value - Y) / Steps
  KStep = (K2_edit.Value - K) / Steps
  Optimization = True
  ActiveDocument.BeginCommandGroup ''
  ActiveDocument.Unit = cdrTenthMicron
  XCoord = ActiveWindow.ActiveView.OriginX - (Steps * 550000 + 500000) * 0.5
  YCoord = ActiveWindow.ActiveView.OriginY - 300000
  For i = Steps To 0 Step -1
    ActiveLayer.CreateRectangle2(XCoord, YCoord + 100000, 500000, 500000).Fill.ApplyUniformFill CreateCMYKColor(Round(C), Round(M), Round(Y), Round(K))
    ActiveLayer.CreateArtisticText(XCoord, YCoord, "C:" & CStr(Round(Abs(C))) & " M:" & CStr(Round(Abs(M))) & " Y:" & CStr(Round(Abs(Y))) & " K:" & CStr(Round(Abs(K)))).SetSizeEx XCoord, YCoord, 500000, 70000
    C = C + CStep
    M = M + MStep
    Y = Y + YStep
    K = K + KStep
    XCoord = XCoord + 550000
  Next
  ActiveDocument.EndCommandGroup
  Optimization = False
  Refresh
  ColorProbeForm.Hide
End Sub

Private Sub UserForm_Initialize()
  With CreateCMYKColor(C1_edit.Value, M1_edit.Value, Y1_edit.Value, K1_edit.Value)
    .ConvertToRGB
    Color1_Label.BackColor = RGB(.RGBRed, .RGBGreen, .RGBBlue)
  End With
  With CreateCMYKColor(C2_edit.Value, M2_edit.Value, Y2_edit.Value, K2_edit.Value)
    .ConvertToRGB
    Color2_Label.BackColor = RGB(.RGBRed, .RGBGreen, .RGBBlue)
  End With
  If VersionMajor * 256 + VersionMinor > &H1103 Then 'SampleColorAtPoint appear sinse 17.4
    ColorPicker1.Enabled = True
    ColorPicker2.Enabled = True
  End If
End Sub
