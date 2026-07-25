VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MarksForm 
   Caption         =   "Метки"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   OleObjectBlob   =   "MarksForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MarksForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddPoint(ByRef CurveElement As CurveElement, ByVal x As Long, ByVal y As Long, ByVal typ As Long)
  With CurveElement
    .ElementType = typ
    .PositionX = x
    .PositionY = y
    .Flags = cdrFlagValid + cdrFlagUser
  End With
End Sub

Private Sub HCount_KeyPress(ByVal Key As MSForms.ReturnInteger)
  Key = IIf((Key > 46) And (Key < 58), Key, 0)
End Sub

Private Sub VCount_KeyPress(ByVal Key As MSForms.ReturnInteger)
  Key = IIf((Key > 46) And (Key < 58), Key, 0)
End Sub

Private Sub Offset_KeyPress(ByVal Key As MSForms.ReturnInteger)
  Key = IIf((Key > 46) And (Key < 58), Key, 0)
End Sub

Private Sub MarksSize_KeyPress(ByVal Key As MSForms.ReturnInteger)
  Key = IIf((Key > 46) And (Key < 58), Key, 0)
End Sub

Private Sub OK_Click()
  Dim curve As curve
  Dim L As Double, B As Double, R As Double, T As Double, stepx As Double, stepy As Double, x As Double, y As Double
  Dim HNum As Long, VNum As Long, Ofs As Long, Size As Long, i As Long, j As Long, k As Long
  If (HCount.Text = "") Or (HCount.Value < 2) Then
    HCount.Value = 2
  End If
  If (VCount.Text = "") Or (VCount.Value < 2) Then
    VCount.Value = 2
  End If
  If Offset.Text = "" Then
    Offset.Value = 0
  End If
  If (MarksSize.Text = "") Or (MarksSize.Value < 1) Then
    MarksSize.Value = 1
  End If
  HNum = HCount.Value
  VNum = VCount.Value
  Ofs = Offset.Value * 10000
  Size = MarksSize.Value * 10000
  MarksForm.Hide
  ReDim CurveElements(HNum * VNum * 4) As CurveElement
  ActiveDocument.BeginCommandGroup "Marks"
  ActiveDocument.Unit = cdrTenthMicron
  Set curve = CreateCurve
  k = 0
  ActiveSelectionRange.GetBoundingBox L, B, R, T, True
  R = L + R + Ofs
  T = B + T + Ofs
  L = L - Ofs
  B = B - Ofs
'Vertical
  stepy = (T - B) / (VNum - 1)
  y = stepy + B - Size
  For i = 3 To VNum
    AddPoint CurveElements(k), L, y, cdrElementStart
    AddPoint CurveElements(k + 1), L, y + Size + Size, cdrElementLine
    AddPoint CurveElements(k + 2), L, y + Size, cdrElementStart
    AddPoint CurveElements(k + 3), L + Size, y + Size, cdrElementLine
    AddPoint CurveElements(k + 4), R, y, cdrElementStart
    AddPoint CurveElements(k + 5), R, y + Size + Size, cdrElementLine
    AddPoint CurveElements(k + 6), R, y + Size, cdrElementStart
    AddPoint CurveElements(k + 7), R - Size, y + Size, cdrElementLine
    y = y + stepy
    k = k + 8
  Next
'Horizontal
  stepx = (R - L) / (HNum - 1)
  x = stepx + L - Size
  For i = 3 To HNum
    AddPoint CurveElements(k), x, B, cdrElementStart
    AddPoint CurveElements(k + 1), x + Size + Size, B, cdrElementLine
    AddPoint CurveElements(k + 2), x + Size, B, cdrElementStart
    AddPoint CurveElements(k + 3), x + Size, B + Size, cdrElementLine
    AddPoint CurveElements(k + 4), x, T, cdrElementStart
    AddPoint CurveElements(k + 5), x + Size + Size, T, cdrElementLine
    AddPoint CurveElements(k + 6), x + Size, T, cdrElementStart
    AddPoint CurveElements(k + 7), x + Size, T - Size, cdrElementLine
    x = x + stepx
    k = k + 8
  Next
'Inner
  If InnerMarks.Value Then
    y = stepy + B - Size
    For i = 3 To VNum
      x = stepx + L - Size
      For j = 3 To HNum
        AddPoint CurveElements(k), x, y + Size, cdrElementStart
        AddPoint CurveElements(k + 1), x + Size + Size, y + Size, cdrElementLine
        AddPoint CurveElements(k + 2), x + Size, y, cdrElementStart
        AddPoint CurveElements(k + 3), x + Size, y + Size + Size, cdrElementLine
        x = x + stepx
        k = k + 4
      Next
      y = y + stepy
    Next
  End If
'Cornenrs
  AddPoint CurveElements(k), L, B + Size, cdrElementStart
  AddPoint CurveElements(k + 1), L, B, cdrElementLine
  AddPoint CurveElements(k + 2), L + Size, B, cdrElementLine
  AddPoint CurveElements(k + 3), R - Size, B, cdrElementStart
  AddPoint CurveElements(k + 4), R, B, cdrElementLine
  AddPoint CurveElements(k + 5), R, B + Size, cdrElementLine
  AddPoint CurveElements(k + 6), R, T - Size, cdrElementStart
  AddPoint CurveElements(k + 7), R, T, cdrElementLine
  AddPoint CurveElements(k + 8), R - Size, T, cdrElementLine
  AddPoint CurveElements(k + 9), L + Size, T, cdrElementStart
  AddPoint CurveElements(k + 10), L, T, cdrElementLine
  AddPoint CurveElements(k + 11), L, T - Size, cdrElementLine
  curve.PutCurveInfo CurveElements, k + 12
  ActiveLayer.CreateCurve(curve).SetPosition L, T
  ActiveDocument.EndCommandGroup
  Refresh
End Sub
