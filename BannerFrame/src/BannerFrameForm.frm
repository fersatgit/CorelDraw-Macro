VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BannerFrameForm 
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   OleObjectBlob   =   "BannerFrameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BannerFrameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Const PI = 3.14
Const PI2 = 6.28

Dim Points() As CurveElement
Dim PointsCount As Long

Private Sub HemHorizontal_KeyPress(ByVal Key As MSForms.ReturnInteger)
  Key = IIf((Key > 46) And (Key < 58), Key, 0)
End Sub

Private Sub HemVertical_KeyPress(ByVal Key As MSForms.ReturnInteger)
  Key = IIf((Key > 46) And (Key < 58), Key, 0)
End Sub

Private Sub MarkdownDistance_KeyPress(ByVal Key As MSForms.ReturnInteger)
  Key = IIf((Key > 46) And (Key < 58), Key, 0)
End Sub

Private Sub BannerFormRectangle_Change()
  Dim enabled As Boolean
  enabled = BannerFormRectangle.Value And Grommets.Value
  GrommetsOnPerimeter.enabled = enabled
  GrommetsInCorners.enabled = enabled
  If BannerFormRectangle.Value = False Then
    GrommetsOnPerimeter.Value = True
    GrommetsInCorners.Value = False
    HemVertical.Value = HemHorizontal.Value
  End If
  HemVertical.enabled = BannerFormRectangle.Value
  HemHorizontal.enabled = BannerFormRectangle.Value
End Sub

Private Sub Grommets_Change()
  With Grommets
    Dim enabled As Boolean
    Markdown.enabled = .Value
    Markdown.Value = Markdown.Value And .Value
    enabled = .Value And BannerFormRectangle.Value
    GrommetsOnPerimeter.enabled = enabled
    GrommetsInCorners.enabled = enabled
    enabled = .Value And Markdown.Value
    MarkdownInside.enabled = enabled
    MarkdownOutside.enabled = enabled
    MarkdownDistanceCaption.enabled = enabled
    MarkdownDistance.enabled = enabled
  End With
End Sub

Private Sub Markdown_Change()
  With Markdown
    MarkdownInside.enabled = .Value
    MarkdownOutside.enabled = .Value
    MarkdownDistanceCaption.enabled = .Value
    MarkdownDistance.enabled = .Value
  End With
End Sub

Private Sub AddPoint(ByRef CurveElement As CurveElement, ByVal x As Long, ByVal y As Long, ByVal typ As Long)
  With CurveElement
    .ElementType = typ
    .PositionX = x
    .PositionY = y
    .Flags = cdrFlagValid + cdrFlagUser
  End With
End Sub

'for older versions of CorelDraw compatibility
Private Sub AppendSubpathCircle(x As Double, y As Double, r As Double)
  AddPoint Points(PointsCount), x + r, y, cdrElementStart
  AddPoint Points(PointsCount + 1), x + r, y - r * 0.551785, cdrElementControl
  AddPoint Points(PointsCount + 2), x + r * 0.551785, y - r, cdrElementControl
  AddPoint Points(PointsCount + 3), x, y - r, cdrElementCurve
  AddPoint Points(PointsCount + 4), x - r * 0.551785, y - r, cdrElementControl
  AddPoint Points(PointsCount + 5), x - r, y - r * 0.551785, cdrElementControl
  AddPoint Points(PointsCount + 6), x - r, y, cdrElementCurve
  AddPoint Points(PointsCount + 7), x - r, y + r * 0.551785, cdrElementControl
  AddPoint Points(PointsCount + 8), x - r * 0.551785, y + r, cdrElementControl
  AddPoint Points(PointsCount + 9), x, y + r, cdrElementCurve
  AddPoint Points(PointsCount + 10), x + r * 0.551785, y + r, cdrElementControl
  AddPoint Points(PointsCount + 11), x + r, y + r * 0.551785, cdrElementControl
  AddPoint Points(PointsCount + 12), x + r, y, cdrElementCurve
  Points(PointsCount + 12).Flags = cdrFlagValid + cdrFlagUser + cdrFlagClosed
  PointsCount = PointsCount + 13
End Sub

Private Sub OkButton_Click()
  BannerFrameForm.Hide
  Optimization = True
  With ActiveDocument
    .BeginCommandGroup "Banner"
    .Unit = cdrTenthMicron
    PointsCount = 0
    With ActiveSelectionRange.UngroupAllEx
      Dim w As Double
      Dim h As Double
      Dim x As Double
      Dim y As Double
      Dim cx As Double
      Dim cy As Double
      Dim HemH As Double
      Dim HemV As Double
      Dim i As Long
      Dim r As Double
      Dim a As Double
      Dim gx1 As Double
      Dim gx2 As Double
      Dim gy1 As Double
      Dim gy2 As Double
      Dim hcount As Long
      Dim vcount As Long
      Dim count As Long
      Dim step As Double
      Dim MarkdownCurve As Curve
      Dim tmpShape As Shape
      Dim MarkDist As Double
      Dim GroupName As String
      w = .SizeWidth
      h = .SizeHeight
      x = .PositionX
      y = .PositionY - h
      HemH = HemHorizontal.Value * 10000
      HemV = HemVertical.Value * 10000
      MarkDist = MarkdownDistance * 10000
      GroupName = Label.Text + " " + CStr(Round(w * 0.0001)) + "x" + CStr(Round(h * 0.0001))
'''''''''''Rectangle
      If BannerFormRectangle.Value Then
Rectangle:
        Set tmpShape = ActiveLayer.CreateRectangle2(x, y, w, h)
        .Add tmpShape
        tmpShape.Outline.Width = 762
        .Add ActiveLayer.CreateRectangle2(x - HemH, y - HemV, w + HemH + HemH, h + HemV + HemV)
        If HemDoubled.Value Then
          .Add ActiveLayer.CreateRectangle2(x - HemH * 3, y - HemV * 3, w + HemH * 6, h + HemV * 6)
        End If
        Set tmpShape = ActiveLayer.CreateArtisticText(0, 0, GroupName + IIf(Grommets.Value, " люверсы ", " ") + IIf(Grommets.Value And GrommetsInCorners.Value, "по углам", ""), cdrRussian, cdrCharSetANSI, "Arial")
        tmpShape.SetBoundingBox x, y + h + HemV * 0.2, w, HemV * 0.6, True, cdrBottomMiddle
        tmpShape.Fill.ApplyUniformFill CreateCMYKColor(5, 5, 5, 5)
        .Add tmpShape
        Set tmpShape = tmpShape.Duplicate
        tmpShape.Rotate 180
        tmpShape.SetPosition x + (w - tmpShape.SizeWidth) * 0.5, y - HemV * 0.2
        .Add tmpShape
        Set tmpShape = tmpShape.Duplicate
        tmpShape.Rotate 270
        tmpShape.SetBoundingBox x - HemH * 0.8, y, HemH * 0.6, h, True, cdrMiddleRight
        .Add tmpShape
        Set tmpShape = tmpShape.Duplicate
        tmpShape.Rotate 180
        tmpShape.SetPosition x + w + HemH * 0.2, y + (h + tmpShape.SizeHeight) * 0.5
        .Add tmpShape
        If Markdown.Value And (MarkDist > 0) Then
          If GrommetsInCorners.Value Then
            hcount = 1
            vcount = 1
          Else
            hcount = Fix((w - 500000) / MarkDist + 0.99)
            vcount = Fix((h - 500000) / MarkDist + 0.99)
          End If
          gx1 = x + 250000
          gx2 = gx1
          gy1 = y + 250000
          gy2 = y + h - 250000
          ReDim Points((vcount + hcount + 2) * 26) As CurveElement
          If MarkdownOutside Then
            gy1 = gy1 - 500000
            gy2 = gy2 + 500000
          End If
          If hcount > 0 Then
            step = (w - 500000) / hcount
            For i = 0 To hcount
              AppendSubpathCircle gx2, gy1, 50000
              AppendSubpathCircle gx2, gy2, 50000
              gx2 = gx2 + step
            Next
            gx2 = gx2 - step
          End If
          If vcount > 0 Then
            step = (h - 500000) / vcount
            If MarkdownOutside Then
              gx1 = gx1 - 500000
              gx2 = gx2 + 500000
              gy1 = gy1 + 500000
              gy2 = gy2 - 500000
            Else
              vcount = vcount - 2
              gy1 = gy1 + step
            End If
            For i = 0 To vcount
              AppendSubpathCircle gx1, gy1, 50000
              AppendSubpathCircle gx2, gy1, 50000
              gy1 = gy1 + step
            Next
          End If
          If PointsCount > 0 Then
            Set MarkdownCurve = ActiveDocument.CreateCurve
            MarkdownCurve.PutCurveInfo Points, PointsCount
            .Add ActiveLayer.CreateCurve(MarkdownCurve)
            With .Shapes.Last.Fill.ApplyPatternFill(cdrTwoColorPattern)
              .TileWidth = 50000
              .TileHeight = 50000
            End With
          End If
          Erase Points
        End If
'''''''''''Circle
      ElseIf BannerFormCircle.Value Then
        cx = x + w * 0.5
        cy = y + h * 0.5
        r = IIf(w > h, w, h) * 0.5
        Set tmpShape = ActiveLayer.CreateEllipse2(cx, cy, r, r)
        .Add tmpShape
        tmpShape.Outline.Width = 762
        .Add ActiveLayer.CreateEllipse2(cx, cy, r + HemH, r + HemV)
        If HemDoubled.Value Then
          .Add ActiveLayer.CreateEllipse2(cx, cy, r + HemH * 3, r + HemV * 3)
        End If
        If Markdown.Value And (MarkDist > 0) Then
          r = r - 250000
          If MarkdownOutside.Value Then
            r = r + 500000
          End If
          count = Fix(PI * (w - 500000) / MarkDist + 0.99)
          If count > 0 Then
            step = PI2 / count
            ReDim Points(count * 13) As CurveElement
            With MarkdownCurve
              a = 0
              Do
                AppendSubpathCircle cx + Sin(a) * r, cy + Cos(a) * r, 50000
                a = a + step
              Loop Until Round(a, 2) >= PI2
            End With
            Set MarkdownCurve = ActiveDocument.CreateCurve
            MarkdownCurve.PutCurveInfo Points, PointsCount
            Erase Points
            .Add ActiveLayer.CreateCurve(MarkdownCurve)
            With .Shapes.Last.Fill.ApplyPatternFill(cdrTwoColorPattern)
              .TileWidth = 50000
              .TileHeight = 50000
            End With
          End If
        End If
'''''''''''Curve
      Else
        r = 0
        For i = 1 To .Shapes.count
          If (.Shapes(i).SizeWidth = .SizeWidth) And (.Shapes(i).SizeHeight = .SizeHeight) Then
            Set tmpShape = .Shapes(i)
          End If
        Next
        If tmpShape Is Nothing Then
          GoTo Rectangle
        Else
          With tmpShape
            Select Case .Type
              Case cdrBitmapShape
                  Set tmpShape = ActiveLayer.CreateCurve(.Bitmap.CropEnvelope)
              Case cdrCurveShape
                  Set tmpShape = ActiveLayer.CreateCurve(.Curve)
              Case cdrEllipseShape, cdrPolygonShape
                  .ConvertToCurves
                  Set tmpShape = ActiveLayer.CreateCurve(.Curve)
              Case Else
                  GoTo Rectangle
            End Select
          End With
        End If
        .Add tmpShape
        .Add tmpShape.CreateContour(cdrContourOutside, HemH, 1, cdrDirectFountainFillBlend, Nothing, Nothing, Nothing, 0, 0, cdrContourRoundCap, cdrContourCornerRound).Separate(1)
        If HemDoubled.Value Then
          .Add tmpShape.CreateContour(cdrContourOutside, HemH * 3, 1, cdrDirectFountainFillBlend, Nothing, Nothing, Nothing, 0, 0, cdrContourRoundCap, cdrContourCornerRound).Separate(1)
        End If
        If Markdown.Value And (MarkDist > 0) Then
          If MarkdownOutside.Value Then
            With tmpShape.CreateContour(cdrContourOutside, 250000, 1, cdrDirectFountainFillBlend, Nothing, Nothing, Nothing, 0, 0, cdrContourRoundCap, cdrContourCornerRound).Separate
              tmpShape.Delete
              Set tmpShape = .Item(1)
            End With
          Else
            With tmpShape.CreateContour(cdrContourInside, 250000, 1, cdrDirectFountainFillBlend, Nothing, Nothing, Nothing, 0, 0, cdrContourRoundCap, cdrContourCornerRound).Separate
              tmpShape.Delete
              Set tmpShape = .Item(1)
            End With
          End If
          count = Fix(tmpShape.Curve.Length / MarkDist + 0.99)
          If count > 2 Then
            With ActiveLayer.CreateEllipse2(0, 0, 50000).CreateBlend(ActiveLayer.CreateEllipse2(0, 0, 50000), count - 2, cdrDirectFountainFillBlend, cdrBlendSteps, count - 1, 0, True, tmpShape).Separate.UngroupAllEx
              tmpShape.Delete
              Set tmpShape = .Combine
            End With
            .Add tmpShape
            With tmpShape.Fill.ApplyPatternFill(cdrTwoColorPattern)
              .TileWidth = 50000
              .TileHeight = 50000
            End With
          End If
        Else
          tmpShape.Delete
        End If
      End If
      With .Group
        .Name = GroupName
        .CreateSelection
      End With
    End With
    .EndCommandGroup
  End With
  Optimization = False
  Refresh
End Sub

Private Sub UserForm_Activate()
  If ActiveSelectionRange.count = 0 Then
    BannerFrameForm.Hide
  End If
End Sub
