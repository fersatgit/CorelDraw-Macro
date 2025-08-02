VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PackForm 
   Caption         =   "SignPack 1.0"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "PackForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PackForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gap As Double, Margins As Double, TopMargin As Double, MarkMargins As Double, RoundRadius As Double, ContourOffset As Double
Dim PageWidth As Double, PageHeight As Double, PackWidth As Double, PackHeight As Double
Dim startX As Double, endX As Double, startY As Double, endY As Double
Dim Width As Double, Height As Double, HalfWidth As Double
Dim SummMargins As Double, CenterX As Double, CenterY As Double, radius As Double
Dim ContourLayer As Layer, CutLayer As Layer, PrintLayer As Layer, Symbol As ShapeRange, All As ShapeRange
Dim Edits(8) As MyEdit
Dim CurveElements() As CurveElement
Dim n As Long, i As Long, j As Long, x As Double, y As Double, Rows As Long

Private Function ToTenthMicron(text As String) As Long
  If text <> "" Then
    ToTenthMicron = CLng(text) * 10000
  Else
    ToTenthMicron = 0
  End If
End Function

Private Sub ReadParams()
  ActiveDocument.Unit = cdrTenthMicron
  PageWidth = ToTenthMicron(PageWidth_edit.text)
  PageHeight = ToTenthMicron(PageHeight_edit.text)
  Gap = ToTenthMicron(Gap_edit.text)
  Margins = ToTenthMicron(Margins_edit.text)
  TopMargin = ToTenthMicron(TopMargin_edit.text)
  MarkMargins = ToTenthMicron(MarkMargins_edit.text)
  SummMargins = Margins + MarkMargins
  RoundRadius = ToTenthMicron(RoundRadius_Edit.text)
  ContourOffset = ToTenthMicron(ContourOffset_edit.text)
  ActiveSelection.GetSize Width, Height
  radius = IIf(Width > Height, Width, Height) * 0.5
  Width = Width + Gap
  Height = Height + Gap
End Sub

Private Sub init(ShapeType As Long)
  Width = Int(Width)
  Height = Int(Height)
  HalfWidth = Width * 0.5
  startX = (PageWidth - SummMargins * 2) Mod Width
  If (ShapeType <> 2) And (startX >= HalfWidth) Then
    startX = startX - HalfWidth
  End If
  startX = startX * 0.5
  endX = Int(SummMargins + startX)
  startX = Round(PageWidth - SummMargins - startX)
  startY = Round(PageHeight - SummMargins)
  ActiveSelectionRange.GetPositionEx cdrCenter, CenterX, CenterY
  Set Symbol = CreateShapeRange
  Set All = CreateShapeRange
  Symbol.AddRange ActiveSelectionRange
  Set PrintLayer = ActiveLayer
  Set CutLayer = ActivePage.Layers.Find("Рез")
  If CutLayer Is Nothing Then
    Set CutLayer = ActivePage.CreateLayer("Рез")
  End If
  Set ContourLayer = ActivePage.Layers.Find("Контур")
  If ContourLayer Is Nothing Then
    Set ContourLayer = ActivePage.CreateLayer("Контур")
  End If
  CutLayer.Printable = False
  CutLayer.Visible = True
  ContourLayer.Printable = False
  ContourLayer.Visible = True
  PrintLayer.Printable = True
  PrintLayer.Visible = True
End Sub

Private Sub Pack(Contour As Shape, Xstep As Double, YStep1 As Double, YStep2 As Double)
Dim HGap As Double, VGap As Double
  If (Height > PageHeight - SummMargins * 2) Or (Width > PageWidth - SummMargins * 2) Then
    MsgBox "Фигура не умещается"
    Contour.Delete
    PackForm.Hide
    ActiveDocument.EndCommandGroup
    Optimization = False
    Refresh
    End
  End If
  Contour.ConvertToCurves
  If VersionMajor > 16 Then
    Contour.curve.AutoReduceNodes
  End If
  With Contour.CreateContour(cdrContourInside, ContourOffset + 1).Separate
    Set Contour = .Item(1)
    Symbol.AddToPowerClip .Item(2)
    Symbol.RemoveAll
    .Item(2).Outline.SetNoOutline
    Symbol.Add .Item(2).ConvertToSymbol("")
    Symbol.MoveToLayer PrintLayer
  End With
  If ContourCheckBox.Value Then
    If RoundRadius > 0 Then
      Contour.Fillet RoundRadius
    End If
    Symbol.Add Contour
  Else
    Contour.Delete
  End If
  VGap = Round(Height - Symbol.SizeHeight) * 0.5
  HGap = Round(Width - Symbol.SizeWidth) * 0.5
  startX = startX + HGap
  endX = endX + HGap
  y = startY - Height + VGap
  endY = 0
  Rows = 0
  Do While y > SumMargins + TopMargin + VGap
    endY = y
    Symbol.Rotate 0
    x = startX - Width
    While x >= endX
      Symbol.SetPositionEx cdrBottomLeft, x, y
      All.AddRange Symbol
      Set Symbol = Symbol.Duplicate(0, 0)
      x = x - Width
    Wend
    y = y - YStep1
    Rows = Rows + 1
    If y < SumMargins + TopMargin + VGap Then
      Exit Do
    End If
    endY = y
    Symbol.Rotate 180
    x = x + Xstep
    If x < endX Then
      x = x + Width
    End If
    While x <= startX - Width
      Symbol.SetPositionEx cdrBottomLeft, x, y
      All.AddRange Symbol
      Set Symbol = Symbol.Duplicate(0, 0)
      x = x + Width
    Wend
    y = y - YStep2
    Rows = Rows + 1
  Loop
  startX = startX - HGap
  endX = endX - HGap
  startY = startY
  endY = endY - VGap
  Symbol.Delete
  PackWidth = startX - endX
  PackHeight = startY - endY
End Sub

Private Sub AddPoint(ByVal x As Long, ByVal y As Long, ByVal typ As Long)
  With CurveElements(n)
    .NodeType = cdrCuspNode
    .ElementType = typ
    .Flags = cdrFlagValid + cdrFlagUser
    .PositionX = x
    .PositionY = y
  End With
  n = n + 1
End Sub

Private Sub Finish(CutPath() As CurveElement)
  Dim curve As New curve, RightBorder As Double, Shape As Shape
  If FastCutCheckBox.Value Then
    curve.PutCurveInfo CutPath, n
    Set Shape = CutLayer.CreateCurve(curve)
    Shape.SetPositionEx cdrBottomLeft, endX, endY
    All.Add Shape
  End If
  endY = endY - MarkMargins
  If MarksCheckBox.Value Then
    ReDim CurveElements(14)
    n = 0
    startY = startY + MarkMargins
    RightBorder = PageWidth - Margins
    AddPoint Margins, endY + 50000, cdrElementStart
    AddPoint Margins, endY, cdrElementLine
    AddPoint Margins + 50000, endY, cdrElementLine
    AddPoint RightBorder - 150000, endY, cdrElementStart
    AddPoint RightBorder - 100000, endY, cdrElementLine
    AddPoint RightBorder - 50000, endY, cdrElementStart
    AddPoint RightBorder, endY, cdrElementLine
    AddPoint RightBorder, endY + 50000, cdrElementLine
    AddPoint RightBorder, startY - 50000, cdrElementStart
    AddPoint RightBorder, startY, cdrElementLine
    AddPoint RightBorder - 50000, startY, cdrElementLine
    AddPoint Margins + 50000, startY, cdrElementStart
    AddPoint Margins, startY, cdrElementLine
    AddPoint Margins, startY - 50000, cdrElementLine
    curve.PutCurveInfo CurveElements, 14
    Set Shape = PrintLayer.CreateCurve(curve)
    All.Add Shape
    With Shape
      .SetPositionEx cdrBottomLeft, Margins, endY
      .Name = "RegMark X00 Y00"
    End With
  End If
  All.OrderReverse
  If TurnAroundCheckBox.Value Then
    All.Rotate 180
  Else
    All.Move 0, Margins - endY
  End If
  PackForm.Hide
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub CircleButton_Click()
  Dim Cols As Long, deltay As Double
  ReadParams
  Width = radius * 2 + Gap
  Height = Width
  init 0
  deltay = Round(Width * 0.866025404)  'sin(PI3/3)
  Pack ContourLayer.CreateEllipse2(CenterX, CenterY, radius, radius), HalfWidth, deltay, deltay
  If FastCutCheckBox.Value Then
    radius = radius + Gap * 0.5
    Cols = Fix((PageWidth - SummMargins * 2) / Width)
    ReDim CurveElements(Rows * Cols * 12 + Rows)
    deltax = radius
    y = startY - radius
    n = 0
    For j = 0 To Rows - 1
      x = startX - radius - (j And 1) * radius
      AddPoint x + radius, y, cdrElementStart
      For i = 1 To Cols
        If (i And 1) = 1 Then
          AddPoint x + radius, y + radius * 0.551785, cdrElementControl
          AddPoint x + radius * 0.551785, y + radius, cdrElementControl
          AddPoint x, y + radius, cdrElementCurve
          AddPoint x - radius * 0.551785, y + radius, cdrElementControl
          AddPoint x - radius, y + radius * 0.551785, cdrElementControl
          AddPoint x - radius, y, cdrElementCurve
        Else
          AddPoint x + radius, y - radius * 0.551785, cdrElementControl
          AddPoint x + radius * 0.551785, y - radius, cdrElementControl
          AddPoint x, y - radius, cdrElementCurve
          AddPoint x - radius * 0.551785, y - radius, cdrElementControl
          AddPoint x - radius, y - radius * 0.551785, cdrElementControl
          AddPoint x - radius, y, cdrElementCurve
        End If
        x = x - Width
        If x < endX + radius Then
          Exit For
        End If
      Next
      For i = i To i + Cols
        x = x + Width
        If x > startX - radius Then
          Exit For
        End If
        If (i And 1) = 0 Then
          AddPoint x - radius, y + radius * 0.551785, cdrElementControl
          AddPoint x - radius * 0.551785, y + radius, cdrElementControl
          AddPoint x, y + radius, cdrElementCurve
          AddPoint x + radius * 0.551785, y + radius, cdrElementControl
          AddPoint x + radius, y + radius * 0.551785, cdrElementControl
          AddPoint x + radius, y, cdrElementCurve
        Else
          AddPoint x - radius, y - radius * 0.551785, cdrElementControl
          AddPoint x - radius * 0.551785, y - radius, cdrElementControl
          AddPoint x, y - radius, cdrElementCurve
          AddPoint x + radius * 0.551785, y - radius, cdrElementControl
          AddPoint x + radius, y - radius * 0.551785, cdrElementControl
          AddPoint x + radius, y, cdrElementCurve
        End If
      Next
      y = y - deltay
    Next
  End If
  Finish CurveElements
End Sub

Private Sub ContourCheckBox_Click()
  ContourLabel.Enabled = ContourCheckBox.Value
  RoundRadius_Edit.Enabled = ContourCheckBox.Value
  ContourOffset_edit.Enabled = ContourCheckBox.Value
End Sub

Private Sub MarksCheckBox_Click()
  MarksLabel.Enabled = MarksCheckBox.Value
  MarkMargins_edit.Enabled = MarksCheckBox.Value
End Sub

Private Sub TriangleButton_Click()
  Dim Cols As Long, StepX As Double, StepY As Double, P1x As Double, P1y As Double, a As Double, b As Double
  ReadParams
  Width = Width - Gap
  Height = Width * 0.866025403784439 + Gap '0.5*Tan(PI/3)
  radius = Width * 0.577350269189626 '1/ Sqr(3)
  Width = (3 * radius + Gap * 4) * 0.577350269189626 '1/ Sqr(3)
  init 1
  Symbol.GetPositionEx cdrBottomLeft, x, y
  CenterY = y + Symbol.SizeWidth * 0.288675135 '0.5*tg(PI/6)
  Pack ContourLayer.CreatePolygon2(CenterX, CenterY, radius, 3), HalfWidth, 0, Height
  If FastCutCheckBox.Value Then
    StepX = PackWidth
    x = startX
    y = startY
    Rows = Fix(PackHeight / Height)
    Cols = Fix(PackWidth / Width)
    ReDim CurveElements((Rows + Cols) * 3)
    n = 0
    For i = 0 To Rows
      StepX = -StepX
      AddPoint x, y, cdrElementStart
      x = x + StepX
      AddPoint x, y, cdrElementLine
      y = y - Height
    Next
    y = y + Height
    
    a = HalfWidth / Height
    b = -Height / HalfWidth
    deltay = PackWidth * 2
    deltax = Int(deltay * a)
    If x = startX Then
      a = -a
      b = -b
      deltax = -deltax
    ElseIf PackWidth Mod Width < 100 Then
      x = x + HalfWidth
      a = -a
      b = -b
      deltax = -deltax
    End If
    n = n - 1
    
    AddPoint x, y, cdrElementLine
    P1x = Round((x - endX) / HalfWidth) * HalfWidth + endX
    P1y = Round((y - endY) / Height) * Height + endY
    Do
      y = y + deltay
      x = x + deltax
      y = Round((y - endY) / Height) * Height + endY
      If y > startY Then
        x = x + (startY - y) * a
        y = startY
        deltay = -deltay
      ElseIf y < endY Then
        x = x + (endY - y) * a
        y = endY
        deltay = -deltay
      End If
      x = Round((x - endX) / HalfWidth) * HalfWidth + endX
      If x > startX Then
        If (y = startY) Or (y = endY) Then
          deltay = -deltay
        End If
        y = y - (startX - x) * b
        x = startX
        deltax = -deltax
      ElseIf x < endX Then
        If (y = startY) Or (y = endY) Then
          deltay = -deltay
        End If
        y = y - (endX - x) * b
        x = endX
        deltax = -deltax
      End If
      x = Round((x - endX) / HalfWidth) * HalfWidth + endX
      y = Round((y - endY) / Height) * Height + endY
      a = -a
      b = -b
      AddPoint x, y, cdrElementLine
      If ((x = P1x) And (y = P1y)) Or ((x = startX) Or (x = endX)) And ((y = startY) Or (y = endY)) Then
        If n > UBound(CurveElements) - Cols Then
          Exit Do
        End If
        x = x + IIf(deltax < 0, Width, -Width)
        P1x = x
        P1y = y
        AddPoint x, y, cdrElementStart
      ElseIf n = UBound(CurveElements) Then
        MsgBox "Что-то пошло не так"
        Exit Do
      End If
    Loop Until False
  End If
  Finish CurveElements
End Sub

Private Sub RectangleButton_Click()
  Dim Cols As Long, StepX As Double, StepY As Double
  ReadParams
  init 2
  StepX = Width - Gap
  StepY = Height - Gap
  Pack ContourLayer.CreateRectangle2(CenterX - StepX * 0.5, CenterY - StepY * 0.5, StepX, StepY), 0, Height, Height
  If FastCutCheckBox.Value Then
    StepX = PackWidth
    x = startX
    y = startY
    Cols = Fix(PackWidth / Width)
    ReDim CurveElements((Rows + Cols) * 2 + 4)
    n = 0
    For i = 0 To Rows
      StepX = -StepX
      AddPoint x, y, cdrElementStart
      x = x + StepX
      AddPoint x, y, cdrElementLine
      y = y - Height
    Next
    StepX = IIf(x = startX, -Width, Width)
    StepY = startY - endY
    y = y + Height
    n = n - 1
    j = n
    For i = 0 To Cols
      AddPoint x, y, cdrElementStart
      y = y + StepY
      AddPoint x, y, cdrElementLine
      x = x + StepX
      StepY = -StepY
    Next
    CurveElements(j).ElementType = cdrElementLine
  End If
  Finish CurveElements
End Sub

Private Sub HexagonButton_Click()
  Dim Cols As Long, deltay As Double, h As Double, w As Double, ofs1 As Double, ofs2 As Double
  ReadParams
  deltay = Height - Gap
  radius = deltay * 0.5
  Width = deltay * 0.866025403784439 + Gap 'Sqr(3)/2
  init 3
  deltay = Round(Width * 0.866025403784439) 'sin(PI3/3)
  Height = Round(deltay * 1.33333333333333 + 0.5)
  Pack ContourLayer.CreatePolygon2(CenterX, CenterY, radius, 6), HalfWidth, deltay, deltay
  If FastCutCheckBox.Value Then
    h = Height * 0.25
    Cols = Fix(PackWidth / Width) + 1
    ReDim CurveElements(Rows * Cols * 6 + Rows)
    w = -HalfWidth
    x = startX
    y = startY - h
    n = 0
    Do While y >= endY
      w = -w
      AddPoint x, y, cdrElementStart
      Do
        x = x - w
        y = y + h
        h = -h
        AddPoint x, y, cdrElementLine
      Loop While x >= endX
      n = n - 1
      x = x + w
      y = y - deltay
      If y < endY Then
        Exit Do
      End If
      w = -w
      AddPoint x, y, cdrElementStart
      Do
        x = x - w
        y = y + h
        h = -h
        AddPoint x, y, cdrElementLine
      Loop While x <= startX
      n = n - 1
      x = x + w
      y = y - deltay + IIf(h < 0, h, 0)
    Loop
    h = Abs(h)
    
    ofs = h
    ofs2 = Height * 1.75
    If (Rows And 1) = 0 Then
      If PackWidth Mod Width < 100 Then
        ofs = Height
      Else
        ofs2 = h
      End If
    End If
    For i = 1 To Cols
      y = endY + ofs
      While y < startY
        AddPoint x, y, cdrElementStart
        y = y + Height * 0.5
        AddPoint x, y, cdrElementLine
        y = y + Height
      Wend
      x = x + w
      If (x < endX) Or (x > startX) Then
        Exit For
      End If
      y = y - ofs2
      While y > endY
        AddPoint x, y, cdrElementStart
        y = y - Height * 0.5
        AddPoint x, y, cdrElementLine
        y = y - Height
      Wend
      x = x + w
      y = y + h
    Next
  End If
  Finish CurveElements
End Sub

Private Sub UserForm_Initialize()
  Dim ctrl As Object
  ActiveDocument.Unit = cdrMillimeter
  ActivePage.GetSize PageWidth, PageHeight
  PageWidth_edit.Value = Round(PageWidth)
  PageHeight_edit.Value = Round(PageHeight)
  n = 0
  For Each ctrl In Controls
    If TypeName(ctrl) = "TextBox" Then
      n = n + 1
      Set Edits(n) = New MyEdit
      Set Edits(n).Edit = ctrl
    End If
  Next ctrl
End Sub
