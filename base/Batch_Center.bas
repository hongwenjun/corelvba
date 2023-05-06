Attribute VB_Name = "Batch_Center"
Private Function Smart_Group() As ShapeRange
If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  ActiveDocument.ReferencePoint = cdrBottomLeft
  ActiveDocument.Unit = cdrMillimeter
  
  Dim OrigSelection As ShapeRange, sr As New ShapeRange
  Dim s1 As Shape, sh As Shape, s As Shape
  Dim x As Double, y As Double, w As Double, h As Double
  Dim eff1 As Effect
  
  Set OrigSelection = ActiveSelectionRange

  '// 遍历物件画矩形
  For Each sh In OrigSelection
    sh.GetBoundingBox x, y, w, h
    If w * h > 4 Then
      Set s = ActiveLayer.CreateRectangle2(x, y, w, h)
      sr.Add s

    '// 轴线 创建轮廓处理
    ElseIf w * h < 0.3 Then
    ' Debug.Print w * h
      Set eff1 = sh.CreateContour(cdrContourOutside, 0.5, 1, cdrDirectFountainFillBlend, CreateRGBColor(26, 22, 35), CreateRGBColor(26, 22, 35), CreateRGBColor(26, 22, 35), 0, 0, cdrContourSquareCap, cdrContourCornerMiteredOffsetBevel, 15#)
      eff1.Separate
    End If
  Next sh

  '// 查找轴线轮廓
  ActivePage.Shapes.FindShapes(Query:="@Outline.Color=RGB(26, 22, 35)").CreateSelection
  ActivePage.Shapes.FindShapes(Query:="@fill.Color=RGB(26, 22, 35)").AddToSelection
  For Each sh In ActiveSelection.Shapes
     sr.Add sh
  Next sh

  '// 新矩形寻找边界，散开，删除刚才画的新矩形
  Set s1 = sr.CustomCommand("Boundary", "CreateBoundary")
  Set brk1 = s1.BreakApartEx
  sr.Delete

  '// 矩形边界智能群组，删除矩形
  Dim retsr As New ShapeRange
  For Each s In brk1
    Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, False)
    retsr.Add sh.Shapes.All.Group
    s.Delete
  Next

  Set Smart_Group = retsr

  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
Exit Function

ErrorHandler:
  Application.Optimization = False
  MsgBox "请先选择一些物件来确定群组范围!"
  On Error Resume Next

End Function


' 这个子程序遍历对象，调用解散物件和居中
Public Sub start_Center()
    Dim s As Shape, ssr As ShapeRange
    Set ssr = Smart_Group
    For Each s In ssr
      Ungroup_Center s
    Next s
End Sub


' 以下函数，解散物件，以面积排序居中
Private Function Ungroup_Center(os As Shape)
    Set grp = os.UngroupEx
    grp.Sort "@shape1.Width * @shape1.Height> @shape2.Width * @shape2.Height"
    cx = grp(1).CenterX
    cy = grp(1).CenterY
    For Each s In grp
      s.CenterX = cx
      s.CenterY = cy
    Next s
End Function
