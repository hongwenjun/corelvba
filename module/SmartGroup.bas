Attribute VB_Name = "SmartGroup"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "智能群组"   SmartGroup  2023.6.30

Public Function Smart_Group(Optional ByVal tr As Double = 0) As ShapeRange
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  API.BeginOpt

  Dim OrigSelection As ShapeRange, sr As New ShapeRange
  Dim s1 As Shape, sh As Shape, s As Shape
  Dim x As Double, Y As Double, w As Double, h As Double
  Dim eff1 As Effect
  
  Set OrigSelection = ActiveSelectionRange

  '// 遍历物件画矩形
  For Each sh In OrigSelection
    sh.GetBoundingBox x, Y, w, h
    If w * h > 4 Then
      Set s = ActiveLayer.CreateRectangle2(x - tr, Y - tr, w + 2 * tr, h + 2 * tr)
      sr.Add s

    '// 轴线 创建轮廓处理
    ElseIf w * h < 0.3 Then
    ' Debug.Print w * h
      Set eff1 = sh.CreateContour(cdrContourOutside, 0.5, 1, cdrDirectFountainFillBlend, CreateRGBColor(26, 22, 35), _
              CreateRGBColor(26, 22, 35), CreateRGBColor(26, 22, 35), 0, 0, cdrContourSquareCap, cdrContourCornerMiteredOffsetBevel, 15#)
      eff1.Separate
    End If
  Next sh

  '// 查找轴线轮廓
  sr.AddRange ActivePage.Shapes.FindShapes(Query:="@Outline.Color=RGB(26, 22, 35)")
  sr.AddRange ActivePage.Shapes.FindShapes(Query:="@fill.Color=RGB(26, 22, 35)")

  '// 新矩形寻找边界，散开，删除刚才画的新矩形
  Dim brk1 As ShapeRange
  Set s1 = sr.CustomCommand("Boundary", "CreateBoundary")
  Set brk1 = s1.BreakApartEx
  sr.Delete

  '// 矩形边界智能群组, RetSR 返回群组 和 删除矩形s
  Dim RetSR As New ShapeRange
  For Each s In brk1
    Set sr = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, False).Shapes.all
    sr.DeleteItem sr.IndexOf(s)
    If sr.Count > 0 Then RetSR.Add sr.Group
  Next s
  
  '// 智能群组返回和选择
  Set Smart_Group = RetSR
  RetSR.CreateSelection
  
ErrorHandler:
  API.EndOpt
End Function

'// 智能群组 原理版
Private Function Smart_Group_ABC()
  ActiveDocument.Unit = cdrMillimeter
  
  Dim OrigSelection As ShapeRange, brk1 As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  Dim s1 As Shape, sh As Shape, s As Shape
  
  Set s1 = OrigSelection.CustomCommand("Boundary", "CreateBoundary")
  Set brk1 = s1.BreakApartEx

  For Each s In brk1
    If s.SizeHeight > 10 Then
      Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, False)
      sh.Shapes.all.Group
    End If
    s.Delete
  Next
End Function

