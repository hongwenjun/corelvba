Attribute VB_Name = "SmartGroup"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "智能群组"   SmartGroup  2023.6.11

Sub 剪贴板物件替换()
  Replace_UI.Show 0
End Sub

Public Sub 智能群组(Optional ByVal tr As Double = 0)
  If 0 = ActiveSelectionRange.Count Then Exit Sub
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  
  ActiveDocument.ReferencePoint = cdrBottomLeft
  ActiveDocument.Unit = cdrMillimeter
  
  Dim OrigSelection As ShapeRange, sr As New ShapeRange
  Dim s1 As Shape, sh As Shape, s As Shape
  Dim X As Double, Y As Double, w As Double, h As Double
  Dim eff1 As Effect
  
  Set OrigSelection = ActiveSelectionRange

  '// 遍历物件画矩形
  For Each sh In OrigSelection
    sh.GetBoundingBox X, Y, w, h
    If w * h > 4 Then
      Set s = ActiveLayer.CreateRectangle2(X - tr, Y - tr, w + 2 * tr, h + 2 * tr)
      sr.Add s

    '// 轴线 创建轮廓处理
    ElseIf w * h < 0.3 Then
    ' Debug.Print w * h
      Set eff1 = sh.CreateContour(cdrContourOutside, 0.5, 1, cdrDirectFountainFillBlend, _
          CreateRGBColor(26, 22, 35), CreateRGBColor(26, 22, 35), CreateRGBColor(26, 22, 35), 0, 0, cdrContourSquareCap, cdrContourCornerMiteredOffsetBevel, 15#)
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
  For Each s In brk1
    Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, False)
    sh.Shapes.all.group
    s.Delete
  Next

  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:   Application.Refresh
Exit Sub

ErrorHandler:
  Application.Optimization = False
  MsgBox "请先选择一些物件来确定群组范围!"
  On Error Resume Next

End Sub

' 智能群组_V1 第一版，储备示例代码
Function 智能群组_V1()
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  ActiveDocument.Unit = cdrMillimeter
  
  Dim OrigSelection As ShapeRange, brk1 As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  Dim s1 As Shape, sh As Shape, s As Shape
  
  Set s1 = OrigSelection.CustomCommand("Boundary", "CreateBoundary")
' s1.Outline.SetProperties Color:=CreateRGBColor(26, 22, 35)
  Set brk1 = s1.BreakApartEx

  For Each s In brk1
    If s.SizeHeight > 10 Then
      Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, False)
      sh.Shapes.all.group
    End If
    s.Delete
  Next
  
' ActiveDocument.ClearSelection
' ActivePage.Shapes.FindShapes(Query:="@colors.find(RGB(26, 22, 35))").CreateSelections

  '// 代码操作结束恢复窗口刷新
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
Exit Function
ErrorHandler:
  Application.Optimization = False
  MsgBox "请先选择一些物件来确定群组范围!"
  On Error Resume Next
End Function

