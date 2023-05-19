Attribute VB_Name = "Container"
' ① 标记容器盒子
Public Function SetBoxName()
  Dim box As ShapeRange, s As Shape
  Set box = ActiveSelectionRange
  
  Application.Optimization = True
  ' 设置物件名字，以供CQL查询
  For Each s In box
    s.Name = "Container"
  Next s
  
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
  MsgBox "标记容器盒子" & vbNewLine & "名字: Container"
  
End Function


' ② 删除容器盒子边界外面的物件    ③④
Public Function Remove_OutsideBox()
  Dim s As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim rmsr As New ShapeRange
  Dim x As Double, y As Double
  
  Set ssr = ActiveSelectionRange
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    x = s.CenterX: y = s.CenterY
    If box(1).IsOnShape(x, y) = cdrOutsideShape Then rmsr.Add s
  Next s

  rmsr.Delete
End Function


Public Function Remove_OnMargin()
  Dim s As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim rmsr As New ShapeRange
  Dim x As Double, y As Double
  
  Set ssr = ActiveSelectionRange
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    x = s.CenterX: y = s.CenterY
    If box(1).IsOnShape(x, y) = cdrOnMarginOfShape Then rmsr.Add s
  Next s

  rmsr.Delete
End Function


Public Function Select_OutsideBox()
  Dim s As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim SelSr As New ShapeRange
  Dim x As Double, y As Double, radius
  
  Set ssr = ActiveSelectionRange
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    x = s.CenterX: y = s.CenterY
    radius = s.SizeWidth / 2
    If box(1).IsOnShape(x, y, radius) = cdrOutsideShape Then SelSr.Add s
  Next s
  
  ActiveDocument.ClearSelection
  SelSr.AddToSelection

End Function


Public Function Select_OnMargin()
  Dim s As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim SelSr As New ShapeRange
  Dim x As Double, y As Double, radius
  
  Set ssr = ActiveSelectionRange
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    x = s.CenterX: y = s.CenterY
    radius = s.SizeWidth / 2
    If box(1).IsOnShape(x, y, radius) = cdrOnMarginOfShape Then SelSr.Add s
  Next s
  
  ActiveDocument.ClearSelection
  SelSr.AddToSelection

End Function



' 图片批量置入容器
Public Sub Batch_ToPowerClip()
  ActiveDocument.BeginCommandGroup ' 一键撤销返回
  Dim s As Shape, ssr As ShapeRange, box As ShapeRange
  
  ' 标记容器，请酌情取消注释
  ' Set box = ActiveSelectionRange
  ' For Each s In box
  '   If s.Type <> cdrBitmapShape Then s.Name = "Container"
  ' Next s
  
  Set ssr = Smart_Group(0.5) ' 智能群组容差 0.5mm
  
  Application.Optimization = True
  For Each s In ssr
    Image_ToPowerClip s
  Next s
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:  Application.Refresh
End Sub


' 图片置入容器，基本函数
Public Function Image_ToPowerClip(arg As Shape)
  Dim box As ShapeRange
  Dim ssr As New ShapeRange, rmsr As New ShapeRange
  Set ssr = arg.UngroupEx
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ssr.AddToPowerClip box(1), 0

End Function

Private Function Smart_Group(Optional ByVal tr As Double = 0) As ShapeRange
If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  Application.Optimization = True
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
      Set s = ActiveLayer.CreateRectangle2(x - tr, y - tr, w + 2 * tr, h + 2 * tr)
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

  '// 矩形边界智能群组, retsr 返回群组 和 删除矩形s
  Dim retsr As New ShapeRange, rmsr As New ShapeRange
  For Each s In brk1
    Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, False)
    s.Delete
    retsr.Add sh.Shapes.All.group
  Next

  Set Smart_Group = retsr
  
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
Exit Function

ErrorHandler:
  Application.Optimization = False
  MsgBox "请先选择一些物件来确定群组范围!"
  On Error Resume Next

End Function
