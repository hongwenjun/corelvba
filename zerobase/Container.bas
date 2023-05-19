Attribute VB_Name = "Container"
' ① 标记容器盒子
Public Function SetBoxName()
  API.BeginOpt "标记容器盒子"
  
  Dim box As ShapeRange, S As Shape
  Set box = ActiveSelectionRange
  
  ' 设置物件名字，以供CQL查询
  For Each S In box
    S.Name = "Container"
  Next S
  
  API.EndOpt
  MsgBox "标记容器盒子" & vbNewLine & "名字: Container"
End Function

' 图片批量置入容器
Public Sub Batch_ToPowerClip()
  API.BeginOpt "批量置入容器"
  Dim S As Shape, ssr As ShapeRange, box As ShapeRange
  Set ssr = Smart_Group(0.5) ' 智能群组容差 0.5mm
  
  For Each S In ssr
    Image_ToPowerClip S
  Next S

  API.EndOpt
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
  
  box.SetOutlineProperties Width:=0, Color:=Nothing
  ssr.AddToPowerClip box(1), 0
  box(1).Name = "powerclip_ok"

End Function

' 图片OneKey置入容器
Public Sub OneKey_ToPowerClip()
  API.BeginOpt "图片OneKey置入容器"
  Dim S As Shape, ssr As ShapeRange, box As ShapeRange
  
  ' 标记容器，设置透明
  Set box = ActiveSelectionRange
  For Each S In box
    If S.Type <> cdrBitmapShape Then S.Name = "Container"
  Next S
  
  Set ssr = Smart_Group(0.5) ' 智能群组容差 0.5mm
  
  Application.Optimization = True
  For Each S In ssr
    Image_ToPowerClip S
  Next S

  API.EndOpt
End Sub

' ② 删除容器盒子边界外面的物件    ③④
Public Function Remove_OutsideBox(radius As Double)
  API.BeginOpt "删除容器盒子边界外面的物"
  On Error GoTo ErrorHandler
  Dim S As Shape, bc As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim rmsr As New ShapeRange
  Dim x As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then GoTo ErrorHandler
  Set bc = box(1).Duplicate(0, 0)
  If bc.Type = cdrTextShape Then bc.ConvertToCurves
  
  For Each S In ssr
    x = S.CenterX: Y = S.CenterY
    If bc.IsOnShape(x, Y, radius) = cdrOutsideShape Then rmsr.Add S
  Next S
  
  rmsr.Add bc: rmsr.Delete: API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
  On Error Resume Next

End Function

Public Function Select_OutsideBox(radius As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt "选择容器外面对象"
  Dim S As Shape, bc As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim SelSr As New ShapeRange
  Dim x As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then GoTo ErrorHandler
  Set bc = box(1).Duplicate(0, 0)
  If bc.Type = cdrTextShape Then bc.ConvertToCurves
  
  ActiveDocument.unit = cdrMillimeter
  For Each S In ssr
    x = S.CenterX: Y = S.CenterY
    If bc.IsOnShape(x, Y, S.SizeWidth / 2 * radius) = cdrOutsideShape Then SelSr.Add S
  Next S
  
  ActiveDocument.ClearSelection
  bc.Delete: SelSr.AddToSelection: API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
End Function

Public Function Select_by_BlendGroup(radius As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt "使用调和群组选择"
  Dim S As Shape, bc As Shape
  Dim ssr As ShapeRange, box As ShapeRange, gp As ShapeRange
  Dim SelSr As New ShapeRange
  Dim x As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then GoTo ErrorHandler
  Set gp = box.Duplicate(0, 0).UngroupAllEx
  Set bc = gp.BreakApartEx.UngroupAllEx.Combine

  ActiveDocument.unit = cdrMillimeter
  For Each S In ssr
    x = S.CenterX: Y = S.CenterY
    If bc.IsOnShape(x, Y, S.SizeWidth / 2 * radius) = cdrOnMarginOfShape Then SelSr.Add S
  Next S
  
  ActiveDocument.ClearSelection
  bc.Delete: SelSr.AddToSelection: API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function

Public Function Select_OnMargin(radius As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt "选择容器边界对象"
  Dim S As Shape, bc As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim SelSr As New ShapeRange
  Dim x As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  ' CQL查找容器盒物件
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then GoTo ErrorHandler
  Set bc = box(1).Duplicate(0, 0)
  If bc.Type = cdrTextShape Then bc.ConvertToCurves  ' 如果是文本转曲

  
  ActiveDocument.unit = cdrMillimeter
  For Each S In ssr
    x = S.CenterX: Y = S.CenterY
    If bc.IsOnShape(x, Y, S.SizeWidth / 2 * radius) = cdrOnMarginOfShape Then SelSr.Add S
  Next S
  
  ActiveDocument.ClearSelection
  bc.Delete: SelSr.AddToSelection: API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
  
End Function


Private Function Smart_Group(Optional ByVal tr As Double = 0) As ShapeRange
If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  Application.Optimization = True
  ActiveDocument.ReferencePoint = cdrBottomLeft
  ActiveDocument.unit = cdrMillimeter
  
  Dim OrigSelection As ShapeRange, sr As New ShapeRange
  Dim s1 As Shape, sh As Shape, S As Shape
  Dim x As Double, Y As Double, w As Double, h As Double
  Dim eff1 As Effect
  
  Set OrigSelection = ActiveSelectionRange

  '// 遍历物件画矩形
  For Each sh In OrigSelection
    sh.GetBoundingBox x, Y, w, h
    If w * h > 4 Then
      Set S = ActiveLayer.CreateRectangle2(x - tr, Y - tr, w + 2 * tr, h + 2 * tr)
      sr.Add S

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
  For Each S In brk1
    Set sh = ActivePage.SelectShapesFromRectangle(S.LeftX, S.TopY, S.RightX, S.BottomY, False)
    S.Delete
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

