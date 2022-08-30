Attribute VB_Name = "拼版裁切线"
Type Coordinate
  X As Double
  Y As Double
End Type

Sub Cut_lines()
  If 0 = ActiveSelectionRange.Count Then Exit Sub
  '// 代码运行时关闭窗口刷新
  Application.Optimization = True
  ActiveDocument.BeginCommandGroup  '一步撤消'
  ActiveDocument.Unit = cdrMillimeter
  Dim OrigSelection As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  
  Dim s1 As Shape, sbd As Shape
  Dim dot As Coordinate
  Dim arr As Variant, border As Variant
  
  ' 当前选择物件的范围边界
  set_lx = OrigSelection.LeftX:   set_rx = OrigSelection.RightX
  set_by = OrigSelection.BottomY: set_ty = OrigSelection.TopY
  set_cx = OrigSelection.CenterX: set_cy = OrigSelection.CenterY
  radius = 8
  Bleed = API.GetSet("Bleed")
  Line_len = API.GetSet("Line_len")
  Outline_Width = API.GetSet("Outline_Width")
  border = Array(set_lx, set_rx, set_by, set_ty, set_cx, set_cy, radius, Bleed, Line_len)
  
  ' 创建边界矩形，用来添加角线
  Set sbd = ActiveLayer.CreateRectangle(set_lx, set_by, set_rx, set_ty)
  OrigSelection.Add sbd
  
  For Each Target In OrigSelection
    Set s1 = Target
    lx = s1.LeftX:   rx = s1.RightX
    by = s1.BottomY: ty = s1.TopY
    cx = s1.CenterX: cy = s1.CenterY
    
    '// 范围边界物件判断
    If Abs(set_lx - lx) < radius Or Abs(set_rx - rx) < radius Or Abs(set_by - by) _
      < radius Or Abs(set_ty - ty) < radius Then
      
      arr = Array(lx, by, rx, by, lx, ty, rx, ty)  '// 物件左下-右下-左上-右上 四个顶点坐标数组
      For i = 0 To 3
        dot.X = arr(2 * i)
        dot.Y = arr(2 * i + 1)
        
        '// 范围边界坐标点判断
        If Abs(set_lx - dot.X) < radius Or Abs(set_rx - dot.X) < radius _
              Or Abs(set_by - dot.Y) < radius Or Abs(set_ty - dot.Y) < radius Then

            draw_line dot, border  '// 以坐标点和范围边界画裁切线
        End If
      Next i
    End If
  Next Target
  
  sbd.Delete  '删除边界矩形
  
  '// 使用CQL 颜色标志查找，然后群组统一设置线宽和注册色
  ActivePage.Shapes.FindShapes(Query:="@colors.find(RGB(26, 22, 35))").CreateSelection
  ActiveSelection.Group
  ActiveSelection.Outline.SetProperties Outline_Width, Color:=CreateRegistrationColor
  
  ActiveDocument.EndCommandGroup
  '// 代码操作结束恢复窗口刷新
  Application.Optimization = False
  ActiveWindow.Refresh
  Application.Refresh
End Sub

'范围边界 border = Array(set_lx, set_rx, set_by, set_ty, set_cx, set_cy, radius, Bleed, Line_len)
Private Function draw_line(dot As Coordinate, border As Variant)
  radius = border(6): Bleed = border(7):  Line_len = border(8)
  Dim line As Shape

  If Abs(dot.Y - border(3)) < radius Then
    Set line = ActiveLayer.CreateLineSegment(dot.X, border(3) + Bleed, dot.X, border(3) + (Line_len + Bleed))
    set_line_color line
  ElseIf Abs(dot.Y - border(2)) < radius Then
    Set line = ActiveLayer.CreateLineSegment(dot.X, border(2) - Bleed, dot.X, border(2) - (Line_len + Bleed))
    set_line_color line
  End If
  
  If Abs(dot.X - border(1)) < radius Then
    Set line = ActiveLayer.CreateLineSegment(border(1) + Bleed, dot.Y, border(1) + (Line_len + Bleed), dot.Y)
    set_line_color line
  ElseIf Abs(dot.X - border(0)) < radius Then
    Set line = ActiveLayer.CreateLineSegment(border(0) - Bleed, dot.Y, border(0) - (Line_len + Bleed), dot.Y)
    set_line_color line
  End If

End Function

'// 旧版本
Private Function draw_line_按点基准(dot As Coordinate, border As Variant)
  Bleed = 2:  Line_len = 3:  radius = border(6)
  Dim line As Shape

  If Abs(dot.Y - border(3)) < radius Then
    Set line = ActiveLayer.CreateLineSegment(dot.X, dot.Y + Bleed, dot.X, dot.Y + (Line_len + Bleed))
    set_line_color line
  ElseIf Abs(dot.Y - border(2)) < radius Then
    Set line = ActiveLayer.CreateLineSegment(dot.X, dot.Y - Bleed, dot.X, dot.Y - (Line_len + Bleed))
    set_line_color line
  End If
  
  If Abs(dot.X - border(1)) < radius Then
    Set line = ActiveLayer.CreateLineSegment(dot.X + Bleed, dot.Y, dot.X + (Line_len + Bleed), dot.Y)
    set_line_color line
  ElseIf Abs(dot.X - border(0)) < radius Then
    Set line = ActiveLayer.CreateLineSegment(dot.X - Bleed, dot.Y, dot.X - (Line_len + Bleed), dot.Y)
    set_line_color line
  End If

End Function

Private Function set_line_color(line As Shape)
   '// 设置轮廓线注册色
  line.Outline.SetProperties Color:=CreateRGBColor(26, 22, 35)
End Function

'// CorelDRAW 物件排列拼版简单代码
Sub arrange()
  On Error GoTo ErrorHandler
  ActiveDocument.Unit = cdrMillimeter
  row = 3     ' 拼版 3 x 4
  List = 4
  sp = 0       '间隔 0mm

  Dim Str, arr, n
  Str = API.GetClipBoardString

  ' 替换 mm x * 换行 TAB 为空格
  Str = VBA.replace(Str, "mm", " ")
  Str = VBA.replace(Str, "x", " ")
  Str = VBA.replace(Str, "X", " ")
  Str = VBA.replace(Str, "*", " ")
  Str = VBA.replace(Str, Chr(13), " ")
  Str = VBA.replace(Str, Chr(9), " ")
  
  Do While InStr(Str, "  ")    '多个空格换成一个空格
      Str = VBA.replace(Str, "  ", " ")
  Loop
  
  arr = Split(Str)

  Dim s1 As Shape
  Dim X As Double, Y As Double
  
  If 0 = ActiveSelectionRange.Count Then
    X = Val(arr(0)):    Y = Val(arr(1))
    row = Int(ActiveDocument.Pages.First.SizeWidth / X)
    List = Int(ActiveDocument.Pages.First.SizeHeight / Y)

    If UBound(arr) > 2 Then
    row = Val(arr(2)):  List = Val(arr(3))
      If row * List > 800 Then
        GoTo ErrorHandler
      ElseIf UBound(arr) > 3 Then
          sp = Val(arr(4))       '间隔
      End If
    End If
     
    '// 建立矩形 Width  x Height 单位 mm
    Set s1 = ActiveLayer.CreateRectangle(0, 0, X, Y)
    
    '// 填充颜色无，轮廓颜色 K100，线条粗细0.3mm
    s1.Fill.ApplyNoFill
    s1.Outline.SetProperties 0.3, OutlineStyles(0), CreateCMYKColor(0, 100, 0, 0), ArrowHeads(0), _
      ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#

  '// 如果当前选择物件，按当前物件拼版
  ElseIf 1 = ActiveSelectionRange.Count Then
    Set s1 = ActiveSelection
    X = s1.SizeWidth:    Y = s1.SizeHeight
    row = Int(ActiveDocument.Pages.First.SizeWidth / X)
    List = Int(ActiveDocument.Pages.First.SizeHeight / Y)
  End If
  
  sw = X:  sh = Y

  '// StepAndRepeat 方法在范围内创建多个形状副本
  Dim dup1 As ShapeRange
  Set dup1 = s1.StepAndRepeat(row - 1, sw + sp, 0#)
  Dim dup2 As ShapeRange
  Set dup2 = ActiveDocument.CreateShapeRangeFromArray(dup1, s1).StepAndRepeat(List - 1, 0#, (sh + sp))
       
  Exit Sub
ErrorHandler:
  Speak_Msg "记事本输入数字,示例: 50x50 4x3 ,复制到剪贴板再运行工具!"
  MsgBox "记事本输入数字,示例: 50x50 4x3 ,复制到剪贴板再运行工具!"
  On Error Resume Next
End Sub


