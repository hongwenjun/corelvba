Type Coordinate
    x As Double
    y As Double
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
  radius = 8:  border = Array(set_lx, set_rx, set_by, set_ty, set_cx, set_cy, radius)
  
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
        dot.x = arr(2 * i)
        dot.y = arr(2 * i + 1)
        
        '// 范围边界坐标点判断
        If Abs(set_lx - dot.x) < radius Or Abs(set_rx - dot.x) < radius _
              Or Abs(set_by - dot.y) < radius Or Abs(set_ty - dot.y) < radius Then

            draw_line dot, border  '// 以坐标点和范围边界画裁切线
        End If
      Next i
    End If
  Next Target
  
  sbd.Delete  '删除边界矩形
  
  '// 使用CQL 颜色标志查找，然后群组统一设置线宽和注册色
  ActivePage.Shapes.FindShapes(Query:="@colors.find(RGB(26, 22, 35))").CreateSelection
  ActiveSelection.Group
  ActiveSelection.Outline.SetProperties 0.1, Color:=CreateRegistrationColor
  
  ActiveDocument.EndCommandGroup
  '// 代码操作结束恢复窗口刷新
  Application.Optimization = False
  ActiveWindow.Refresh
  Application.Refresh
End Sub

'范围边界 border = Array(set_lx, set_rx, set_by, set_ty, set_cx, set_cy, radius)
Private Function draw_line(dot As Coordinate, border As Variant)
    Bleed = 2:  line_len = 3:  radius = border(6)
    Dim line As Shape

    If Abs(dot.y - border(3)) < radius Then
        Set line = ActiveLayer.CreateLineSegment(dot.x, border(3) + Bleed, dot.x, border(3) + (line_len + Bleed))
        set_line_color line
    ElseIf Abs(dot.y - border(2)) < radius Then
        Set line = ActiveLayer.CreateLineSegment(dot.x, border(2) - Bleed, dot.x, border(2) - (line_len + Bleed))
        set_line_color line
    End If
    
    If Abs(dot.x - border(1)) < radius Then
        Set line = ActiveLayer.CreateLineSegment(border(1) + Bleed, dot.y, border(1) + (line_len + Bleed), dot.y)
        set_line_color line
    ElseIf Abs(dot.x - border(0)) < radius Then
        Set line = ActiveLayer.CreateLineSegment(border(0) - Bleed, dot.y, border(0) - (line_len + Bleed), dot.y)
        set_line_color line
    End If

End Function

'// 旧版本
Private Function draw_line_按点基准(dot As Coordinate, border As Variant)
    Bleed = 2:  line_len = 3:  radius = border(6)
    Dim line As Shape

    If Abs(dot.y - border(3)) < radius Then
        Set line = ActiveLayer.CreateLineSegment(dot.x, dot.y + Bleed, dot.x, dot.y + (line_len + Bleed))
        set_line_color line
    ElseIf Abs(dot.y - border(2)) < radius Then
        Set line = ActiveLayer.CreateLineSegment(dot.x, dot.y - Bleed, dot.x, dot.y - (line_len + Bleed))
        set_line_color line
    End If
    
    If Abs(dot.x - border(1)) < radius Then
        Set line = ActiveLayer.CreateLineSegment(dot.x + Bleed, dot.y, dot.x + (line_len + Bleed), dot.y)
        set_line_color line
    ElseIf Abs(dot.x - border(0)) < radius Then
        Set line = ActiveLayer.CreateLineSegment(dot.x - Bleed, dot.y, dot.x - (line_len + Bleed), dot.y)
        set_line_color line
    End If

End Function

Private Function set_line_color(line As Shape)
    '// 设置线宽和注册色
   line.Outline.SetProperties Color:=CreateRGBColor(26, 22, 35)
End Function
