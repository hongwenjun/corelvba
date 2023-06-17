Attribute VB_Name = "CutLines"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "裁切线"   CutLines  2023.6.9

'// 选中多个物件批量制作四角裁切线
Public Function Batch_CutLines()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  API.BeginOpt
  Bleed = API.GetSet("Bleed")
  Line_len = API.GetSet("Line_len")
  Outline_Width = API.GetSet("Outline_Width")

  '// 定义当前选择物件 分别获得 左右下上中心坐标(x,y)和尺寸信息
  Dim s1 As Shape, OrigSelection As ShapeRange, sr As New ShapeRange
  Set OrigSelection = ActiveSelectionRange

  For Each s1 In OrigSelection
    lx = s1.LeftX:      rx = s1.RightX
    By = s1.BottomY:    ty = s1.TopY
    cx = s1.CenterX:    cy = s1.CenterY
    sw = s1.SizeWidth:  sh = s1.SizeHeight
    
    '//  添加裁切线，分别左下-右下-左上-右上
    Dim s2, s3, s4, s5, s6, s7, s8, s9 As Shape
    Set s2 = ActiveLayer.CreateLineSegment(lx - Bleed, By, lx - (Bleed + Line_len), By)
    Set s3 = ActiveLayer.CreateLineSegment(lx, By - Bleed, lx, By - (Bleed + Line_len))

    Set s4 = ActiveLayer.CreateLineSegment(rx + Bleed, By, rx + (Bleed + Line_len), By)
    Set s5 = ActiveLayer.CreateLineSegment(rx, By - Bleed, rx, By - (Bleed + Line_len))

    Set s6 = ActiveLayer.CreateLineSegment(lx - Bleed, ty, lx - (Bleed + Line_len), ty)
    Set s7 = ActiveLayer.CreateLineSegment(lx, ty + Bleed, lx, ty + (Bleed + Line_len))

    Set s8 = ActiveLayer.CreateLineSegment(rx + Bleed, ty, rx + (Bleed + Line_len), ty)
    Set s9 = ActiveLayer.CreateLineSegment(rx, ty + Bleed, rx, ty + (Bleed + Line_len))

    '// 选中裁切线 群组 设置线宽和注册色
    ActiveDocument.AddToSelection s2, s3, s4, s5, s6, s7, s8, s9
    ActiveSelection.Group
    sr.Add ActiveSelection
  Next s1

  '// 设置线宽和颜色，再选择
   sr.SetOutlineProperties Outline_Width
   sr.SetOutlineProperties Color:=CreateRegistrationColor
   sr.AddToSelection
   
  API.EndOpt
End Function

'// 标注尺寸标记线
Public Function Dimension_MarkLines(Optional ByVal mark As cdrAlignType = cdrAlignTop, Optional ByVal mirror As Boolean = False)
  If 0 = ActiveSelectionRange.Count Then Exit Function
  API.BeginOpt
  Bleed = API.GetSet("Bleed")
  Line_len = API.GetSet("Line_len")
  Outline_Width = API.GetSet("Outline_Width")

  '// 定义当前选择物件 分别获得 左右下上中心坐标(x,y)和尺寸信息
  Dim s As Shape, s1 As Shape, OrigSelection As ShapeRange, sr As New ShapeRange
  Set OrigSelection = ActiveSelectionRange

  For Each s1 In OrigSelection
    lx = s1.LeftX:      rx = s1.RightX
    By = s1.BottomY:    ty = s1.TopY
    
    '//  添加使用 左-上 标注尺寸标记线
    Dim s2, s6, s7, s8, s9 As Shape
    
    If mark = cdrAlignTop Then
      Set s7 = ActiveLayer.CreateLineSegment(lx, ty + Bleed, lx, ty + (Bleed + Line_len))
      Set s9 = ActiveLayer.CreateLineSegment(rx, ty + Bleed, rx, ty + (Bleed + Line_len))
      sr.Add s7: sr.Add s9
    Else
      Set s2 = ActiveLayer.CreateLineSegment(lx - Bleed, By, lx - (Bleed + Line_len), By)
      Set s6 = ActiveLayer.CreateLineSegment(lx - Bleed, ty, lx - (Bleed + Line_len), ty)
      sr.Add s2: sr.Add s6
    End If
  Next s1

  '// 获得页面中心点 x,y
'  px = ActiveDocument.Pages.First.CenterX
'  py = ActiveDocument.Pages.First.CenterY
  '// 物件范围边界
  px = OrigSelection.LeftX
  py = OrigSelection.TopY
  mpx = OrigSelection.RightX
  mpy = OrigSelection.BottomY
  
  '// 页面边缘对齐
  For Each s In sr
    s.Name = "DMKLine"
    If mark = cdrAlignTop Then
      s.TopY = py + Line_len + Bleed
    Else
      s.LeftX = px - Line_len - Bleed
    End If
  Next s
  
  '// 简单删除重复
  RemoveDuplicates sr
  
  '// 设置线宽和颜色，再选择
   sr.SetOutlineProperties Outline_Width
   sr.SetOutlineProperties Color:=CreateCMYKColor(80, 40, 0, 20)
   sr.AddToSelection
   
   If mirror Then
    If mark = cdrAlignTop Then
      sr.BottomY = mpy - Line_len - Bleed
    Else
      sr.RightX = mpx + Line_len + Bleed
    End If
   End If
   
  API.EndOpt
End Function

 '// 简单删除重复线和物件算法算法
Public Function RemoveDuplicates(sr As ShapeRange)
  Dim s As Shape, cnt As Integer, rms As New ShapeRange
  cnt = 1
  
  #If VBA7 Then
     sr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
  #Else
    ' X4 不支持 ShapeRange.sort
  #End If

  For Each s In sr
    If cnt > 1 Then
      If Check_duplicate(sr(cnt - 1), sr(cnt)) Then rms.Add sr(cnt)
    End If
    cnt = cnt + 1
  Next s
  
  rms.Delete
End Function

 '// 检查重复算法
Private Function Check_duplicate(s1 As Shape, s2 As Shape) As Boolean
  Check_duplicate = False
  Jitter = 0.3
  X = Abs(s1.CenterX - s2.CenterX)
  Y = Abs(s1.CenterY - s2.CenterY)
  w = Abs(s1.SizeWidth - s2.SizeWidth)
  h = Abs(s1.SizeHeight - s2.SizeHeight)
  If X < Jitter And Y < Jitter And w < Jitter And h < Jitter Then
    Check_duplicate = True
  End If
End Function


'// 单线条转裁切线 - 放置到页面四边
Public Function SelectLine_to_Cropline()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  '// 代码运行时关闭窗口刷新
  Application.Optimization = True
  ActiveDocument.Unit = cdrMillimeter
  
  ActiveDocument.BeginCommandGroup  '一步撤消'
  
  '// 获得页面中心点 x,y
  px = ActiveDocument.Pages.First.CenterX
  py = ActiveDocument.Pages.First.CenterY
  Bleed = API.GetSet("Bleed")
  Line_len = API.GetSet("Line_len")
  Outline_Width = API.GetSet("Outline_Width")
  
  Dim s As Shape
  Dim line As Shape
  
  '// 遍历选择的线条
  For Each s In ActiveSelection.Shapes
  
    lx = s.LeftX
    rx = s.RightX
    By = s.BottomY
    ty = s.TopY
    
    cx = s.CenterX
    cy = s.CenterY
    sw = s.SizeWidth
    sh = s.SizeHeight
   
   '// 判断横线(高度小于宽度)，在页面左边还是右边
   If sh <= sw Then
    s.Delete
    If cx < px Then
        Set line = ActiveLayer.CreateLineSegment(0, cy, 0 + Line_len, cy)
    Else
        Set line = ActiveLayer.CreateLineSegment(px * 2, cy, px * 2 - Line_len, cy)
    End If
   End If
 
   '// 判断竖线(高度大于宽度)，在页面下边还是上边
   If sh > sw Then
    s.Delete
    If cy < py Then
        Set line = ActiveLayer.CreateLineSegment(cx, 0, cx, 0 + Line_len)
    Else
        Set line = ActiveLayer.CreateLineSegment(cx, py * 2, cx, py * 2 - Line_len)
    End If
   End If

    line.Outline.SetProperties Outline_Width
    line.Outline.SetProperties Color:=CreateRegistrationColor
  Next s
  
  ActiveDocument.EndCommandGroup
  '// 代码操作结束恢复窗口刷新
  Application.Optimization = False
  ActiveWindow.Refresh
  Application.Refresh
End Function


'// 拼版裁切线
Public Function Draw_Lines()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  API.BeginOpt
  
  Dim OrigSelection As ShapeRange, sr As ShapeRange
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
    By = s1.BottomY: ty = s1.TopY
    cx = s1.CenterX: cy = s1.CenterY
    
    '// 范围边界物件判断
    If Abs(set_lx - lx) < radius Or Abs(set_rx - rx) < radius Or Abs(set_by - By) _
      < radius Or Abs(set_ty - ty) < radius Then
      
      arr = Array(lx, By, rx, By, lx, ty, rx, ty)  '// 物件左下-右下-左上-右上 四个顶点坐标数组
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
  
  '// 使用CQL 颜色标志查
  Set sr = ActivePage.Shapes.FindShapes(Query:="@colors.find(RGB(26, 22, 35))")
  
  '// 简单删除重复
  RemoveDuplicates sr
  
  '// 设置线宽和颜色，再选择
   sr.SetOutlineProperties Outline_Width, Color:=CreateRegistrationColor
   sr.Group
   sr.AddRange OrigSelection
   sr.AddToSelection

  API.EndOpt
End Function

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

Private Function set_line_color(line As Shape)
   '// 设置轮廓线注册色
  line.Outline.SetProperties Color:=CreateRGBColor(26, 22, 35)
End Function


