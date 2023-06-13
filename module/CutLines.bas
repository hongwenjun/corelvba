Attribute VB_Name = "裁切线"
' Attribute VB_Name = "裁切线"
Sub start()
If 0 = ActiveSelectionRange.Count Then Exit Sub
  '// 代码运行时关闭窗口刷新
  Application.Optimization = True
  ActiveDocument.BeginCommandGroup  '一步撤消'

   '// 设置当前文档 尺寸单位mm 出血和线长和线宽
  ActiveDocument.Unit = cdrMillimeter
  Bleed = API.GetSet("Bleed")
  Line_len = API.GetSet("Line_len")
  Outline_Width = API.GetSet("Outline_Width")

  Dim OrigSelection As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  
  '// 定义当前选择物件 分别获得 左右下上中心坐标(x,y)和尺寸信息
  Dim s1 As Shape

  For Each Target In OrigSelection
    Set s1 = Target
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
    ActiveSelection.Outline.SetProperties Outline_Width
    ActiveSelection.Outline.SetProperties Color:=CreateRegistrationColor
  
  Next Target

  ActiveDocument.EndCommandGroup
  '// 代码操作结束恢复窗口刷新
  Application.Optimization = False
  ActiveWindow.Refresh
  Application.Refresh
End Sub



'// 单线条转裁切线 - 放置到页面四边
Sub SelectLine_to_Cropline()
  If 0 = ActiveSelectionRange.Count Then Exit Sub
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
End Sub
