Attribute VB_Name = "裁切线"
Sub start()

     '// 设置当前文档 尺寸单位mm 出血和线长
    ActiveDocument.Unit = cdrMillimeter
    Bleed = 2
    line_len = 3

    '// 定义当前选择物件 分别获得 左右下上中心坐标(x,y)和尺寸信息
    Dim s1 As Shape
    Set s1 = ActiveSelection
    
    lx = s1.LeftX
    rx = s1.RightX
    by = s1.BottomY
    ty = s1.TopY
    
    cx = s1.CenterX
    cy = s1.CenterY
    sw = s1.SizeWidth
    sh = s1.SizeHeight

    '//  添加裁切线，分别左下-右下-左上-右上 
    Dim s2, s3, s4, s5, s6, s7, s8, s9 As Shape
    Set s2 = ActiveLayer.CreateLineSegment(lx - Bleed, by, lx - (Bleed + line_len), by)
    Set s3 = ActiveLayer.CreateLineSegment(lx, by - Bleed, lx, by - (Bleed + line_len))

    Set s4 = ActiveLayer.CreateLineSegment(rx + Bleed, by, rx + (Bleed + line_len), by)
    Set s5 = ActiveLayer.CreateLineSegment(rx, by - Bleed, rx, by - (Bleed + line_len))

    Set s6 = ActiveLayer.CreateLineSegment(lx - Bleed, ty, lx - (Bleed + line_len), ty)
    Set s7 = ActiveLayer.CreateLineSegment(lx, ty + Bleed, lx, ty + (Bleed + line_len))

    Set s8 = ActiveLayer.CreateLineSegment(rx + Bleed, ty, rx + (Bleed + line_len), ty)
    Set s9 = ActiveLayer.CreateLineSegment(rx, ty + Bleed, rx, ty + (Bleed + line_len))

    '// 选中裁切线 群组 设置线宽和注册色
    ActiveDocument.AddToSelection s2, s3, s4, s5, s6, s7, s8, s9
    ActiveSelection.Group
    ActiveSelection.Outline.SetProperties 0.1
    ActiveSelection.Outline.SetProperties Color:=CreateRegistrationColor

End Sub
