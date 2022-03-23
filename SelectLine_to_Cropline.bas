'// 单线条转裁切线 - 放置到页面四边
Sub SelectLine_to_Cropline()

    '// 代码运行时关闭窗口刷新
    Application.Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    
    '// 获得页面中心点 x,y
    px = ActiveDocument.Pages.First.CenterX
    py = ActiveDocument.Pages.First.CenterY
    Bleed = 2
    line_len = 3
    
    Dim s As Shape
    Dim line As Shape
    
    '// 遍历选择的线条
    For Each s In ActiveSelection.Shapes
    
        lx = s.LeftX
        rx = s.RightX
        by = s.BottomY
        ty = s.TopY
        
        cx = s.CenterX
        cy = s.CenterY
        sw = s.SizeWidth
        sh = s.SizeHeight
       
       '// 判断横线(高度小于宽度)，在页面左边还是右边
       If sh < sw Then
        s.Delete
        If cx < px Then
            Set line = ActiveLayer.CreateLineSegment(0, cy, 0 + line_len, cy)
        Else
            Set line = ActiveLayer.CreateLineSegment(px * 2, cy, px * 2 - line_len, cy)
        End If
       End If
     
       '// 判断竖线(高度大于宽度)，在页面下边还是上边
       If sh > sw Then
        s.Delete
        If cy < py Then
            Set line = ActiveLayer.CreateLineSegment(cx, 0, cx, 0 + line_len)
        Else
            Set line = ActiveLayer.CreateLineSegment(cx, py * 2, cx, py * 2 - line_len)
        End If
       End If
    
        line.Outline.SetProperties 0.1
        line.Outline.SetProperties Color:=CreateRegistrationColor
    Next s
    
    '// 代码操作结束恢复窗口刷新
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

