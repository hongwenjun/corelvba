Sub cut_lines()
    ActiveDocument.Unit = cdrMillimeter
    Bleed = 2
    line_len = 3
    Dim s As Shape
    Dim line As Shape
    For Each s In ActiveSelection.Shapes
       cx = s.CenterX
       cy = s.CenterY
       sw = s.SizeWidth
       sh = s.SizeHeight
       
       If sw > sh Then
        s.Delete
        Set line = ActiveLayer.CreateLineSegment(0, cy, 0 + line_len, cy)
       End If
       
       If sw < sh Then
        s.Delete
        Set line = ActiveLayer.CreateLineSegment(cx, 0, cx, 0 + line_len)
       End If
       
    Next s
End Sub
