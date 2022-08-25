Attribute VB_Name = "裁切编号"
Sub ShapesRange()
    '// 代码运行时关闭窗口刷新
    Application.Optimization = True
        
    Dim d As Document
    Dim number As Shape
    Dim cnt As Integer
    cnt = 1
    Set d = ActiveDocument
     
    With ActiveLayer.Shapes
        MsgBox "总共有物件个数 " & .Count
    End With
    
    Dim s1 As Shape
    For Each Target In ActiveLayer.Shapes
        Set s1 = Target
        '设置颜色 s1.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
        
        cx = s1.CenterX
        cy = s1.CenterY
        sw = s1.SizeWidth
        sh = s1.SizeHeight
        
        Text = Trim(Str(cnt))
        Set number = d.ActiveLayer.CreateArtisticText(cx, cy, Text)
        cnt = cnt + 1
    Next Target
    
    '// 代码操作结束恢复窗口刷新
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

