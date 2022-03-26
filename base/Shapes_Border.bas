Sub Shapes_Border()
    '// 建立文件 testfile.txt 输出物件坐标信息
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile("R:\testfile.txt", True)
    
    ActiveDocument.Unit = cdrMillimeter
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    
    '// 代码运行时关闭窗口刷新
    Application.Optimization = True
    
    ' 当前选择物件的范围边界
    set_lx = OrigSelection.LeftX
    set_rx = OrigSelection.RightX
    set_by = OrigSelection.BottomY
    set_ty = OrigSelection.TopY
    set_cx = OrigSelection.CenterX
    set_cy = OrigSelection.CenterY
    radius = 20
    
    Dim s1 As Shape
    cnt = 1
    For Each Target In OrigSelection
        Set s1 = Target
        lx = s1.LeftX
        rx = s1.RightX
        by = s1.BottomY
        ty = s1.TopY
        
        If Abs(set_lx - lx) < radius Or Abs(set_rx - rx) < radius Or Abs(set_by - by) _
            < radius Or Abs(set_ty - ty) < radius Then
        
            '// 遍历物件，输出左下-右下-左上-右上四点坐标
            f.WriteLine (cnt & "号物件修改颜色: 绿色")
            s1.Fill.UniformColor.CMYKAssign 60, 0, 100, 0
        End If
        cnt = cnt + 1
    Next Target
    
    f.Close
    '// 代码操作结束恢复窗口刷新
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub
