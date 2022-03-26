Sub Shapes_Get_Coordinates()
    '// 建立文件 testfile.txt 输出物件坐标信息
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile("R:\testfile.txt", True)
    
    ActiveDocument.Unit = cdrMillimeter
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    
     '// 代码运行时关闭窗口刷新
    Application.Optimization = True
    
    With OrigSelection
        MsgBox "选择物件个数 " & OrigSelection.Count & "   尺寸:" & .SizeWidth & " x " & .SizeHeight
        f.WriteLine ("选择物件个数 " & OrigSelection.Count & "   尺寸:" & .SizeWidth & " x " & .SizeHeight)
        
        lx = OrigSelection.LeftX
        rx = OrigSelection.RightX
        by = OrigSelection.BottomY
        ty = OrigSelection.TopY
        
        f.WriteLine ("选择物件集合坐标范围: " & "(" & lx & "," & by & ") " & "(" & rx & "," & by & ") " _
            & "(" & lx & "," & ty & ") " & "(" & rx & "," & ty & ") ")
        f.WriteLine ("--------- 分割 ---------")
    End With
    
    Dim s1 As Shape
    cnt = 1
    For Each Target In OrigSelection
        Set s1 = Target
        lx = s1.LeftX
        rx = s1.RightX
        by = s1.BottomY
        ty = s1.TopY
        
        '// 遍历物件，输出左下-右下-左上-右上四点坐标
        f.WriteLine (cnt & "号物件坐标: " & "(" & lx & "," & by & ") " & "(" & rx & "," & by & ") " _
            & "(" & lx & "," & ty & ") " & "(" & rx & "," & ty & ") ")
        cnt = cnt + 1
    Next Target
    
    f.Close
    '// 代码操作结束恢复窗口刷新
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub
