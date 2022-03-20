'// CorelDRAW 物件排列拼版简单代码
Sub arrange()

    ActiveDocument.Unit = cdrMillimeter
    Bleed = 2
    line_len = 3
    
    Size = 50   '尺寸 50x50mm
    sp = 3      '间隔 3mm
    row = 3     ' 拼版 3 x 4
    List = 4

    '// 当前选择物件 按行3列4间隔3mm拼版
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    
    '// StepAndRepeat 方法在范围内创建多个形状副本
    Dim dup1 As ShapeRange
    Set dup1 = OrigSelection.StepAndRepeat(row - 1, Size + sp, 0#)
    Dim dup2 As ShapeRange
    Set dup2 = ActiveDocument.CreateShapeRangeFromArray _
         (dup1, OrigSelection).StepAndRepeat(List - 1, 0#, -(Size + sp))
End Sub
