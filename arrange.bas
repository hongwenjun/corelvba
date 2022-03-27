'// CorelDRAW 物件排列拼版简单代码
Sub arrange()
    On Error GoTo ErrorHandler
    ActiveDocument.Unit = cdrMillimeter
    row = 3     ' 拼版 3 x 4
    List = 4
    sp = 0       '间隔 0mm
    
    Dim Str, arr, n
    Str = GetClipBoardString

    ' 替换 mm x * 换行 TAB 为空格
    Str = VBA.Replace(Str, "mm", " ")
    Str = VBA.Replace(Str, "x", " ")
    Str = VBA.Replace(Str, "*", " ")
    Str = VBA.Replace(Str, Chr(13), " ")
    Str = VBA.Replace(Str, Chr(9), " ")
    
    Do While InStr(Str, "  ") '多个空格换成一个空格
        Str = VBA.Replace(Str, "  ", " ")
    Loop
    
    arr = Split(Str)

    Dim x As Double
    Dim y As Double
    x = Val(arr(0))
    y = Val(arr(1))
    
    If UBound(arr) > 2 Then
    row = Val(arr(2))     ' 拼版 3 x 4
    List = Val(arr(3))
        If UBound(arr) > 3 Then
            sp = Val(arr(4))       '间隔
        End If
    End If
    
    Dim s1 As Shape
    '// 建立矩形 Width  x Height 单位 mm
    Set s1 = ActiveLayer.CreateRectangle(0, 0, x, y)
    
    '// 填充颜色无，轮廓颜色 K100，线条粗细0.3mm
    s1.Fill.ApplyNoFill
    s1.Outline.SetProperties 0.3, OutlineStyles(0), CreateCMYKColor(0, 100, 0, 0), ArrowHeads(0), _
        ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#

    sw = x
    sh = y
    
    '// StepAndRepeat 方法在范围内创建多个形状副本
    Dim dup1 As ShapeRange
    Set dup1 = s1.StepAndRepeat(row - 1, sw + sp, 0#)
    Dim dup2 As ShapeRange
    Set dup2 = ActiveDocument.CreateShapeRangeFromArray _
         (dup1, s1).StepAndRepeat(List - 1, 0#, (sh + sp))
         
    Exit Sub
ErrorHandler:
     MsgBox "记事本输入数字,示例: 50x50 4x3 ,复制到剪贴板再运行工具!"
    On Error Resume Next
End Sub

Private Function GetClipBoardString() As String
    On Error Resume Next
    Dim MyData As New DataObject
    GetClipBoardString = ""
    MyData.GetFromClipboard
    GetClipBoardString = MyData.GetText
    Set MyData = Nothing
End Function
