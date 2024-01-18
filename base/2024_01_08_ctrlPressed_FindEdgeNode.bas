Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Sub 工具栏图标Ctrl扩展功能()
    Dim ctrlPressed As Boolean
    
    ' 获取Ctrl键的状态
    ctrlPressed = GetAsyncKeyState(17) And &H8000
    
    ' 检查Ctrl键是否按下
    If ctrlPressed Then
        MsgBox "Ctrl键被按下了。 我要运行强大的扩展功能了"

         autogroup("group", 1).CreateSelection
    Else

    Tools.guideangle ActiveSelectionRange, 0#   ' 右键 0距离贴紧
    
        MsgBox "Ctrl键未被按下。我躺平了。"
    End If
End Sub


'// 日醺Apollo 2024-01-18 快速查极点测试成功，代码如下，供参考。

Sub FindEdgeNode()
    Dim s As Shape, nd As Node
    Dim ndIndex As Integer
    ActiveDocument.Unit = cdrMillimeter
    Dim x As Double, y As Double, w As Double, h As Double
    Set s = ActiveShape
    s.GetBoundingBox x, y, w, h
    s.SetBoundingBox x, y, 1, h
    Set nd = s.Curve.FindNodeAtPoint(x, y + h)
    If Not nd Is Nothing Then
        ndIndex = nd.Index
        MsgBox "当前顶点的序号为" & ndIndex & "   其座标x为" & nd.PositionX & "   y为" & nd.PositionY, vbCritical
    End If
End Sub

