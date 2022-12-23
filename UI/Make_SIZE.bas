VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Make_SIZE 
   Caption         =   " 标注尺寸"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3690
   OleObjectBlob   =   "Make_SIZE.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Make_SIZE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    With Tis
        .BackColor = RGB(0, 150, 255)
        .BorderColor = RGB(30, 150, 255)
        .ForeColor = RGB(255, 255, 255)
    End With
End Sub

Private Function 按钮移入(T)
    With T
        .BackColor = RGB(0, 150, 255)
        .BorderColor = RGB(30, 150, 255)
        .ForeColor = RGB(255, 255, 255)
    End With
End Function

Private Function 命令按钮(T As Label)
    With T
        .BackColor = RGB(240, 240, 240)
        .BorderColor = RGB(100, 100, 100)
        .ForeColor = RGB(0, 0, 0)
    End With
End Function

Private Sub CheckBox1_Click()
    If CheckBox1 Then CheckBox4 = False
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2 Then CheckBox4 = False
End Sub

Private Sub CheckBox3_Click()
    If CheckBox3 Then CheckBox1 = False: CheckBox2 = False: CheckBox4 = False
End Sub

Private Sub CheckBox4_Click()
    If CheckBox4 Then CheckBox1 = False: CheckBox2 = False: CheckBox3 = False
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call 命令按钮(标注)
    Call 命令按钮(删除)
End Sub

Private Sub SpinButton1_SpinDown()
    选中标注字号减少
End Sub

Private Sub SpinButton1_SpinUp()
    选中标注字号增加
End Sub

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    选中标注字号
End Sub

Private Sub 标注_Click()
    If CheckBox1 Or CheckBox2 Then Call 标注宽高度
    If CheckBox3 Then Call 标注线长
    If CheckBox4 Then Call 标注线段长
End Sub

Private Sub 标注_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call 按钮移入(标注)
End Sub

Private Sub 删除_Click()
    删除标注
End Sub

Private Sub 删除_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call 按钮移入(删除)
End Sub


Private Sub 标注宽高度()
    ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape, st1 As Shape, st2 As Shape
    Set s = ActiveShape
    If s Is Nothing Then Exit Sub
    Optimization = True '优化启动
    For Each s In ActiveSelection.Shapes
        If CheckBox1 Then
            Set st1 = ActiveLayer.CreateArtisticText(s.LeftX, s.TopY + 4, round(s.SizeWidth, 0) & "mm", , , "微软雅黑", TextBox1.value, , , , cdrCenterAlignment)
                st1.text.Story.CharSpacing = 0 '字符间距
                st1.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
                st1.Move s.SizeWidth / 2, 0
                st1.Name = "Text" ' 设置名
            Set sox = ActiveLayer.CreateLineSegment(s.LeftX, s.TopY + 3, s.RightX, s.TopY + 3)
                sox.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                sox.Name = "line"
            Set sox1 = ActiveLayer.CreateLineSegment(s.LeftX, s.TopY + 1, s.LeftX, s.TopY + 3)
                sox1.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                sox1.Name = "line"
            Set sox2 = ActiveLayer.CreateLineSegment(s.RightX, s.TopY + 1, s.RightX, s.TopY + 3)
                sox2.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                sox2.Name = "line"
            s.CreateSelection
        End If
        If CheckBox2 Then
            Set st2 = ActiveLayer.CreateArtisticText(s.LeftX - 4, s.BottomY, round(s.SizeHeight, 0) & "mm", , , "微软雅黑", TextBox1.value, , , , cdrCenterAlignment)
            st2.text.Story.CharSpacing = 0 '字符间距
            st2.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
            st2.Rotate 90
            st2.Move -st2.SizeWidth / 2, s.SizeHeight / 2
            st2.Name = "Text" ' 设置名
            Set soy = ActiveLayer.CreateLineSegment(s.LeftX - 3, s.BottomY, s.LeftX - 3, s.TopY)
                soy.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                soy.Name = "line"
            Set soy1 = ActiveLayer.CreateLineSegment(s.LeftX - 1, s.BottomY, s.LeftX - 3, s.BottomY)
                soy1.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                soy1.Name = "line"
            Set soy2 = ActiveLayer.CreateLineSegment(s.LeftX - 1, s.TopY, s.LeftX - 3, s.TopY)
                soy2.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                soy2.Name = "line"
            s.CreateSelection
        End If
    Next
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub 标注线段长()
    ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape, s1 As Shape, s2 As Shape, sc As Shape, st1 As Shape, st2 As Shape
    Set s = ActiveShape
    If s Is Nothing Then Exit Sub
    Optimization = True '优化启动
    For Each s In ActiveSelection.Shapes
        If s.Type <> cdrTextShape Then
            s.Copy
            Set sc = ActiveLayer.Paste
            sc.ConvertToCurves
            sc.Curve.Nodes.All.BreakApart
            sc.BreakApart
            For Each s1 In ActiveSelection.Shapes
                Set st1 = ActiveLayer.CreateArtisticText(0, 0, round(s1.Curve.Length, 0), , , , TextBox1.value)
                st1.text.Story.CharSpacing = 0 '字符间距
                st1.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
                st1.text.FitToPath s1
                ' 获取或设置文本与路径的偏移量
                st1.Effects(1).TextOnPath.Offset = s1.Curve.Length * 0.5 - st1.SizeWidth * 0.55
                ' 获取或设置文本与路径的距离
                st1.Effects(1).TextOnPath.DistanceFromPath = 1
                st1.Name = "Text" ' 设置名
                s1.Outline.SetNoOutline
                s1.OrderToBack
                s1.Name = "line"
            Next
            Set st2 = ActiveLayer.CreateArtisticText(s.RightX + 3, s.BottomY, "单位：mm", , , "微软雅黑", TextBox1.value)
            st2.text.Story.CharSpacing = 0 '字符间距
            st2.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
            st2.Name = "Text" ' 设置名
         End If
    Next
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub 标注线长()
    ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape, st1 As Shape
    Set s = ActiveShape
    If s Is Nothing Then Exit Sub
    Optimization = True '优化启动
    For Each s In ActiveSelection.Shapes
        If s.Type <> cdrTextShape Then
            X = s.LeftX
            Y = s.BottomY
            Set st1 = ActiveLayer.CreateArtisticText(X, Y, "线条长：" & round(s.DisplayCurve.Length, 0) & "mm", , , "微软雅黑", TextBox1.value, , , , cdrLeftAlignment)
            st1.text.Story.CharSpacing = 0 '字符间距
            st1.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
            st1.Move 0, -st1.SizeHeight * 2
            st1.Name = "Text" ' 设置名
            s.CreateSelection
        End If
    Next
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub 选中标注字号增加()
    Dim s As Shape
    Optimization = True '优化启动
    If TextBox1.value > 0 Then
        TextBox1.value = TextBox1.value + 1
        For Each s In ActiveSelection.Shapes.FindShapes(query:="@type ='text:artistic' and @Name='Text' ")
            s.text.Story.size = s.text.Story.size + 1
        Next
    End If
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub 选中标注字号减少()
    Dim s As Shape
    Optimization = True '优化启动
    If TextBox1.value > 0 Then
        TextBox1.value = TextBox1.value - 1
        For Each s In ActiveSelection.Shapes.FindShapes(query:="@type ='text:artistic' and @Name='Text' ")
            s.text.Story.size = s.text.Story.size - 1
        Next
    End If
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub 选中标注字号()
    Dim s As Shape
    Optimization = True '优化启动
    If TextBox1.value > 0 Then
        For Each s In ActiveSelection.Shapes.FindShapes(query:="@type ='text:artistic' and @Name='Text' ")
            s.text.Story.size = TextBox1.value
        Next
    End If
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub 删除标注()
    If ActiveSelection.Shapes.Count > 0 Then
        ActiveSelection.Shapes.FindShapes(query:="@type ='text:artistic' and @Name='Text' ").Delete
        ActiveSelection.Shapes.FindShapes(query:="@Name='line' ").Delete
    Else
        ActivePage.Shapes.FindShapes(query:="@type ='text:artistic' and @Name='Text' ").Delete
        ActivePage.Shapes.FindShapes(query:="@Name='line' ").Delete
    End If
End Sub

