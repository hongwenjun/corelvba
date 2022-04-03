'// 请先选择要印刷的物件群组，本插件完成设置页面大小，自动中线色阶条对准线功能
Sub Auto_ColorMark()
    On Error GoTo ErrorHandler
    ActiveDocument.BeginCommandGroup:  Application.Optimization = True
    Dim doc As Document: Set doc = ActiveDocument: doc.Unit = cdrMillimeter

    ' 物件群组，设置页面大小
    Call set_page_size

    '// 获得页面中心点 x,y
    px = ActiveDocument.ActivePage.CenterX
    py = ActiveDocument.ActivePage.CenterY
    '// 导入色阶条中线对准线标记文件 ColorMark.cdr 解散群组
    doc.ActiveLayer.Import Path & "GMS\ColorMark.cdr"
    ActiveDocument.ReferencePoint = cdrBottomMiddle
    ActiveDocument.Selection.SetPosition px, -100
    ActiveDocument.Selection.Ungroup

    Dim sh As Shape, shs As Shapes
    Set shs = ActiveSelection.Shapes
    '// 按 MarkName 名称查找放置中线对准线标记等
    For Each sh In shs
    ActiveDocument.ClearSelection
    sh.CreateSelection
    If "CenterLine" = sh.ObjectData("MarkName").Value Then
        put_center_line sh
        
    ElseIf "TargetLine" = sh.ObjectData("MarkName").Value Then
        put_target_line sh

    ElseIf "ColorStrip" = sh.ObjectData("MarkName").Value Then
        put_ColorStrip sh   ' 放置彩色色阶条

        ' sh.Delete  ' 工厂定置不用色阶条

    ElseIf "ColorMark" = sh.ObjectData("MarkName").Value Then
        ' CMYK四色标记放置咬口
        If (px > py) Then
        sh.SetPosition px + 30#, 0
        Else
        sh.Rotate 270#
        ActiveDocument.ReferencePoint = cdrBottomLeft
        sh.SetPosition 0, py - 52#
        End If
    Else
        sh.Delete ' 没找到标记 ColorMark 删除
    
    End If
    Next sh

    ' 标准页面大小和添加页面框
    put_page_size
    put_page_line
    
    '// 代码操作结束恢复窗口刷新
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh:    Application.Refresh
Exit Sub
ErrorHandler:
    MsgBox "请先选择要印刷的物件群组，本插件完成设置页面大小，自动中线色阶条对准线功能!"
    Application.Optimization = False
    On Error Resume Next
End Sub

Private Sub set_page_size()
    ' 实践应用: 选择物件群组,页面设置物件大小,物件页面居中
    ActiveDocument.Unit = cdrMillimeter
    Dim OrigSelection As ShapeRange, sh As Shape
    Set OrigSelection = ActiveSelectionRange
    Set sh = OrigSelection.Group

    ' MsgBox "选择物件尺寸: " & sh.SizeWidth & "x" & sh.SizeHeight
    ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
    sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
End Sub

Private Function put_target_line(sh As Shape)
    ' 在页面四角放置套准标记线  Set sh = ActiveDocument.Selection
    sh.AlignToPage cdrAlignLeft + cdrAlignTop
    sh.Duplicate 0, 0
    sh.Rotate 180
    sh.AlignToPage cdrAlignRight + cdrAlignBottom
    sh.Duplicate 0, 0
    sh.Flip cdrFlipHorizontal   ' 物件镜像
    sh.AlignToPage cdrAlignLeft + cdrAlignBottom
    sh.Duplicate 0, 0
    sh.Rotate 180
    sh.AlignToPage cdrAlignRight + cdrAlignTop
End Function

Private Function put_center_line(sh As Shape)
    ' 在页面四边放置中线 Set sh = ActiveDocument.Selection
    sh.AlignToPage cdrAlignHCenter + cdrAlignTop
    sh.Duplicate 0, 0
    sh.Rotate 180
    sh.AlignToPage cdrAlignHCenter + cdrAlignBottom
    sh.Duplicate 0, 0
    sh.Rotate 90
    sh.AlignToPage cdrAlignVCenter + cdrAlignRight
    sh.Duplicate 0, 0
    sh.Rotate 180
    sh.AlignToPage cdrAlignVCenter + cdrAlignLeft
End Function

Private Function put_ColorStrip(sh As Shape)
  ' 在页面四边放置中线 Set sh = ActiveDocument.Selection
    sh.OrderToBack
  If ActivePage.SizeWidth >= ActivePage.SizeHeight Then
    sh.AlignToPage cdrAlignLeft + cdrAlignTop
    sh.Duplicate 10, 0
    sh.AlignToPage cdrAlignRight + cdrAlignTop
    sh.Duplicate -25, 0
    sh.Rotate 90
    sh.AlignToPage cdrAlignLeft + cdrAlignBottom
    sh.Duplicate 0, 10
    sh.AlignToPage cdrAlignRight + cdrAlignBottom
    sh.Move 0, 10
  Else
    sh.AlignToPage cdrAlignLeft + cdrAlignTop
    sh.Duplicate 10, 0
    sh.AlignToPage cdrAlignLeft + cdrAlignBottom
    sh.Duplicate 10, 0
    sh.Rotate 270
    sh.AlignToPage cdrAlignRight + cdrAlignTop
    sh.Duplicate 0, -10
    sh.AlignToPage cdrAlignRight + cdrAlignBottom
    sh.Move 0, 25
  End If
End Function

Private Function put_page_line()
    ' 添加页面框线
    Dim s1 As Shape
    Set s1 = ActiveLayer.CreateRectangle2(0, 0, ActivePage.SizeWidth, ActivePage.SizeHeight)
    s1.Fill.ApplyNoFill:    s1.OrderToBack
    s1.Outline.SetProperties 0.04, Color:=CreateCMYKColor(0, 100, 0, 0)
End Function

Private Function put_page_size()
    ' 添加文字 页面大小
    Dim st As Shape
    size = Trim(Str(Int(ActivePage.SizeWidth))) + "x" + Trim(Str(Int(ActivePage.SizeHeight))) + "mm"
    Set st = ActiveLayer.CreateArtisticText(0, 0, size, , , "Arial", 8)
    st.AlignToPage cdrAlignRight + cdrAlignTop
    st.Move -3, -0.2
End Function

