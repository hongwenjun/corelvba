VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CardsToolsForm 
   Caption         =   "CardsTools 2025"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5025
   OleObjectBlob   =   "CardsToolsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CardsToolsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private DIY_SIZE(1 To 2) As Double
Private flag_size As Boolean


' 这里修改绑定编号
Private Sub Combo_Material_Change()
    If Combo_Material.ListIndex >= 0 Then
        If Combo_Material.ListIndex <= 1 Then
            Text_SerialNumber.text = "2159"
        Else
            Text_SerialNumber.text = "2054"
        End If
    End If
End Sub


Private Sub UserForm_Initialize()
    ' Combo_Material 材质
    With Combo_Material
        .AddItem "亮"  '// 文件名 替换成 过
        .AddItem "不"  '// 前两项， 编号 2159
        
        .AddItem "星"  '// 后面项， 编号 2054
        .AddItem "虹"
        .AddItem "珠光"
        .AddItem "碎"
        .AddItem "厚亮"
        .AddItem "厚过"
        .AddItem "厚星"
        .AddItem "厚虹"
        .AddItem "厚碎"
        .ListIndex = 0 ' 默认选中第一项
        
        ' 设置列表显示行数（等于或大于项目总数）
        .ListRows = .ListCount  ' 显示所有项目
    End With

    ' Combo_Single_Double 单双面
    With Combo_Single_Double
        .AddItem "双面"
        .AddItem "单面"
        .ListIndex = 0 ' 默认选中第一项
    End With

    ' Combo_Quantity 数量
    With Combo_Quantity
        .AddItem "(1)"
        .AddItem "(2)"
        .AddItem "(5)"
        .AddItem "(10)"
        .AddItem "(20)"
        .AddItem "(30)"
        .AddItem "(40)"
        .ListIndex = 2 ' 默认选中第一项
    End With

    ' Combo_StyleCount 款数
    With Combo_StyleCount
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .AddItem "10"
        .ListIndex = 0 ' 默认选中第一项
        
        ' 设置列表显示行数（等于或大于项目总数）
        .ListRows = .ListCount  ' 显示所有项目
    End With

    ' Combo_Process 工艺
    With Combo_Process
        .AddItem ""
        .AddItem "后工[切圆角(圆四角)]"
        .AddItem "后工[特规模切(圆角85X54)]"
        .AddItem "后工[特规模切(票根120X60)]"
        .AddItem "后工[特规模切(票根140X70)]"
        .AddItem "后工[压痕(居中横向压1痕)]"
        .AddItem "后工[压痕(居中竖向压1痕)]"
        .ListIndex = 0 ' 默认选中第一项
        
        ' 设置列表显示行数（等于或大于项目总数）
        .ListRows = .ListCount  ' 显示所有项目
    End With
End Sub

Private Sub MakeRectangle(w As Double, h As Double, Optional ByVal onekey_images As Boolean = False)
    If Documents.count = 0 Then CreateDocument
    API.BeginOpt
    If onekey_images Then
        Call Images2NewDoc
    End If
    Call MakeRectangleToPowerClip(w, h)
    DIY_SIZE(1) = w: DIY_SIZE(2) = h
    API.EndOpt
End Sub

'///***** 批量尺寸按钮代码 *****///
Private Sub BT_54x85mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(54, 85)
    Else
        Call MakeRectangle(54, 85, True)
    End If
End Sub

Private Sub BT_85x54mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(85, 54)
    Else
        Call MakeRectangle(85, 54, True)
    End If
End Sub

Private Sub BT_90x54mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(90, 54)
    Else
        Call MakeRectangle(90, 54, True)
    End If
End Sub

Private Sub BT_54x90mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(54, 90)
    Else
        Call MakeRectangle(54, 90, True)
    End If
End Sub

Private Sub BT_90x90mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(90, 90)
    Else
        Call MakeRectangle(90, 90, True)
    End If
End Sub

Private Sub BT_89x58mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(89, 58)
    Else
        Call MakeRectangle(89, 58, True)
    End If
End Sub

Private Sub BT_58x89mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(58, 89)
    Else
        Call MakeRectangle(58, 89, True)
    End If
End Sub

Private Sub BT_140x95mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(140, 95)
    Else
        Call MakeRectangle(140, 95, True)
    End If
End Sub

Private Sub BT_95x140mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(95, 140)
    Else
        Call MakeRectangle(95, 140, True)
    End If
End Sub

Private Sub BT_150x100mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(150, 100)
    Else
        Call MakeRectangle(150, 100, True)
    End If
End Sub

Private Sub BT_100x150mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(100, 150)
    Else
        Call MakeRectangle(100, 150, True)
    End If
End Sub

Private Sub BT_100x100mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(100, 100)
    Else
        Call MakeRectangle(100, 100, True)
    End If
End Sub

Private Sub BT_54x54mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(54, 54)
    Else
        Call MakeRectangle(54, 54, True)
    End If
End Sub

Private Sub BT_60x120mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(60, 120)
    Else
        Call MakeRectangle(60, 120, True)
    End If
End Sub

Private Sub BT_120x60mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(120, 60)
    Else
        Call MakeRectangle(120, 60, True)
    End If
End Sub

Private Sub BT_70x140mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(70, 140)
    Else
        Call MakeRectangle(70, 140, True)
    End If
End Sub

Private Sub BT_140x70mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(140, 70)
    Else
        Call MakeRectangle(140, 70, True)
    End If
End Sub

Private Sub BT_50x150mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(50, 150)
    Else
        Call MakeRectangle(50, 150, True)
    End If
End Sub

Private Sub BT_150x50mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(150, 50)
    Else
        Call MakeRectangle(150, 50, True)
    End If
End Sub

Private Sub BT_100x300mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(100, 300)
    Else
        Call MakeRectangle(100, 300, True)
    End If
End Sub

Private Sub BT_300x100mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(300, 100)
    Else
        Call MakeRectangle(300, 100, True)
    End If
End Sub

Private Sub BT_150x450mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(150, 450)
    Else
        Call MakeRectangle(150, 450, True)
    End If
End Sub

Private Sub BT_450x150mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(450, 150)
    Else
        Call MakeRectangle(450, 150, True)
    End If
End Sub

Private Sub BT_210x140mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(210, 140)
    Else
        Call MakeRectangle(210, 140, True)
    End If
End Sub

Private Sub BT_140x210mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(140, 210)
    Else
        Call MakeRectangle(140, 210, True)
    End If
End Sub

Private Sub BT_297x210mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(297, 210)
    Else
        Call MakeRectangle(297, 210, True)
    End If
End Sub

Private Sub BT_210x297mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(210, 297)
    Else
        Call MakeRectangle(210, 297, True)
    End If
End Sub

Private Sub BT_108x86mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(108, 86)
    Else
        Call MakeRectangle(108, 86, True)
    End If
End Sub

Private Sub BT_86x108mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(86, 108)
    Else
        Call MakeRectangle(86, 108, True)
    End If
End Sub

Private Sub BT_127x89mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(127, 89)
    Else
        Call MakeRectangle(127, 89, True)
    End If
End Sub

Private Sub BT_89x127mm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(89, 127)
    Else
        Call MakeRectangle(89, 127, True)
    End If
End Sub

'//////////////////////////////////


' 生成格式化字符串的函数
Public Function GenerateFormattedString() As String
    Dim result As String
    Dim separator As String
    Dim size_xy As String
    Dim mtl As String
    
    
    separator = "-" ' 分隔符
    
    
    ' 构建各部分
    result = Trim(Text_SerialNumber.text) & separator & _
             Replace(Trim(Text_OrderNumber.text), "-", "") & separator & "@名片"
    
    ' 添加材质（如果选择了）
    If Combo_Material.ListIndex >= 0 Then
        mtl = Combo_Material.text
        If mtl = "亮" Then mtl = "过"
        
        result = result & "_" & mtl
    End If
    
    ' 添加尺寸（如果有）
    If DIY_SIZE(1) > 10 And DIY_SIZE(2) > 10 Then
        size_xy = DIY_SIZE(1) & "X" & DIY_SIZE(2)
        
        If size_xy = "89X58" Then
          size_xy = Replace(size_xy, "89X58", "85X54")
        End If
        
        If size_xy = "58X89" Then
          size_xy = Replace(size_xy, "58X89", "54X85")
        End If
        
        result = result & "_" & size_xy
    End If
    
    ' 添加单双面（如果选择了）
    If Combo_Single_Double.ListIndex >= 0 Then
        ' 去掉前后的下划线（如果不需要的话）
        Dim singleDouble As String
        singleDouble = Combo_Single_Double.text
        singleDouble = Replace(singleDouble, "_", "")
        result = result & "_" & singleDouble
    End If
    
    ' 添加数量（如果选择了）
    If Combo_Quantity.ListIndex >= 0 Then
        ' 去掉括号和下划线
        Dim quantity As String
        quantity = Combo_Quantity.text
        quantity = Replace(quantity, "_", "")
        result = result & "_数量" & quantity
    End If
    
    ' 添加款数（如果选择了）
    If Combo_StyleCount.ListIndex >= 0 Then
        result = result & "_" & Combo_StyleCount.text & "款"
    End If
    
    ' 添加工艺（如果选择了且不是空项）
    If Combo_Process.ListIndex >= 1 Then
        Dim processText As String
        processText = Combo_Process.text
        
        ' 去掉前导下划线
        If Left(processText, 1) = "_" Then
            processText = Mid(processText, 2)
        End If
        
        result = result & "_" & processText
    End If
    
    GenerateFormattedString = result
End Function


Private Sub BT_ReadFileName_Click()
'    Dim clipText As String
    ' 从剪贴板获取文本
'    clipText = GetClipBoardString()

    ' 检查剪贴板内容是否为空
'    If clipText = "" Or clipText = vbNullString Then
'       CDRX4_FileName.text = "请先准备好文件名文字复制到剪贴板"
'    Else
'        CDRX4_FileName.text = clipText
'    End If

   ' 验证必填项
    If Trim(Text_SerialNumber.text) = "" Then
        MsgBox "请填写编号", vbExclamation
        Text_SerialNumber.SetFocus
        Exit Sub
    End If
    
    If Trim(Text_OrderNumber.text) = "" Then
        MsgBox "请填写订单号", vbExclamation
        Text_OrderNumber.SetFocus
        Exit Sub
    End If
    
    ' 生成格式化字符串
    Dim formattedText As String
    formattedText = GenerateFormattedString()
    
    ' 显示结果（可以根据需要复制到剪贴板或显示在文本框中）
    ' MsgBox "生成的格式：" & vbCrLf & vbCrLf & formattedText, vbInformation
    
    CDRX4_FileName.text = formattedText


End Sub


Private Sub ClearText_OrderNumber_FileName()
    On Error Resume Next
    CDRX4_FileName.text = ""
    Text_OrderNumber.text = ""

'//  填加重置  工艺 和 自定义尺寸到默认
    Combo_Material.ListIndex = 0
    SIZE_WIDTH.text = ""
    SIZE_HEIGHT.text = ""
    
End Sub

Private Sub BT_SaveCDRX4_Click()
    file = "D:\Cards\CDR保存CDR文件\" & CDRX4_FileName.text & ".cdr"
    Save_CdrX4_File (file)
    ClearText_OrderNumber_FileName
End Sub

Private Sub Photo_Import_Click()
    Call Import_Images
End Sub

Private Sub PWC_Extract_Click()
    Call PowerClip_ExtractShapes
End Sub

Private Sub SIZE_WIDTH_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim Numbers As String
    Numbers = "1234567890"
    If InStr(Numbers, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

' 在KeyPress事件中只控制输入
Private Sub SIZE_HEIGHT_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim Numbers As String
    Numbers = "1234567890"
    If InStr(Numbers, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

' 新增Change事件处理
Private Sub SIZE_HEIGHT_Change()
    UpdateSizePreview
End Sub

Private Sub SIZE_WIDTH_Change()
    UpdateSizePreview
End Sub

' 统一更新函数
Private Sub UpdateSizePreview()
    On Error Resume Next

    Dim sx As Integer, sy As Integer

    ' 转换为整数
    sx = CInt(SIZE_WIDTH.value)
    sy = CInt(SIZE_HEIGHT.value)

    ' 检查有效值
    If sx > 29 And sy > 29 Then
        Dim txt As String
        txt = sx & "x" & sy & "mm"
        BT_DIY_SIZE.Caption = txt
        DIY_SIZE(1) = sx
        DIY_SIZE(2) = sy
        flag_size = True
    Else
        BT_DIY_SIZE.Caption = "自定义尺寸"
        flag_size = False
    End If
End Sub

Private Sub BT_DIY_SIZE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If BT_DIY_SIZE.Caption = "自定义尺寸" Then
        Exit Sub
    End If

    Dim sx As Double
    Dim sy As Double
    If flag_size = True Then
        sx = DIY_SIZE(1)
        sy = DIY_SIZE(2)
    End If

    If Button = 2 Then

    ElseIf Shift = fmCtrlMask Then
        Call MakeRectangle(sx, sy)
    Else
        Call MakeRectangle(sx, sy, True)
    End If
End Sub

Private Sub BT_GET_Size_Click()
    ActiveDocument.Unit = cdrMillimeter
    Set sr = ActiveSelectionRange
    sx = sr.SizeWidth: sy = sr.SizeHeight
    sx = Int(sx + 0.5): sy = Int(sy + 0.5)
    txt = sx & "x" & sy & "mm"
    BT_DIY_SIZE.Caption = txt
    DIY_SIZE(1) = sx
    DIY_SIZE(2) = sy
    flag_size = True
End Sub
