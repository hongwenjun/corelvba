VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Toolbar 
   Caption         =   "Toolbar"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   OleObjectBlob   =   "Toolbar.frx":0000
End
Attribute VB_Name = "Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#End If
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&

'Constants for transparency
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1                  'Chroma key for fading a certain color on your Form
Private Const LWA_ALPHA = &H2                     'Only needed if you want to fade the entire userform

Public UIL_Key As Boolean
Public pic1, pic2

Private Sub MakeUserFormTransparent(frm As Object, Optional Color As Variant)
  'set transparencies on userform
  Dim formhandle As Long
  Dim bytOpacity As Byte
  
  formhandle = FindWindow(vbNullString, Me.Caption)
  If IsMissing(Color) Then Color = vbWhite 'default to vbwhite
  bytOpacity = 100 ' variable keeping opacity setting
  
  SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
  'The following line makes only a certain color transparent so the
  ' background of the form and any object whose BackColor you've set to match
  ' vbColor (default vbWhite) will be transparent.
  Me.BackColor = Color
  SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY
End Sub

Private Sub Change_UI_Close_Voice_Click()
  Speak_Msg "修改UI图片更换界面  注册表关闭语音 详QQ群"
  MsgBox "请给我支持!" & vbNewLine & "您的支持，我才能有动力添加更多功能." & vbNewLine & "蘭雅CorelVBA中秋节版" & vbNewLine & "coreldrawvba插件交流群  8531411"
End Sub


Private Sub UserForm_Initialize()
  Dim IStyle As Long
  Dim hWnd As Long
  
  hWnd = FindWindow("ThunderDFrame", Me.Caption)

  IStyle = GetWindowLong(hWnd, GWL_STYLE)
  IStyle = IStyle And Not WS_CAPTION
  SetWindowLong hWnd, GWL_STYLE, IStyle
  DrawMenuBar hWnd
  IStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
  SetWindowLong hWnd, GWL_EXSTYLE, IStyle
  
With Me
  .StartUpPosition = 0
  .Left = Val(GetSetting("262235.xyz", "Settings", "Left", "400"))  ' 设置工具栏位置
  .Top = Val(GetSetting("262235.xyz", "Settings", "Top", "55"))
  .Height = 30
  .Width = 336
End With

  OutlineKey = True
  OptKey = True

  ' 读取角线设置
  Bleed.text = API.GetSet("Bleed")
  Line_len.text = API.GetSet("Line_len")
  Outline_Width.text = GetSetting("262235.xyz", "Settings", "Outline_Width", "0.2")
  
  UIFile = Path & "GMS\262235.xyz\ToolBar.jpg"
  If API.ExistsFile_UseFso(UIFile) Then
    UI.Picture = LoadPicture(UIFile)   '换UI图
    Set pic1 = LoadPicture(UIFile)
  End If

  UIL = Path & "GMS\262235.xyz\ToolBar1.jpg"
  If API.ExistsFile_UseFso(UIL) Then
    Set pic2 = LoadPicture(UIL)
    UIL_Key = True
  End If

  ' 窗口透明, 最小化只显示一个图标
  #If VBA7 Then
    MakeUserFormTransparent Me, RGB(26, 22, 35)
  #Else
  ' CorelDRAW X4 / Windows7 自用关闭透明
  #End If
End Sub

Private Sub UI_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  UI.Visible = False
  If Y > 1 And Y < 16 And UIL_Key Then
    UI.Picture = pic2
  ElseIf Y > 16 And UIL_Key Then
    UI.Picture = pic1
  End If
  UI.Visible = True

  ' Debug.Print X & " , " & Y
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button Then
        mx = x
        my = Y
    End If
    
  With Me
    .Height = 30
  End With

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button Then
    Me.Left = Me.Left - mx + x
    Me.Top = Me.Top - my + Y
  End If
End Sub

Private Sub LOGO_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Abs(x - 14) < 14 And Abs(Y - 14) < 14 And Button = 2 Then
    Me.Width = 336
    OPEN_UI_BIG.Left = 322
    UI.Visible = True
    LOGO.Visible = False
    TOP_ALIGN_BT.Visible = False
    LEFT_ALIGN_BT.Visible = False
    Exit Sub
  ElseIf Shift = fmCtrlMask Then
      mx = x
      my = Y
  Else
    Unload Me   ' Ctrl + 鼠标 关闭工具
  End If
End Sub

Private Sub LOGO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button Then
    Me.Left = Me.Left - mx + x
    Me.Top = Me.Top - my + Y
  End If
End Sub

Private Sub UI_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Dim c As New Color
  ' 定义图标坐标pos
  Dim pos_x As Variant
  Dim pos_y As Variant
  pos_y = Array(14)
  pos_x = Array(14, 41, 67, 94, 121, 148, 174, 201, 228, 254, 281, 308, 334, 361, 388, 415, 441, 468, 495)

  '// 鼠标右键 扩展键按钮优先  收缩工具栏  标记范围框  居中页面 尺寸取整数  单色黑中线标记 扩展工具栏  排列工具  扩展工具栏收缩
  If Abs(x - pos_x(0)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    Me.Width = 30
    Me.Height = 30
    UI.Visible = False
    LOGO.Visible = True
    Exit Sub

  ElseIf Abs(x - pos_x(1)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    Tools.居中页面
    Exit Sub

  ElseIf Abs(x - pos_x(2)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    Tools.Mark_Range_Box
    Exit Sub

  ElseIf Abs(x - pos_x(3)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    Tools.尺寸取整
    Exit Sub
  
  ElseIf Abs(x - pos_x(5)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    自动中线色阶条.Auto_ColorMark_K
    Exit Sub
  
  '//分分合合把几个功能按键合并到一起，定义到右键上
  ElseIf Abs(x - pos_x(4)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    Tools.分分合合
    Exit Sub
  
  ElseIf Abs(x - pos_x(6)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    智能群组和查找.智能群组 API.Create_Tolerance
    Exit Sub
  
  ElseIf Abs(x - pos_x(8)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    '// 右键扩展工具栏
    Me.Height = 30 + 45
    Exit Sub
  
  ElseIf Abs(x - pos_x(9)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    '// 右键拆分线段
    Tools.Split_Segment
    Exit Sub
  
  ElseIf Abs(x - pos_x(10)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    '// 右键排列工具
    TOP_ALIGN_BT.Visible = True
    LEFT_ALIGN_BT.Visible = True
    Exit Sub

  ElseIf Abs(x - pos_x(11)) < 14 And Abs(Y - pos_y(0)) < 14 And Button = 2 Then
    '// 右键扩展工具栏收缩
    Me.Height = 30
    Exit Sub

  End If
  
  '// 鼠标左键 单击按钮功能  按工具栏上图标正常功能
  If Abs(x - pos_x(0)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    裁切线.start
    
  ElseIf Abs(x - pos_x(1)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    剪贴板尺寸建立矩形.start
    
  ElseIf Abs(x - pos_x(2)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    裁切线.SelectLine_to_Cropline
    
  ElseIf Abs(x - pos_x(3)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    拼版裁切线.arrange
    
  ElseIf Abs(x - pos_x(4)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    拼版裁切线.Cut_lines
    
  ElseIf Abs(x - pos_x(5)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    自动中线色阶条.Auto_ColorMark
    
  ElseIf Abs(x - pos_x(6)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    智能群组和查找.智能群组
    
  ElseIf Abs(x - pos_x(7)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    CQL_FIND_UI.Show 0
    
  ElseIf Abs(x - pos_x(8)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    Replace_UI.Show 0
    
  ElseIf Abs(x - pos_x(9)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    Tools.TextShape_ConvertToCurves
    
  ElseIf Abs(x - pos_x(10)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    '// 扩展工具栏
    Me.Height = 30 + 45
    
    Speak_Msg "左右键有不同功能"
    
  ElseIf Abs(x - pos_x(11)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    If Me.Height > 30 Then
      Me.Height = 30
    Else
      '// 最小化
      Me.Width = 30
      Me.Height = 30
      OPEN_UI_BIG.Left = 31
      UI.Visible = False
      LOGO.Visible = True
  
      ' 保存工具条位置 Left 和 Top
      SaveSetting "262235.xyz", "Settings", "Left", Me.Left
      SaveSetting "262235.xyz", "Settings", "Top", Me.Top
    
      Speak_Msg "左键缩小 右键收缩"
    End If
  End If

End Sub

Private Sub X_EXIT_Click()
  Unload Me    ' 关闭
End Sub


Private Sub 调用多页合并工具()
  Dim value As Integer
  value = GMSManager.RunMacro("合并多页工具", "合并多页运行.run")
End Sub


'''///  贪心商人和好玩工具等  ///'''

Private Sub Cdr_Nodes_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    TSP.Nodes_To_TSP
  ElseIf Shift = fmCtrlMask Then
    TSP.CDR_TO_TSP
  Else
    ' Ctrl + 鼠标  空
  End If
End Sub

Private Sub Cdr_Nodes_BT_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  TSP_L1.ForeColor = RGB(0, 150, 255)
End Sub

Private Sub START_TSP_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  TSP_L2.ForeColor = RGB(0, 150, 255)
End Sub

Private Sub PATH_TO_TSP_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  TSP_L3.ForeColor = RGB(0, 150, 255)
End Sub

Private Sub TSP2DRAW_LINE_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  TSP_L4.ForeColor = RGB(0, 150, 255)
End Sub

Private Sub TSP2DRAW_LINE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    TSP.TSP_TO_DRAW_LINE
  ElseIf Shift = fmCtrlMask Then
    TSP.TSP_TO_DRAW_LINES
  Else
    ' Ctrl + 鼠标  空
  End If
End Sub


Private Sub START_TSP_Click()
  TSP.START_TSP
End Sub

Private Sub PATH_TO_TSP_Click()
  TSP.MAKE_TSP
End Sub

Private Sub BITMAP_BUILD_Click()
  Tools.Python_BITMAP
End Sub

Private Sub BITMAP_BUILD2_Click()
  Tools.Python_BITMAP2
End Sub

Private Sub BITMAP_MAKE_DOTS_Click()
  TSP.BITMAP_MAKE_DOTS
End Sub

'''///  Python脚本和二维码等  ///'''
Private Sub Organize_Size_Click()
  Tools.Python_Organize_Size
End Sub

Private Sub Get_Number_Click()
  Tools.Python_Get_Barcode_Number
End Sub

Private Sub Make_QRCode_Click()
  Tools.Python_Make_QRCode
  Tools.QRCode_replace
End Sub

Private Sub QR2Vector_Click()
  Tools.QRCode_to_Vector
End Sub

Private Sub OPEN_UI_BIG_Click()
  Unload Me
  CorelVBA.Show 0
End Sub

Private Sub Settings_Click()
  If 0 < Val(Bleed.text) * Val(Line_len.text) < 100 Then
   SaveSetting "262235.xyz", "Settings", "Bleed", Bleed.text
   SaveSetting "262235.xyz", "Settings", "Line_len", Line_len.text
   SaveSetting "262235.xyz", "Settings", "Outline_Width", Outline_Width.text
  End If

  ' 保存工具条位置 Left 和 Top
  SaveSetting "262235.xyz", "Settings", "Left", Me.Left
  SaveSetting "262235.xyz", "Settings", "Top", Me.Top
  
  Me.Height = 30
End Sub


'''/////////  图标鼠标左右点击功能调用   /////////'''

Private Sub Tools_Icon_Click()
  ' 调用语句
  I = GMSManager.RunMacro("CorelDRAW_VBA", "学习CorelVBA.start")
End Sub

'''////  选择多物件，组合然后拆分线段，为角线爬虫准备  ////'''
Private Sub Split_Segment_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    MsgBox "鼠标右键，功能待定"
    Exit Sub
  End If
  
  If Button Then
      Tools.Split_Segment
  End If
End Sub

Private Sub Split_Segment_Copy_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    MsgBox "鼠标右键，功能待定"
    Exit Sub
  End If
  
  If Button Then
      Tools.Split_Segment
  End If
  Speak_Msg "拆分线段"
End Sub

'''////  CorelDRAW 与 Adobe_Illustrator 剪贴板转换  ////'''
Private Sub Adobe_Illustrator_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Dim value As Integer
  If Button = 2 Then
    value = GMSManager.RunMacro("AIClipboard", "CopyPaste.PasteAIFormat")
    Exit Sub
  End If
  
  If Button Then
    value = GMSManager.RunMacro("AIClipboard", "CopyPaste.CopyAIFormat")
    MsgBox "CorelDRAW 与 Adobe_Illustrator 剪贴板转换" & vbNewLine & "鼠标左键复制，鼠标右键粘贴"
  End If
End Sub

'''////  标记画框 支持容差  ////'''
Private Sub Mark_CreateRectangle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Mark_CreateRectangle True
  ElseIf Shift = fmCtrlMask Then
    Tools.Mark_CreateRectangle False
  Else
    Create_Tolerance
  End If
  Speak_Msg "标记画框  右键支持容差"
End Sub

'''////  一键拆开多行组合的文字字符  ////'''
Private Sub Batch_Combine_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Batch_Combine
  ElseIf Shift = fmCtrlMask Then
    Tools.Take_Apart_Character
  Else
    Create_Tolerance
  End If

  Speak_Msg "智能拆字"
End Sub

'''////  简单一刀切  ////'''
Private Sub Single_Line_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Single_Line_Vertical
  ElseIf Shift = fmCtrlMask Then
    Tools.Single_Line
  Else
    Tools.Single_Line_LastNode
  End If
  
  Speak_Msg "简单一刀切"
End Sub

'''////  傻瓜火车排列  ////'''
Private Sub TOP_ALIGN_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.傻瓜火车排列 3#
  ElseIf Shift = fmCtrlMask Then
    Tools.傻瓜火车排列 0#
  Else
    Tools.傻瓜火车排列 Set_Space_Width
  End If
End Sub

'''////  傻瓜阶梯排列  ////'''
Private Sub LEFT_ALIGN_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.傻瓜阶梯排列 3#
  ElseIf Shift = fmCtrlMask Then
    Tools.傻瓜阶梯排列 0#
  Else
    Tools.傻瓜阶梯排列 Set_Space_Width
  End If
End Sub


'''////  左键-多页合并一页工具   右键-批量多页居中 ////'''
Private Sub UniteOne_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.批量多页居中
  ElseIf Shift = fmCtrlMask Then
    UniteOne.Show 0
    Speak_Msg "多页合并一页"
  Else
    ' Ctrl + 鼠标  空
  End If
End Sub

'''////  Adobe AI EPS INDD PDF和CorelDRAW 缩略图工具  ////'''
Private Sub AdobeThumbnail_Click()
    Dim h As Long, r As Long
    mypath = Path & "GMS\262235.xyz\"
    App = mypath & "GuiAdobeThumbnail.exe"
    
    h = FindWindow(vbNullString, "CorelVBA 青年节 By 蘭雅sRGB")
    I = ShellExecute(h, "", App, "", mypath, 1)
End Sub

'''////  快速颜色选择  ////'''
Private Sub Quick_Color_Select_Click()
  Tools.quickColorSelect
End Sub

Private Sub Cut_Cake_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.divideVertically
  ElseIf Shift = fmCtrlMask Then
    Tools.divideHorizontally
  Else
    ' Ctrl + 鼠标  空
  End If
End Sub

'// 安全辅助线功能，三键控制，讨厌辅助线的也可以用来删除辅助线
Private Sub Safe_Guideangle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.guideangle CorelDRAW.ActiveSelectionRange, 0#   ' 右键 0距离贴紧
  ElseIf Shift = fmCtrlMask Then
    Tools.guideangle CorelDRAW.ActiveSelectionRange, 4    ' 左键安全范围 4mm
  Else
    Tools.guideangle CorelDRAW.ActiveSelectionRange, -Set_Space_Width     ' Ctrl + 鼠标左键
  End If
End Sub


'// 小工具快速启动
Private Sub Open_Calc_Click()
  Launcher.START_Calc
End Sub

Private Sub Open_Notepad_Click()
  Launcher.START_Notepad
End Sub

Private Sub terminal_Click()
  Launcher.START_GitBash
End Sub

Private Sub Video_Camera_Click()
  Launcher.START_Bandicam
End Sub

