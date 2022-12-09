VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CorelVBA 
   Caption         =   "CorelVBA 中秋节版 By 蘭雅sRGB 2022"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   OleObjectBlob   =   "CorelVBA.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "CorelVBA"
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
    
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private switch As Boolean

Private Sub Close_Icon_Click()
  WebHelp "https://262235.xyz/index.php/tag/vba/"
  Unload Me    ' 关闭
End Sub

Private Sub ToolBar_show_Click()
  Unload Me
  Toolbar.Show 0
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
  '  .StartUpPosition = 0
  '  .Left = 500
  '  .Top = 200
    .Width = 385.5
    .Height = 271.45
  End With
  
  UIFile = Path & "GMS\262235.xyz\UI.jpg"
  If API.ExistsFile_UseFso(UIFile) Then
    UI.Picture = LoadPicture(UIFile)   '换UI图
  End If
End Sub

Private Sub LOGO_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button Then
    mx = x
    my = Y
  End If
End Sub

Private Sub LOGO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button Then
    Me.Left = Me.Left - mx + x
    Me.Top = Me.Top - my + Y
  End If
End Sub

Private Sub About_Cmd_Click()
  MsgBox "请给我支持!" & vbNewLine & "您的支持，我才能有动力添加更多功能." & vbNewLine & "蘭雅CorelVBA中秋节版" & vbNewLine & "coreldrawvba插件交流群  8531411"
End Sub

Private Sub UI_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

  ' 定义图标坐标pos
  Dim pos_x As Variant
  Dim pos_y As Variant
  pos_x = Array(32, 110, 186, 265, 345)
  pos_y = Array(50, 135, 215)

  If Abs(x - pos_x(0)) < 30 And Abs(Y - pos_y(0)) < 30 Then
    物件角线
  ElseIf Abs(x - pos_x(1)) < 30 And Abs(Y - pos_y(0)) < 30 Then
    绘制矩形
  ElseIf Abs(x - pos_x(2)) < 30 And Abs(Y - pos_y(0)) < 30 Then
    角线爬虫
  ElseIf Abs(x - pos_x(3)) < 30 And Abs(Y - pos_y(0)) < 30 Then
    矩形拼版
  ElseIf Abs(x - pos_x(4)) < 30 And Abs(Y - pos_y(0)) < 30 Then
    拼版角线
  End If

  If Abs(x - pos_x(0)) < 30 And Abs(Y - pos_y(1)) < 30 Then
    Tools.居中页面
  ElseIf Abs(x - pos_x(1)) < 30 And Abs(Y - pos_y(1)) < 30 Then
    拼版标记
  ElseIf Abs(x - pos_x(2)) < 30 And Abs(Y - pos_y(1)) < 30 Then
    智能群组
  ElseIf Abs(x - pos_x(3)) < 30 And Abs(Y - pos_y(1)) < 30 Then
    CQL选择
  ElseIf Abs(x - pos_x(4)) < 30 And Abs(Y - pos_y(1)) < 30 Then
    批量替换
  End If

  If Abs(x - pos_x(0)) < 30 And Abs(Y - pos_y(2)) < 30 Then
    Tools.尺寸取整
  ElseIf Abs(x - pos_x(1)) < 30 And Abs(Y - pos_y(2)) < 30 Then
    Tools.TextShape_ConvertToCurves
  ElseIf Abs(x - pos_x(2)) < 30 And Abs(Y - pos_y(2)) < 30 Then
    Dim h As Long, r As Long
    mypath = Path & "GMS\262235.xyz\"
    App = mypath & "GuiAdobeThumbnail.exe"
    
    h = FindWindow(vbNullString, "CorelVBA 青年节 By 蘭雅sRGB")
    I = ShellExecute(h, "", App, "", mypath, 1)

  ElseIf Abs(x - pos_x(3)) < 30 And Abs(Y - pos_y(2)) < 30 Then
    If switch Then
      switch = Not switch
      Tools.傻瓜火车排列 0#
    Else
      switch = Not switch
      Tools.傻瓜阶梯排列 0#
    End If
    
  ElseIf Abs(x - pos_x(4)) < 30 And Abs(Y - pos_y(2)) < 30 Then
    学习CorelVBA实验室
  End If

  
  If Abs(x - 210) < 30 And Abs(Y - 261) < 8 Then
    WebHelp "https://262235.xyz/index.php/tag/vba/"
  End If

End Sub

Function WebHelp(url As String)
Dim h As Long, r As Long
h = FindWindow(vbNullString, "CorelVBA 青年节 By 蘭雅sRGB")
r = ShellExecute(h, "", url, "", "", 1)
End Function


Private Sub 绘制矩形()
  剪贴板尺寸建立矩形.start
End Sub

Private Sub 角线爬虫()
  裁切线.SelectLine_to_Cropline
End Sub

Private Sub 矩形拼版()
  拼版裁切线.arrange
End Sub

Private Sub 批量替换()
  CorelVBA.Hide
  Replace_UI.Show 0
End Sub

Private Sub 拼版标记()
  自动中线色阶条.Auto_ColorMark
End Sub

Private Sub 拼版角线()
  拼版裁切线.Cut_lines
End Sub

Private Sub 物件角线()
  裁切线.start
End Sub

Private Sub 智能群组()
  智能群组和查找.智能群组
End Sub

Private Sub CQL选择()
  CorelVBA.Hide
  CQL_FIND_UI.Show 0
End Sub


Private Sub 学习CorelVBA实验室()
  CorelVBA.Hide
  ' 调用语句
  I = GMSManager.RunMacro("CorelDRAW_VBA", "学习CorelVBA.start")
End Sub
