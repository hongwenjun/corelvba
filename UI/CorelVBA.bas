#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    
#Else
  Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private switch As Boolean

Private Sub Close_Icon_Click()
  Unload Me    ' 关闭
End Sub

Private Sub ToolBar_show_Click()
  Unload Me
  ToolBar.Show 0
End Sub

Private Sub UserForm_Initialize()
  Dim IStyle As Long
  Dim Hwnd As Long
  
  Hwnd = FindWindow("ThunderDFrame", Me.Caption)

  IStyle = GetWindowLong(Hwnd, GWL_STYLE)
  IStyle = IStyle And Not WS_CAPTION
  SetWindowLong Hwnd, GWL_STYLE, IStyle
  DrawMenuBar Hwnd
  IStyle = GetWindowLong(Hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
  SetWindowLong Hwnd, GWL_EXSTYLE, IStyle

  With Me
  '  .StartUpPosition = 0
  '  .Left = 500
  '  .Top = 200
    .Width = 385.5
    .Height = 271.45
  End With
  
End Sub

Private Sub LOGO_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button Then
    mX = x
    mY = y
  End If
End Sub

Private Sub LOGO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button Then
    Me.Left = Me.Left - mx + x
    Me.Top = Me.Top - my + y
  End If
End Sub

Private Sub CommandButton1_Click()
  MsgBox "请给我支持!" & vbNewLine & "您的支持，我才能有动力添加更多功能." & vbNewLine & "蘭雅CorelVBA青年节版公测" & vbNewLine & "coreldrawvba插件交流群  8531411"
End Sub

Private Sub UI_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

  ' 定义图标坐标pos
  Dim pos_x As Variant
  Dim pos_y As Variant
  pos_x = Array(32, 110, 186, 265, 345)
  pos_y = Array(50, 135, 215)

  If Abs(x - pos_x(0)) < 30 And Abs(y - pos_y(0)) < 30 Then
    物件角线
  ElseIf Abs(x - pos_x(1)) < 30 And Abs(y - pos_y(0)) < 30 Then
    绘制矩形
  ElseIf Abs(x - pos_x(2)) < 30 And Abs(y - pos_y(0)) < 30 Then
    角线爬虫
  ElseIf Abs(x - pos_x(3)) < 30 And Abs(y - pos_y(0)) < 30 Then
    矩形拼版
  ElseIf Abs(x - pos_x(4)) < 30 And Abs(y - pos_y(0)) < 30 Then
    拼版角线
  End If

  If Abs(x - pos_x(0)) < 30 And Abs(y - pos_y(1)) < 30 Then
    Tools.居中页面
  ElseIf Abs(x - pos_x(1)) < 30 And Abs(y - pos_y(1)) < 30 Then
    拼版标记
  ElseIf Abs(x - pos_x(2)) < 30 And Abs(y - pos_y(1)) < 30 Then
    智能群组
  ElseIf Abs(x - pos_x(3)) < 30 And Abs(y - pos_y(1)) < 30 Then
    CQL选择
  ElseIf Abs(x - pos_x(4)) < 30 And Abs(y - pos_y(1)) < 30 Then
    批量替换
  End If

  If Abs(x - pos_x(0)) < 30 And Abs(y - pos_y(2)) < 30 Then
    Tools.尺寸取整
  ElseIf Abs(x - pos_x(1)) < 30 And Abs(y - pos_y(2)) < 30 Then
    Tools.TextShape_ConvertToCurves
  ElseIf Abs(x - pos_x(2)) < 30 And Abs(y - pos_y(2)) < 30 Then
    Dim h As Long, r As Long
    mypath = Path & "GMS\262235.xyz\"
    app = mypath & "GuiAdobeThumbnail.exe"
    
    h = FindWindow(vbNullString, "CorelVBA 青年节 By 蘭雅sRGB")
    i = ShellExecute(h, "", app, "", mypath, 1)

  ElseIf Abs(x - pos_x(3)) < 30 And Abs(y - pos_y(2)) < 30 Then
    If switch Then
      switch = Not switch
      Tools.傻瓜火车排列
    Else
      switch = Not switch
      Tools.傻瓜阶梯排列
    End If
    
  ElseIf Abs(x - pos_x(4)) < 30 And Abs(y - pos_y(2)) < 30 Then
    学习CorelVBA实验室
  End If

  
  If Abs(x - 210) < 30 And Abs(y - 261) < 8 Then
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
  i = GMSManager.RunMacro("CorelDRAW_VBA", "学习CorelVBA.start")
End Sub
