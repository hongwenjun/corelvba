#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
  Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Option Explicit

Private Sub CommandButton1_Click()
  TextBox1.Value = "设置出血和裁切线功能目前有个想法。谁有兴趣制作漂亮的图标请联系我."
  MsgBox "请每天点击右边Logo，点击博客广告一次!" & vbNewLine & "您的支持，我才能有动力添加更多功能."
End Sub

Private Sub UI_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  ' 定义图标坐标pos
  Dim pos_x As Variant
  Dim pos_Y As Variant
  pos_x = Array(32, 110, 186, 265, 345)
  pos_Y = Array(50, 135, 215)

 ' MsgBox "图标坐标: " & X & " , " & Y
  
  If Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(0)) < 30 Then
    物件角线
  ElseIf Abs(X - pos_x(1)) < 30 And Abs(Y - pos_Y(0)) < 30 Then
    绘制矩形
  ElseIf Abs(X - pos_x(2)) < 30 And Abs(Y - pos_Y(0)) < 30 Then
    角线爬虫
  ElseIf Abs(X - pos_x(3)) < 30 And Abs(Y - pos_Y(0)) < 30 Then
    矩形拼版
  ElseIf Abs(X - pos_x(4)) < 30 And Abs(Y - pos_Y(0)) < 30 Then
    拼版角线
  End If

  If Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(1)) < 30 Then
    居中页面
  ElseIf Abs(X - pos_x(1)) < 30 And Abs(Y - pos_Y(1)) < 30 Then
    拼版标记
  ElseIf Abs(X - pos_x(2)) < 30 And Abs(Y - pos_Y(1)) < 30 Then
    智能群组
  ElseIf Abs(X - pos_x(3)) < 30 And Abs(Y - pos_Y(1)) < 30 Then
    CQL选择
  ElseIf Abs(X - pos_x(4)) < 30 And Abs(Y - pos_Y(1)) < 30 Then
    批量替换
  End If

  If Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(2)) < 30 Then
    尺寸取整
  ElseIf Abs(X - pos_x(1)) < 30 And Abs(Y - pos_Y(2)) < 30 Then
    Dim r As Long
  ElseIf Abs(X - pos_x(2)) < 30 And Abs(Y - pos_Y(2)) < 30 Then
    WebHelp
  ElseIf Abs(X - pos_x(3)) < 30 And Abs(Y - pos_Y(2)) < 30 Then
    WebHelp
  ElseIf Abs(X - pos_x(4)) < 30 And Abs(Y - pos_Y(2)) < 30 Then
    WebHelp
  End If

End Sub

Function WebHelp()
Dim h As Long, r As Long
h = FindWindow(vbNullString, "262235.xyz 老人关怀版  By 蘭雅sRGB 2022")
r = ShellExecute(h, "", "https://262235.xyz", "", "", 1)
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
  智能群组和查找.剪贴板物件替换
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
  CQL查找相同.属性选择
End Sub

Private Sub 居中页面()
  ' 实践应用: 选择物件群组,页面设置物件大小,物件页面居中
  ActiveDocument.Unit = cdrMillimeter
  Dim OrigSelection As ShapeRange, sh As Shape
  Set OrigSelection = ActiveSelectionRange
  Set sh = OrigSelection.Group

  ' MsgBox "选择物件尺寸: " & sh.SizeWidth & "x" & sh.SizeHeight
  ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
  sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
End Sub

Private Sub 尺寸取整()
  ActiveDocument.Unit = cdrMillimeter
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String, size As String
  For Each sh In shs
    size = Int(sh.SizeWidth + 0.5) & "x" & Int(sh.SizeHeight + 0.5) & "mm"
    sh.SetSize Int(sh.SizeWidth + 0.5), Int(sh.SizeHeight + 0.5)
    
    s = s & size & vbNewLine
  Next sh

  MsgBox "物件尺寸信息到剪贴板" & vbNewLine & s
  API.WriteClipBoard s
End Sub
