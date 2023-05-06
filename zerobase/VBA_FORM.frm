VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBA_FORM 
   Caption         =   "Hello_VBA"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6180
   OleObjectBlob   =   "VBA_FORM.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "VBA_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AutoRotate_Click()
  Tools.自动旋转角度
End Sub

Private Sub btn_autoalign_bycolumn_Click()
  autogroup("group", 1).CreateSelection
End Sub

Private Sub btn_corners_off_Click()
  Tools.corner_off
End Sub

Private Sub CommandButton1_Click()
  autogroup("group", 2).CreateSelection
End Sub


Private Sub CB_AQX_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then
    Tools.guideangle ActiveSelectionRange, 0#   ' 右键 0距离贴紧
  ElseIf Shift = fmCtrlMask Then
    Tools.guideangle ActiveSelectionRange, 4    ' 左键安全范围 4mm
  Else
    Tools.guideangle ActiveSelectionRange, -10     ' Ctrl + 鼠标左键
  End If
End Sub

Private Sub CB_BZCC_Click()
  Tools.尺寸标注
End Sub


Private Sub CB_ECWZ_Click()
  Tools.填入居中文字 GetClipBoardString
End Sub

Private Sub CB_JDZP_Click()
  Tools.角度转平
End Sub

Private Sub CB_JHDX_Click()
  Tools.交换对象
End Sub

Private Sub CB_make_sizes_Click()
  Tools.Make_Sizes
End Sub

Private Sub CB_PLBZ_Click()
  Tools.批量标注
End Sub

Private Sub CB_PLDYJZ_Click()
  Tools.批量多页居中
End Sub

Private Sub CB_PLWZ_Click()
  Tools.批量居中文字 "CorelVBA批量文字"
End Sub

Private Sub CB_QZJZ_Click()
  Tools.群组居中页面
End Sub


Private Sub CB_SIZESORT_Click()
    splash.Show 1
End Sub

Private Sub CB_VBA_Click()
  MsgBox "你好 CorelVBA!"
End Sub

Private Sub CB_VBA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  CB_VBA.BackColor = RGB(255, 0, 0)
End Sub


Private Sub CB_ZDJD_Click()
  Tools.自动旋转角度
End Sub

Private Sub CB_mirror_by_line_Click()
  Tools.参考线镜像
End Sub


Private Sub CommandButton2_Click()
  Tools.服务器T
End Sub

Private Sub CommandButton3_Click()
    Dim sr As ShapeRange
    Dim shr As ShapeRange

    Set sr = ActiveSelectionRange
    Set shr = ActivePage.Shapes.All

    If sr.Shapes.Count = 0 Then
        shr.CreateSelection '所有对象
    Else
        shr.RemoveRange sr
        shr.CreateSelection '不在原选择范围内的对象
    End If
End Sub

Private Sub ExportNodePot_Click()
  Tools.ExportNodePositions
End Sub

Private Sub Photo_Form_Click()
  PhotoForm.Show 0
End Sub

Private Sub SetNames_Click()
  Tools.SetNames
End Sub

Private Sub SplitSegment_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then
    MsgBox "左键拆分线段，Ctrl合并线段"
  ElseIf Shift = fmCtrlMask Then
    Tools.Split_Segment
  Else
    ActiveSelection.CustomCommand "ConvertTo", "JoinCurves"
    Application.Refresh
  End If
End Sub

Private Sub Image4_Click()
    cmd_line = "Notepad  D:\备忘录.txt"
    Shell cmd_line, vbNormalNoFocus
End Sub

Private Sub Image5_Click()
  Shell "Calc"
End Sub

Private Sub LevelRuler_Click()
  Tools.角度转平
End Sub

Private Sub MakeSizes_Click()
  ZCOPY.Show 0
End Sub

Private Sub MirrorLine_Click()
  Tools.参考线镜像
End Sub

Private Sub SortCount_Click()
  Tools.按面积排列 50
End Sub

Private Sub SwapShape_Click()
  Tools.交换对象
End Sub


Private Sub ZNQZ_Click()
  Tools.智能群组
End Sub

Private Sub 读取文本_Click()
  AutoCutLines.AutoCutLines
End Sub

Sub 读取每一行数据()
    Dim txt As Object, t As Object, path As String
    Set txt = CreateObject("Scripting.FileSystemObject")
    
    Dim a
    ' 指定路径
    path = "R:\Temp.txt"
    ' “1”表示只读打开，“2”表示写入，True表示目标文件不存在时是创建
    Set t = txt.OpenTextFile(path, 1, True)
    '--------------------------
    ' 读取每一行并把内容显示出来
    Do While Not t.AtEndOfStream
'        a = t.ReadLine
        a = a & t.ReadLine & vbNewLine
    TextBox1.Value = a
    Loop
    '--------------------------
    ' 打开文档，注意“notepad.exe ”最后有空格
    Shell "notepad.exe " & path, vbNormalFocus
    ' 释放变量
    Set t = Nothing
    Set txt = Nothing
End Sub



Private Sub 裁切线_Click()
 AutoCutLines.AutoCutLines
 
End Sub


Private Sub 手动拼版_Click()
  ArrangeForm.Show 0
End Sub

Private Sub 算法计算_Click()
  ChatGPT.计算行列
End Sub

Private Sub Z序排列_Click()
    ChatGPT.Z序排列
End Sub

Private Sub U序排列_Click()
  ChatGPT.正式U序排列
End Sub
