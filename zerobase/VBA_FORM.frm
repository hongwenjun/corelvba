VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBA_FORM 
   Caption         =   "Hello_VBA"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6180
   OleObjectBlob   =   "VBA_FORM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VBA_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AutoRotate_Click()
  Tools.�Զ���ת�Ƕ�
End Sub

Private Sub btn_autoalign_bycolumn_Click()
  autogroup("group", 1).CreateSelection
End Sub

Private Sub btn_corners_off_Click()
  Tools.corner_off
End Sub

Private Sub btn_ExpandForm_Click()
  With Me
    If .Width = 200 Then
      .Width = 260: .Height = 132
    ElseIf .Height = 132 Then
      .Height = 206
    Else
      .Width = 200: .Height = 105
    End If
  End With
End Sub

Private Sub cmd_Batch_Center_Click()
  Container.Batch_Center
End Sub

Private Sub CommandButton1_Click()
  autogroup("group", 2).CreateSelection
End Sub


Private Sub CB_AQX_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.guideangle ActiveSelectionRange, 0#   ' �Ҽ� 0��������
  ElseIf Shift = fmCtrlMask Then
    Tools.guideangle ActiveSelectionRange, 4    ' �����ȫ��Χ 4mm
  Else
    Tools.guideangle ActiveSelectionRange, -10     ' Ctrl + ������
  End If
End Sub

Private Sub CB_BZCC_Click()
  Tools.�ߴ��ע
End Sub


Private Sub CB_ECWZ_Click()
  Tools.����������� GetClipBoardString
End Sub

Private Sub CB_JDZP_Click()
  Tools.�Ƕ�תƽ
End Sub

Private Sub CB_JHDX_Click()
  Tools.��������
End Sub

Private Sub CB_make_sizes_Click()
  Tools.Make_Sizes
End Sub

Private Sub CB_PLBZ_Click()
  Tools.������ע
End Sub

Private Sub CB_PLDYJZ_Click()
  Tools.������ҳ����
End Sub

Private Sub CB_PLWZ_Click()
  Tools.������������ "CorelVBA��������"
End Sub

Private Sub CB_QZJZ_Click()
  Tools.Ⱥ�����ҳ��
End Sub


Private Sub CB_SIZESORT_Click()
    splash.Show 1
End Sub

Private Sub CB_VBA_Click()
  MsgBox "��� CorelVBA!"
End Sub

Private Sub CB_VBA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  CB_VBA.BackColor = RGB(255, 0, 0)
End Sub


Private Sub CB_ZDJD_Click()
  Tools.�Զ���ת�Ƕ�
End Sub

Private Sub CB_mirror_by_line_Click()
  Tools.�ο��߾���
End Sub


Private Sub CommandButton2_Click()
  Tools.������T
End Sub

Private Sub CommandButton3_Click()
    Dim sr As ShapeRange
    Dim shr As ShapeRange

    Set sr = ActiveSelectionRange
    Set shr = ActivePage.Shapes.All

    If sr.Shapes.Count = 0 Then
        shr.CreateSelection '���ж���
    Else
        shr.RemoveRange sr
        shr.CreateSelection '����ԭѡ��Χ�ڵĶ���
    End If
End Sub

Private Sub ExportNodePot_Click()
  Tools.ExportNodePositions
End Sub

Private Sub Image7_Click()
arrow.Show 0
Unload Me
End Sub

Private Sub Image8_Click()
    frmSelectSame.Show 0
    Unload Me
End Sub


Private Sub OneKeyToPowerClip_Click()
  Container.OneKey_ToPowerClip
End Sub

Private Sub Photo_Form_Click()
  PhotoForm.Show 0
End Sub

Private Sub BatchToPowerClip_Click()
  Container.Batch_ToPowerClip
End Sub


Private Sub Print_Page_Click()
  ActivePage.Shapes.All.Move ActivePage.CenterX - ActiveSelectionRange.CenterX, ActivePage.CenterY - ActiveSelectionRange.CenterY
  
  ' �ȼ����漸�д���
  ' Dim sr As ShapeRange, shr As ShapeRange
  ' Set sr = ActiveSelectionRange
  ' Set shr = ActivePage.Shapes.All
  
  ' X = sr.CenterX
  ' Y = sr.CenterY
  ' px = ActivePage.CenterX
  ' py = ActivePage.CenterY
  ' shr.Move px - X, py - Y
  
End Sub

Private Sub RemoveShapes_OutsideBox_Click()
  Container.Remove_OutsideBox
End Sub

Private Sub SelectOnMargin_Click()
  Container.Select_OnMargin
End Sub

Private Sub SelectOutsideBox_Click()
  Container.Select_OutsideBox
End Sub

Private Sub Set_BoxName_Click()
  Container.SetBoxName
End Sub

Private Sub SetNames_Click()
  Tools.SetNames
End Sub

Private Sub SplitSegment_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    MsgBox "�������߶Σ�Ctrl�ϲ��߶�"
  ElseIf Shift = fmCtrlMask Then
    Tools.Split_Segment
  Else
    ActiveSelection.CustomCommand "ConvertTo", "JoinCurves"
    Application.Refresh
  End If
End Sub

Private Sub Image4_Click()
    cmd_line = "Notepad  D:\����¼.txt"
    Shell cmd_line, vbNormalNoFocus
End Sub

Private Sub Image5_Click()
  Shell "Calc"
End Sub

Private Sub LevelRuler_Click()
  Tools.�Ƕ�תƽ
End Sub

Private Sub MakeSizes_Click()
  ZCOPY.Show 0
End Sub

Private Sub MirrorLine_Click()
  Tools.�ο��߾���
End Sub

Private Sub SortCount_Click()
  Tools.��������� 50
End Sub

Private Sub SwapShape_Click()
  Tools.��������
End Sub



Private Sub TESTPIC__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    TESTPIC.SpecialEffect = fmSpecialEffectSunken

End Sub
Private Sub UserForm_Click()

End Sub

Private Sub ZNQZ_Click()
  Tools.����Ⱥ��
End Sub

Private Sub ��ȡ�ı�_Click()
  AutoCutLines.AutoCutLines
End Sub

Sub ��ȡÿһ������()
    Dim txt As Object, t As Object, path As String
    Set txt = CreateObject("Scripting.FileSystemObject")
    
    Dim a
    ' ָ��·��
    path = "R:\Temp.txt"
    ' ��1����ʾֻ���򿪣���2����ʾд�룬True��ʾĿ���ļ�������ʱ�Ǵ���
    Set t = txt.OpenTextFile(path, 1, True)
    '--------------------------
    ' ��ȡÿһ�в���������ʾ����
    Do While Not t.AtEndOfStream
'        a = t.ReadLine
        a = a & t.ReadLine & vbNewLine
    TextBox1.Value = a
    Loop
    '--------------------------
    ' ���ĵ���ע�⡰notepad.exe ������пո�
    Shell "notepad.exe " & path, vbNormalFocus
    ' �ͷű���
    Set t = Nothing
    Set txt = Nothing
End Sub



Private Sub ������_Click()
 AutoCutLines.AutoCutLines
 
End Sub


Private Sub �ֶ�ƴ��_Click()
  ArrangeForm.Show 0
End Sub

Private Sub �㷨����_Click()
  ChatGPT.��������
End Sub

Private Sub Z������_Click()
    ChatGPT.Z������
End Sub

Private Sub U������_Click()
  ChatGPT.��ʽU������
End Sub
