VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Toolbar 
   Caption         =   "Toolbar"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   OleObjectBlob   =   "Toolbar.frx":0000
End
Attribute VB_Name = "Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

Private Const Github_Version = 1

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
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
  SaveSetting "LYVBA", "Settings", "SpeakHelp", "0"
  MsgBox "�����֧��!" & vbNewLine & "����֧�֣��Ҳ����ж�����Ӹ��๦��." & vbNewLine & "�m��CorelVBA���� ������ѿ�Դ"
End Sub

Private Sub I18N_LNG_Click()
  LNG_CODE = Val(GetSetting("LYVBA", "Settings", "I18N_LNG", "1033"))
  If LNG_CODE = 1033 Then
    LNG_CODE = 2052
  Else
    LNG_CODE = 1033
  End If
  SaveSetting "LYVBA", "Settings", "I18N_LNG", LNG_CODE
  MsgBox "��Ӣ�������л���ɣ����������!", vbOKOnly, "����VBA�������"
End Sub

Private Sub UserForm_Initialize()
  Dim IStyle As Long
  Dim hwnd As Long
  
  hwnd = FindWindow("ThunderDFrame", Me.Caption)

  IStyle = GetWindowLong(hwnd, GWL_STYLE)
  IStyle = IStyle And Not WS_CAPTION
  SetWindowLong hwnd, GWL_STYLE, IStyle
  DrawMenuBar hwnd
  IStyle = GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
  SetWindowLong hwnd, GWL_EXSTYLE, IStyle
  
With Me
  .StartUpPosition = 0
  .Left = Val(GetSetting("LYVBA", "Settings", "Left", "400"))  ' ���ù�����λ��
  .Top = Val(GetSetting("LYVBA", "Settings", "Top", "55"))
  .Height = 30
  .width = 336
End With

  OutlineKey = True
  OptKey = True

  ' ��ȡ��������
  Bleed.text = API.GetSet("Bleed")
  Line_len.text = API.GetSet("Line_len")
  Outline_Width.text = GetSetting("LYVBA", "Settings", "Outline_Width", "0.2")
  
  UIFile = Path & "GMS\LYVBA\" & HDPI.GetHDPIPercentage & "\ToolBar.jpg"
  If API.ExistsFile_UseFso(UIFile) Then
    UI.Picture = LoadPicture(UIFile)   '��UIͼ
    Set pic1 = LoadPicture(UIFile)
  End If

  UIL = Path & "GMS\LYVBA\ToolBar1.jpg"
  If API.ExistsFile_UseFso(UIL) Then
    Set pic2 = LoadPicture(UIL)
    UIL_Key = True
  End If

  ' ����͸��, ��С��ֻ��ʾһ��ͼ��
  #If VBA7 Then
    MakeUserFormTransparent Me, RGB(26, 22, 35)
  #Else
  ' CorelDRAW X4 / Windows7 ���ùر�͸��
  #End If
End Sub

Private Sub UI_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  UI.Visible = False
  If Y > 1 And Y < 16 And UIL_Key Then
    UI.Picture = pic2
  ElseIf Y > 16 And UIL_Key Then
    UI.Picture = pic1
  End If
  UI.Visible = True

  ' Debug.Print X & " , " & Y
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button Then
        mx = X: my = Y
    End If
    
  With Me
    .Height = 30
  End With

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button Then
    Me.Left = Me.Left - mx + X
    Me.Top = Me.Top - my + Y
  End If
End Sub

Private Sub LOGO_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Abs(X - 14) < 14 And Abs(Y - 14) < 14 And Button = 2 Then
    Me.width = 336
    OPEN_UI_BIG.Left = 322
    UI.Visible = True
    LOGO.Visible = False
    TOP_ALIGN_BT.Visible = False
    LEFT_ALIGN_BT.Visible = False
    Exit Sub
  ElseIf Shift = fmCtrlMask Then
      mx = X: my = Y
  Else
    Unload Me   ' Ctrl + ��� �رչ���
  End If
End Sub

Private Sub LOGO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button Then
    Me.Left = Me.Left - mx + X
    Me.Top = Me.Top - my + Y
  End If
End Sub

Private Sub UI_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Dim c As New Color
  ' ����ͼ������pos
  Dim pos_x As Variant, pos_y As Variant
  pos_y = Array(14)
  pos_x = Array(14, 41, 67, 94, 121, 148, 174, 201, 228, 254, 281, 308, 334, 361, 388, 415, 441, 468, 495)

  '// ����Ctrl���������ȴ����߹���
  If Shift = 2 Then
    If Abs(X - pos_x(0)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ��ȫ�ߣ����������
      Tools.guideangle ActiveSelectionRange, 3    ' ��� 3mm ��Ѫ
      
    ElseIf Abs(X - pos_x(1)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// Adobe AI EPS INDD PDF��CorelDRAW ����ͼ����
      AdobeThumbnail_Click
      
    ElseIf Abs(X - pos_x(2)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ���������߶�
      Tools.Split_Segment
      
    ElseIf Abs(X - pos_x(3)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ���ܲ���
      Tools.Take_Apart_Character
      
    ElseIf Abs(X - pos_x(4)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ��ʱ��
      
    ElseIf Abs(X - pos_x(5)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ��ʱ��
      
    ElseIf Abs(X - pos_x(6)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ľͷ������Ⱥ�飬����Ⱥ��
      autogroup("group", 1).CreateSelection
      
    ElseIf Abs(X - pos_x(8)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// CTRL��չ������
      Me.Height = 30 + 45
      
    ElseIf Abs(X - pos_x(9)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      ' �ı�ת��  ���� all=1 ��֧�ֿ�ѡ��ͼ������ڵ��ı�
      ' Tools.TextShape_ConvertToCurves 1
    End If
    Exit Sub
  End If


  '// ����Ҽ� ��չ����ť����  ����������  ��Ƿ�Χ��  ����ҳ�� �ߴ�ȡ����  ��ɫ�����߱�� ��չ������  ���й���  ��չ����������
  If Button = 2 Then
    If Abs(X - pos_x(0)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ����������
      Me.width = 30: Me.Height = 30
      UI.Visible = False: LOGO.Visible = True

    ElseIf Abs(X - pos_x(1)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ����ҳ��
      Tools.Align_Page_Center

    ElseIf Abs(X - pos_x(2)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    
      If Github_Version = 1 Then
        '// ������ת������ - ���õ�ҳ���ı�
        CutLines.SelectLine_to_Cropline
      Else
        '// ��Ƿ�Χ��
        Tools.Mark_Range_Box
      End If

    ElseIf Abs(X - pos_x(3)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// ������������ߴ�����
      Tools.Size_to_Integer
    
    '//�ַֺϺϰѼ������ܰ����ϲ���һ�𣬶��嵽�Ҽ���
    ElseIf Abs(X - pos_x(4)) < 14 And Abs(Y - pos_y(0)) < 14 Then
     '// Tools.�ַֺϺ�

    ElseIf Abs(X - pos_x(5)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// �Զ�����ɫ���� �ڰ�
      AutoColorMark.Auto_ColorMark_K

    ElseIf Abs(X - pos_x(6)) < 14 And Abs(Y - pos_y(0)) < 14 Then
     '// ����Ⱥ��
      SmartGroup.Smart_Group API.Create_Tolerance
      
    ElseIf Abs(X - pos_x(7)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    If Github_Version = 1 Then
      CQL_FIND_UI.Show 0
    Else
      '// ѡ����ͬ������ǿ��
      frmSelectSame.Show 0
    End If

    ElseIf Abs(X - pos_x(8)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// �Ҽ���չ������
      Me.Height = 30 + 45
      
    ElseIf Abs(X - pos_x(9)) < 14 And Abs(Y - pos_y(0)) < 14 Then
     '// �ı�ͳ����Ϣ
     Application.FrameWork.Automation.InvokeItem "bf3bd8fe-ca26-4fe0-91b0-3b5c99786fb6"

    ElseIf Abs(X - pos_x(10)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// �Ҽ����й���
      TOP_ALIGN_BT.Visible = True
      LEFT_ALIGN_BT.Visible = True

    ElseIf Abs(X - pos_x(11)) < 14 And Abs(Y - pos_y(0)) < 14 Then
      '// �Ҽ���չ����������
      Me.Height = 30
      
    End If
    Exit Sub
  End If
  
  '// ������ ������ť����  ����������ͼ����������
  If Abs(X - pos_x(0)) < 14 And Abs(Y - pos_y(0)) < 14 Then
   '// ������: �������������
    CutLines.Batch_CutLines
    
  ElseIf Abs(X - pos_x(1)) < 14 And Abs(Y - pos_y(0)) < 14 Then
  '// ������ߴ罨������
    ClipbRectangle.Build_Rectangle
    
  ElseIf Abs(X - pos_x(2)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    If Github_Version = 1 Then
      MakeSizePlus.Show 0
    Else
      '// ������ת������ - ���õ�ҳ���ı�
      CutLines.SelectLine_to_Cropline
    End If
  ElseIf Abs(X - pos_x(3)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    '// ƴ��.Arrange
    Arrange.Arrange
    
  ElseIf Abs(X - pos_x(4)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    '// ƴ�������
    CutLines.Draw_Lines
    
  ElseIf Abs(X - pos_x(5)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    '// �Զ�����ɫ���� ��ɫ
    AutoColorMark.Auto_ColorMark
    
  ElseIf Abs(X - pos_x(6)) < 14 And Abs(Y - pos_y(0)) < 14 Then
   '// ����Ⱥ�� û�ݲ�
    SmartGroup.Smart_Group
    
  ElseIf Abs(X - pos_x(7)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    If Github_Version = 1 Then
       '// ѡ����ͬ������ǿ��
      frmSelectSame.Show 0
    Else
      CQL_FIND_UI.Show 0
    End If

  ElseIf Abs(X - pos_x(8)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    Replace_UI.Show 0
    
  ElseIf Abs(X - pos_x(9)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    ' ���ı�ת��
    Tools.TextShape_ConvertToCurves 0
    
  ElseIf Abs(X - pos_x(10)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    '// ��չ������
    Me.Height = 30 + 45
    
    Speak_Msg "���Ҽ��в�ͬ����"
    
  ElseIf Abs(X - pos_x(11)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    If Me.Height > 30 Then
      Me.Height = 30
    Else
      '// ��С��
      Me.width = 30
      Me.Height = 30
      OPEN_UI_BIG.Left = 31
      UI.Visible = False
      LOGO.Visible = True
  
      ' ���湤����λ�� Left �� Top
      SaveSetting "LYVBA", "Settings", "Left", Me.Left
      SaveSetting "LYVBA", "Settings", "Top", Me.Top
    
      Speak_Msg "�����С �Ҽ�����"
    End If
  End If

End Sub

Private Sub X_EXIT_Click()
  Unload Me    ' �ر�
End Sub

'// ��ҳ�ϲ����ߣ��Ѿ��ϲ������߹���
' Private Sub ���ö�ҳ�ϲ�����()
'  Dim value As Integer
'  value = GMSManager.RunMacro("�ϲ���ҳ����", "�ϲ���ҳ����.run")
' End Sub

'''///  ̰�����˺ͺ��湤�ߵ�  ///'''
Private Sub Cdr_Nodes_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    TSP.Nodes_To_TSP
  ElseIf Shift = fmCtrlMask Then
    TSP.CDR_TO_TSP
  Else
    ' Ctrl + ���  ��
  End If
End Sub

Private Sub Cdr_Nodes_BT_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  TSP_L1.ForeColor = RGB(0, 150, 255)
End Sub

Private Sub START_TSP_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  TSP_L2.ForeColor = RGB(0, 150, 255)
End Sub

Private Sub PATH_TO_TSP_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  TSP_L3.ForeColor = RGB(0, 150, 255)
End Sub

Private Sub TSP2DRAW_LINE_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  TSP_L4.ForeColor = RGB(0, 150, 255)
End Sub

Private Sub TSP2DRAW_LINE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    TSP.TSP_TO_DRAW_LINE
  ElseIf Shift = fmCtrlMask Then
    TSP.TSP_TO_DRAW_LINES
  Else
    ' Ctrl + ���  ��
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

'''///  Python�ű��Ͷ�ά���  ///'''
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
  MsgBox "�����֧��!" & vbNewLine & "����֧�֣��Ҳ����ж�����Ӹ��๦��." & vbNewLine & "�m��CorelVBA���� ������ѿ�Դ" _
       & vbNewLine & "Դ����ַ:" & vbNewLine & "https://github.com/hongwenjun/corelvba"
End Sub

Private Sub Settings_Click()
  If 0 < Val(Bleed.text) * Val(Line_len.text) < 100 Then
   SaveSetting "LYVBA", "Settings", "Bleed", Bleed.text
   SaveSetting "LYVBA", "Settings", "Line_len", Line_len.text
   SaveSetting "LYVBA", "Settings", "Outline_Width", Outline_Width.text
  End If

  ' ���湤����λ�� Left �� Top
  SaveSetting "LYVBA", "Settings", "Left", Me.Left
  SaveSetting "LYVBA", "Settings", "Top", Me.Top
  
  Me.Height = 30
End Sub


'''/////////  ͼ��������ҵ�����ܵ���   /////////'''

Private Sub Tools_Icon_Click()
  ' �������
  i = GMSManager.RunMacro("ZeroBase", "Hello_VBA.run")
End Sub

Private Sub Split_Segment_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    MsgBox "�������߶Σ�Ctrl�ϲ��߶�"
  ElseIf Shift = fmCtrlMask Then
    Tools.Split_Segment
  Else
    ActiveSelection.CustomCommand "ConvertTo", "JoinCurves"
    Application.Refresh
  End If
  
  Speak_Msg "����߶Σ�Ctrl�ϲ��߶�"
End Sub

'''////  CorelDRAW �� Adobe_Illustrator ������ת��  ////'''
Private Sub Adobe_Illustrator_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Dim value As Integer
  If Button = 2 Then
    value = GMSManager.RunMacro("AIClipboard", "CopyPaste.PasteAIFormat")
    Exit Sub
  End If
  
  If Button Then
    value = GMSManager.RunMacro("AIClipboard", "CopyPaste.CopyAIFormat")
    MsgBox "CorelDRAW �� Adobe_Illustrator ������ת��" & vbNewLine & "���������ƣ�����Ҽ�ճ��"
  End If
End Sub

'''////  ��ǻ��� ֧���ݲ�  ////'''
Private Sub Mark_CreateRectangle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Mark_CreateRectangle True
  ElseIf Shift = fmCtrlMask Then
    Tools.Mark_CreateRectangle False
  Else
    Create_Tolerance
  End If
  Speak_Msg "��ǻ���  �Ҽ�֧���ݲ�"
End Sub

'''////  һ���𿪶�����ϵ������ַ�  ////'''
Private Sub Batch_Combine_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Batch_Combine
  ElseIf Shift = fmCtrlMask Then
    Tools.Take_Apart_Character
  Else
    Create_Tolerance
  End If
End Sub

'''////  ��һ����  ////'''
Private Sub Single_Line_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Single_Line_Vertical
  ElseIf Shift = fmCtrlMask Then
    Tools.Single_Line
  Else
    Tools.Single_Line_LastNode
  End If
End Sub

'''////  ɵ�ϻ�����  ////'''
Private Sub TOP_ALIGN_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Simple_Train_Arrangement 3#
  ElseIf Shift = fmCtrlMask Then
    Tools.Simple_Train_Arrangement 0#
  Else
    Tools.Simple_Train_Arrangement Set_Space_Width
  End If
End Sub

'''////  ɵ�Ͻ�������  ////'''
Private Sub LEFT_ALIGN_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Simple_Ladder_Arrangement 3#
  ElseIf Shift = fmCtrlMask Then
    Tools.Simple_Ladder_Arrangement 0#
  Else
    Tools.Simple_Ladder_Arrangement Set_Space_Width
  End If
End Sub


'''////  ���-��ҳ�ϲ�һҳ����   �Ҽ�-������ҳ���� ////'''
Private Sub UniteOne_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.Batch_Align_Page_Center
  ElseIf Shift = fmCtrlMask Then
    UniteOne.Show 0
  Else
    ' Ctrl + ���  ��
  End If
End Sub

'''////  Adobe AI EPS INDD PDF��CorelDRAW ����ͼ����  ////'''
Private Sub AdobeThumbnail_Click()
    Dim h As Long, r As Long
    mypath = Path & "GMS\LYVBA\"
    App = mypath & "GuiAdobeThumbnail.exe"
    
    h = FindWindow(vbNullString, "CorelVBA ����� By �m��sRGB")
    i = ShellExecute(h, "", App, "", mypath, 1)
End Sub

'''////  ������ɫѡ��  ////'''
Private Sub Quick_Color_Select_Click()
  Tools.quickColorSelect
End Sub

Private Sub Cut_Cake_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.divideVertically
  ElseIf Shift = fmCtrlMask Then
    Tools.divideHorizontally
  Else
    ' Ctrl + ���  ��
  End If
End Sub

'// ��ȫ�����߹��ܣ��������ƣ����Ḩ���ߵ�Ҳ��������ɾ��������
Private Sub Safe_Guideangle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.guideangle ActiveSelectionRange, 0#   ' �Ҽ�0��������
  ElseIf Shift = fmCtrlMask Then
    Tools.guideangle ActiveSelectionRange, 3    ' ��� 3mm ��Ѫ
  Else
    Tools.guideangle ActiveSelectionRange, -Set_Space_Width     ' Ctrl + ������ �Զ�����
  End If
End Sub

'// ��׼�ߴ磬����Ҽ�Ctrl�������ƣ�����������ʽ
Private Sub btn_makesizes_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Make_SIZE.Show 0   ' �Ҽ�
  ElseIf Shift = fmCtrlMask Then
    #If VBA7 Then
      MakeSizePlus.Show 0
    #Else  ' X4 ʹ��
      Make_SIZE.Show 0
    #End If
  Else
    Tools.Simple_Label_Numbers  ' Ctrl + ���  ���������ֱ�ע
  End If
End Sub

'// ����תͼƬ�͵���ͼƬ�ļ�
Private Sub Photo_Form_Click()
  PhotoForm.Show 0
End Sub

'// �޸�Բ��ȱ�ǵ�ֱ��
Private Sub btn_corners_off_Click()
  Tools.corner_off
End Sub

Private Sub SortCount_Click()
  Tools.Count_byArea 30
End Sub

Private Sub LevelRuler_Click()
  Tools.Angle_to_Horizon
End Sub

Private Sub MirrorLine_Click()
  Tools.Mirror_ByGuide
End Sub

Private Sub AutoRotate_Click()
  Tools.Auto_Rotation_Angle
End Sub

Private Sub SwapShape_Click()
  Tools.Exchange_Object
End Sub


'// С���߿�������
Private Sub Open_Calc_Click()
  Launcher.START_Calc
End Sub

Private Sub Open_Notepad_Click()
  Launcher.START_Notepad
End Sub

Private Sub ImageReader_Click()
  Launcher.START_Barcode_ImageReader
End Sub

Private Sub Video_Camera_Click()
  Launcher.START_Bandicam
End Sub

Private Sub myfonts_Click()
  Launcher.START_whatthefont
End Sub

Private Sub VectorMagic_Click()
  Launcher.START_Vector_Magic
End Sub

Private Sub waifu2x_Click()
  Launcher.START_waifu2x
End Sub
