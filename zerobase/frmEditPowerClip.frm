VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditPowerClip 
   Caption         =   "������ݵ���"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3510
   OleObjectBlob   =   "frmEditPowerClip.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmEditPowerClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim xzbj As Boolean
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call commdanliu(Lab001)
    Call commdanliu(Lab002)
    Call commdanliu(Lab003)
    Call commdanliu(Lab004)
    Call commdanliu(Lab005)
    Call commdanliu(Lab006)
    Call commdanliu(Lab007)
    Call commdanliu(Lab008)
End Sub
Private Sub Lab001_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call anliumove(Lab001)
End Sub
Private Sub Lab002_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call anliumove(Lab002)
End Sub
Private Sub Lab003_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call anliumove(Lab003)
End Sub
Private Sub Lab004_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call anliumove(Lab004)
End Sub
Private Sub Lab005_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call anliumove(Lab005)
End Sub
Private Sub Lab006_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call anliumove(Lab006)
End Sub
Private Sub Lab007_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call anliumove(Lab007)
End Sub
Private Sub Lab008_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Call anliumove(Lab008)
End Sub
Private Sub Lab001_Click()
    BeginOpt "��ȡ���п�����"
    Container.Extractall (IIf(CheckBox1.Value, True, False))
    EndOpt
End Sub
Private Sub Lab002_Click()
    BeginOpt "��ղ��п�"
    Container.Emptyall
    EndOpt
End Sub
Private Sub Lab003_Click()
    BeginOpt "��������������"
        Container.Bilingtznr
    EndOpt
End Sub
Private Sub Lab004_Click()
    BeginOpt "���������"
        Container.Bilintianchun
    EndOpt
End Sub
Private Sub Lab005_Click()
    BeginOpt "��չ���"
    Container.Qiangzhitianmian
    EndOpt
End Sub
Private Sub Lab006_Click()
    BeginOpt "������ȷ�ü�"
    Container.Lockall True
    EndOpt
End Sub
Private Sub Lab007_Click()
    BeginOpt "������ȷ�ü�"
        Container.Lockall False
    EndOpt
End Sub
Private Sub Lab008_Click()
    BeginOpt "���ݾ���"
    Container.CenterToPC
    EndOpt
End Sub
Private Sub txtNilai_Change()
   Dim i As Integer
   Dim s As String
   With txtNilai
      For i = 1 To VBA.Len(.text)
           s = VBA.Mid(.text, i, 1)
            Select Case s
              Case ".", "0" To "9"
              Case Else
               .text = VBA.Replace(.text, s, "")
            End Select
         Next
     End With
End Sub
Private Sub SpinButton1_SpinUp()
    txtNilai.text = VBA.CStr(txtNilai.Value + 1)
End Sub
Private Sub SpinButton1_SpinDown()
    If txtNilai.Value <= 1 Then Exit Sub
    txtNilai.text = VBA.CStr(txtNilai.Value - 1)
End Sub
Private Sub UserForm_Initialize()
    If Strbjini = "" Then Strbjini = VBA.GetSetting(xylAppName, xylSection, "Apppath"): BJAppLJ = Strbjini & "\DaTa\dat\"
    If GetmdbValue(BJAppLJ & "xylTools.ini", "Form", "rqtzFr_l", "") <> "" Then
        Me.StartUpPosition = 0
        Me.Left = GetmdbValue(BJAppLJ & "xylTools.ini", "Form", "rqtzFr_l", "")
        Me.Top = GetmdbValue(BJAppLJ & "xylTools.ini", "Form", "rqtzFr_t", "")
    End If
    Call AddStroyComandBox(Me.cboUnit, "����,����,Ӣ��,����")
    Me.cboUnit.text = GetmdbValue(BJAppLJ & "xylTools.ini", "Rongqibjtz", "��λ", "����")
    xzbj = False
    cboUnit.Enabled = False
    txtNilai.Enabled = False
    SpinButton1.Enabled = False
    spnPositionX.Enabled = False
    spnPositionY.Enabled = False
    spnZoom.Enabled = False
    spnRotate.Enabled = False
    Me.Tis.BackColor = RGB(0, 147, 222)
    Me.Tis.ForeColor = RGB(255, 255, 255)
    Me.Tis.Caption = "  ����ѡ��һ����������������"
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SetmdbValue BJAppLJ & "xylTools.ini", "Form", "rqtzFr_l", frmEditPowerClip.Left
    SetmdbValue BJAppLJ & "xylTools.ini", "Form", "rqtzFr_t", frmEditPowerClip.Top
    SetmdbValue BJAppLJ & "xylTools.ini", "Rongqibjtz", "��λ", Me.cboUnit.text
End Sub
Sub getShapeByUser()
re:
    Dim doc As Document, retval As Long
    Dim x As Double, Y As Double, Shift As Long
    Dim o_seleksi As ShapeRange
    Set doc = ActiveDocument
    doc.ReferencePoint = cdrCenter
    retval = doc.GetUserClick(x, Y, Shift, 10, True, cdrCursorPick)
    doc.ActivePage.SelectShapesAtPoint x, Y, True
    Dim SC As Shape
    Dim sp As PowerClip
    Set SC = ActiveShape
    If SC Is Nothing Then xzbj = False: Me.Show: Exit Sub
    Set sp = SC.PowerClip
    If sp Is Nothing Then
        AutoMsgbox "ѡ�������������" & vbCrLf & "��������ѡ�񣬻����հ״��˳���", vbCritical, "��ӡ����ʾ": GoTo re
    Else
        If sp.Shapes.Count = 0 Then
            AutoMsgbox "����Ϊ�գ�" & vbCrLf & "��������ѡ�񣬻����հ״��˳���", vbCritical, "��ӡ����ʾ": GoTo re
        End If
    End If
    xzbj = True
End Sub
Sub doAction(ByVal doAction As String, Optional ByVal bolUp As Boolean = False)
    doAction = VBA.LCase(doAction)
    ActiveDocument.ReferencePoint = cdrCenter
    If cboUnit.ListIndex = 0 Then
        ActiveDocument.Unit = cdrMillimeter
    ElseIf cboUnit.ListIndex = 1 Then
        ActiveDocument.Unit = cdrCentimeter
    ElseIf cboUnit.ListIndex = 2 Then
        ActiveDocument.Unit = cdrInch
    ElseIf cboUnit.ListIndex = 3 Then
        ActiveDocument.Unit = cdrPixel
    End If '
    Dim setNilai As Double
    setNilai = CDbl(txtNilai.Value)
    If bolUp = False Then setNilai = -setNilai
    Dim s As Shape, sr As ShapeRange
    Set sr = ActiveSelectionRange
    For Each s In sr
        Call checkPowerClip(s, doAction, setNilai, bolUp)
    Next s
End Sub
Private Function checkPowerClip(s As Shape, ByVal doAction As String, ByVal setNilai As Double, ByVal bolUp As Boolean)
    Dim pwc As PowerClip, sr As ShapeRange
    If Not s.PowerClip Is Nothing Then
        Set pwc = s.PowerClip
        Set sr = pwc.Shapes.FindShapes
        If doAction = "position_x" Then
            sr.PositionX = sr.PositionX + setNilai
        ElseIf doAction = "position_y" Then
            sr.PositionY = sr.PositionY + setNilai
        ElseIf doAction = "rotate" Then
            sr.Rotate setNilai
        ElseIf doAction = "zoom" Then
            sr.Stretch sr.SizeWidth / (sr.SizeWidth + setNilai)
        End If
    End If
End Function
Private Sub cmdPickObject_Click()
    Me.Hide
    Call getShapeByUser
    If xzbj = True Then
       Me.Tis.Caption = "  ��������ѡ��һ������������"
       If cmdPickObject.ControlTipText = "ѡ������" Then
          cboUnit.Enabled = True
          txtNilai.Enabled = True
          SpinButton1.Enabled = True
          spnPositionX.Enabled = True
          spnPositionY.Enabled = True
          spnZoom.Enabled = True
          spnRotate.Enabled = True
       End If
       Me.Show
       cmdPickObject.ControlTipText = "����ѡ��һ������"
    End If
End Sub
Private Sub spnPositionX_SpinDown()
    Call doAction("position_x", False)
End Sub
Private Sub spnPositionX_SpinUp()
    Call doAction("position_x", True)
End Sub
Private Sub spnPositionY_SpinDown()
    Call doAction("position_y", False)
End Sub
Private Sub spnPositionY_SpinUp()
    Call doAction("position_y", True)
End Sub
Private Sub spnRotate_SpinUp()
    Call doAction("rotate", False)
End Sub
Private Sub spnRotate_SpinDown()
    Call doAction("rotate", True)
End Sub
Private Sub spnZoom_SpinUp()
    Call doAction("zoom", False)
End Sub
Private Sub spnZoom_SpinDown()
    Call doAction("zoom", True)
End Sub
Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
     cmdPickObject.SpecialEffect = fmSpecialEffectEtched
End Sub
Private Sub cmdPickObject_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
   cmdPickObject.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub cmdPickObject_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    cmdPickObject.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub cmdPickObject_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = 0 Then
        cmdPickObject.SpecialEffect = fmSpecialEffectRaised
    ElseIf Button = 1 Then
        cmdPickObject.SpecialEffect = fmSpecialEffectSunken
    End If
End Sub
