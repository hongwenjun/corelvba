VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContainerForm 
   Caption         =   "Everything Object as Select        github.com/hongwenjun"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5730
   OleObjectBlob   =   "ContainerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContainerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
  LNG_CODE = API.GetLngCode
  Set_BoxName.Caption = i18n("Define as", LNG_CODE) & vbCrLf & i18n("Select", LNG_CODE)
  LabelTOL.Caption = i18n("TOL:", LNG_CODE) & GlobalUserData("Tolerance", 1)

  Me.Caption = i18n("Everything Object as Select", LNG_CODE) & "        github.com/hongwenjun"
  Init_Translations Me, LNG_CODE
  
  txtInfo.text = i18n("Usage", LNG_CODE) & ": A->Left B->Right C->Ctrl"
  
End Sub

Private Sub Set_BoxName_Click()
  Container.SetBoxName
  Create_Tolerance
  LabelTOL.Caption = i18n("TOL:", LNG_CODE) & GlobalUserData("Tolerance", 1)

End Sub

Private Sub RemoveShapes_OutsideBox_Click()
  If GlobalUserData.Exists("Tolerance", 1) Then text = GlobalUserData("Tolerance", 1)
  Container.Remove_OutsideBox Val(text)
End Sub

Private Sub SelectOnMargin_Click()
  Container.Select_SideBox cdrOnMarginOfShape
End Sub

Private Sub AreaSelect_Click()
  Container.Select_SideBox cdrOnMarginOfShape + cdrInsideShape
End Sub

Private Sub SelectOutsideBox_Click()
  Container.Select_SideBox cdrOutsideShape
End Sub

Private Sub SelectInsideBox_Click()
  Container.Select_SideBox cdrInsideShape
End Sub

Private Sub OneKeyToPowerClip_Click()
  Container.OneKey_ToPowerClip
End Sub

Private Sub BatchToPowerClip_Click()
  Container.Batch_ToPowerClip
End Sub

Private Sub Select_byBlendGroup_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If GlobalUserData.Exists("Tolerance", 1) Then text = GlobalUserData("Tolerance", 1)

  If Button = 2 Then
    Container.Select_by_BlendGroup Val(text)
    ContainerForm.Caption = i18n("If you like this feature, please visit.", LNG_CODE) & "  github.com/hongwenjun"
    Exit Sub
  ElseIf Shift = fmCtrlMask Then
    Container.Select_Quick_BlendGroup Val(text)
    LabelTOL.Caption = i18n("Right Click is Better", LNG_CODE)
  Else
     ' Ctrl + ЪѓБъ
  End If
End Sub

Private Sub MADD_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Store_Instruction 2, "add"
  ElseIf Shift = fmCtrlMask Then
    Store_Instruction 1, "add"
  Else
    Store_Instruction 3, "add"
  End If
  txtInfo.text = StoreCount
End Sub

Private Sub MSUB_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Store_Instruction 2, "sub"
  ElseIf Shift = fmCtrlMask Then
    Store_Instruction 1, "sub"
  Else
    Store_Instruction 3, "sub"
  End If
  txtInfo.text = StoreCount
End Sub

Private Sub MRLW_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Store_Instruction 2, "lw"
  ElseIf Shift = fmCtrlMask Then
    Store_Instruction 1, "lw"
  Else
    Store_Instruction 3, "lw"
  End If
  txtInfo.text = StoreCount
End Sub

Private Sub MZERO_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    Store_Instruction 2, "zero"
  ElseIf Shift = fmCtrlMask Then
    Store_Instruction 1, "zero"
  Else
    Store_Instruction 3, "zero"
  End If
  txtInfo.text = StoreCount
End Sub
