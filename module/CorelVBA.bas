Attribute VB_Name = "CORELVBA"
Public Sub Start()
  Toolbar.Show 0
'  CorelVBA.show 0
'  MsgBox "�����֧��!" & vbNewLine & "����֧�֣��Ҳ����ж�����Ӹ��๦��." & vbNewLine & "�m��CorelVBA����ڰ�" & vbNewLine & "coreldrawvba�������Ⱥ  8531411"
'  Speak_Msg "��л��ʹ�� �m��VBA����"
End Sub

Sub Start_Dimension()
  '// �ߴ��ע��ǿ��
  MakeSizePlus.Show 0
End Sub

Public Sub Init_StartButton()
  SaveSetting "LYVBA", "Settings", "StartButton", "0"
  MsgBox "Please Restart CorelDRAW!"
End Sub

