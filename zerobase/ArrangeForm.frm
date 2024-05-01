VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArrangeForm 
   Caption         =   "�m��sRGB �ֶ�ƴ�� �� ��������"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   OleObjectBlob   =   "ArrangeForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "ArrangeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
  On Error GoTo ErrorHandler
  ActiveDocument.Unit = cdrMillimeter
  Dim ls As Integer, hs As Integer
  Dim lj As Double, hj As Double
  Dim matrix As Variant
  Dim s As ShapeRange
  
  ls = Val(TextBox1.text)
  hs = Val(TextBox2.text)
  lj = Val(TextBox3.text)
  hj = Val(TextBox4.text)
  matrix = Array(ls, hs, lj, hj)
  
  Set s = ActiveSelectionRange

  If ls * hs = 0 Then Exit Sub
  If ls = 1 Or hs = 1 Then
    arrange_Clone_one matrix, s
    Exit Sub
  End If
  
  '// ��������ʱ�رմ���ˢ��
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  '// ƴ�����
  arrange_Clone matrix, s

  ActiveDocument.EndCommandGroup
  Unload Me
  
  '// ������������ָ�����ˢ��
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
  Exit Sub
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Sub

'// ƴ�����  matrix = Array(ls,hs,lj,hj)
Private Function arrange_Clone(matrix As Variant, s As ShapeRange)
  ls = matrix(0): hs = matrix(1)
  lj = matrix(2): hj = matrix(3)
  x = s.SizeWidth: Y = s.SizeHeight
  Set s1 = s.Clone
  '// StepAndRepeat �����ڷ�Χ�ڴ��������״����
  Set dup1 = s1.StepAndRepeat(ls - 1, x + lj, 0#)
  Set dup2 = ActiveDocument.CreateShapeRangeFromArray(dup1, s1).StepAndRepeat(hs - 1, 0#, -(Y + hj))
  s1.Delete
End Function

Private Function arrange_Clone_one(matrix As Variant, s As ShapeRange)
  ls = matrix(0): hs = matrix(1)
  lj = matrix(2): hj = matrix(3)
  x = s.SizeWidth: Y = s.SizeHeight
  Set s1 = s.Clone
  '// StepAndRepeat �����ڷ�Χ�ڴ��������״����
  If ls > 1 Then
    Set dup1 = s1.StepAndRepeat(ls - 1, x + lj, 0#)
  Else
    Set dup1 = s1
  End If
  If hs > 1 Then
    Set dup2 = ActiveDocument.CreateShapeRangeFromArray(dup1, s1).StepAndRepeat(hs - 1, 0#, -(Y + hj))
  End If
  s1.Delete
End Function

