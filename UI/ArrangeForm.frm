VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArrangeForm 
   Caption         =   "蘭雅sRGB 自动拼版 │ 嘉盟赞助"
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
'// 用户窗口初始化
Private Sub UserForm_Initialize()
  ActiveDocument.Unit = cdrMillimeter
  Dim sr As ShapeRange
  Dim ls, hs, lj, hj, pw, ph As Double
  
  pw = ActiveDocument.Pages.First.SizeWidth
  ph = ActiveDocument.Pages.First.SizeHeight
  TextBox1.text = 2
  TextBox2.text = 5
  TextBox3.text = 0
  TextBox4.text = 0
  
  Set sr = ActiveSelectionRange
  If sr.Count > 0 Then
    ls = Int(sr.SizeWidth + 0.5)
    hs = Int(sr.SizeHeight + 0.5)
    Label_Size.Caption = "尺寸: " & ls & "×" & hs & "mm"
    
    lj = Int(pw / ls)
    hj = Int(ph / hs)
    
    Dim jh, jl, t As Double
    jl = Int(pw / hs)
    jh = Int(ph / ls)
    
'//  Debug.Print lj, hj, jl, jh
    If jh * jl > hj * lj Then
      lj = jl
      hj = jh
      If lj * ls > pw Or hj * hs > ph Then
        t = lj
        lj = hj
        hj = t
      End If
    End If
    
    
    TextBox1.text = lj
    TextBox2.text = hj
  End If
End Sub

Private Sub CommandButton1_Click()
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim ls, hs As Integer: Dim lj, hj As Double
  Dim matrix As Variant: Dim sr As ShapeRange
  
  ls = Val(TextBox1.text)
  hs = Val(TextBox2.text)
  lj = Val(TextBox3.text)
  hj = Val(TextBox4.text)
  matrix = Array(ls, hs, lj, hj)
  
  Set sr = ActiveSelectionRange

  If ls * hs = 0 Then Exit Sub
  If ls = 1 Or hs = 1 Then
    arrange_Clone_one matrix, sr
    GoTo ErrorHandler
  End If
  
  '// 代码运行时关闭窗口刷新
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  '// 拼版矩阵
  arrange_Clone matrix, sr
  Unload Me
  
ErrorHandler:
  API.EndOpt
End Sub

'// 拼版矩阵  matrix = Array(ls, hs, lj, hj)
Private Function arrange_Clone(matrix As Variant, sr As ShapeRange)
  ls = matrix(0): hs = matrix(1)
  lj = matrix(2): hj = matrix(3)
  X = sr.SizeWidth: Y = sr.SizeHeight
  Set s1 = sr '// Set s1 = sr.Clone
  '// StepAndRepeat 方法在范围内创建多个形状副本
  
'//  Set dup1 = s1.StepAndRepeat(ls - 1, x + lj, 0#)
'//  Set dup2 = ActiveDocument.CreateShapeRangeFromArray(dup1, s1).StepAndRepeat(hs - 1, 0#, -(Y + hj))

Set dup1 = s1.StepAndRepeat(hs - 1, 0#, -(Y + hj))
Set dup2 = ActiveDocument.CreateShapeRangeFromArray(dup1, s1).StepAndRepeat(ls - 1, X + lj, 0#)

  '// s1.Delete
End Function

Private Function arrange_Clone_one(matrix As Variant, sr As ShapeRange)
  ls = matrix(0): hs = matrix(1)
  lj = matrix(2): hj = matrix(3)
  X = sr.SizeWidth: Y = sr.SizeHeight
  Set s1 = sr '// Set s1 = sr.Clone
  '// StepAndRepeat 方法在范围内创建多个形状副本
  If ls > 1 Then
    Set dup1 = s1.StepAndRepeat(ls - 1, X + lj, 0#)
  Else
    Set dup1 = s1
  End If
  If hs > 1 Then
    Set dup2 = ActiveDocument.CreateShapeRangeFromArray(dup1, s1).StepAndRepeat(hs - 1, 0#, -(Y + hj))
  End If
  '// s1.Delete
End Function

