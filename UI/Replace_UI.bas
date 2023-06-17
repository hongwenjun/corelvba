VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Replace_UI 
   Caption         =   "使剪贴板上的物件替换选择的目标物件"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   OleObjectBlob   =   "Replace_UI.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Replace_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#If VBA7 Then
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    
#Else
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&


Private Sub Close_Icon_Click()
  Unload Me    '// 关闭
End Sub

Private Sub UserForm_Initialize()
  Dim IStyle As Long
  Dim hWnd As Long
  
  hWnd = FindWindow("ThunderDFrame", Me.Caption)

  IStyle = GetWindowLong(hWnd, GWL_STYLE)
  IStyle = IStyle And Not WS_CAPTION
  SetWindowLong hWnd, GWL_STYLE, IStyle
  DrawMenuBar hWnd
  IStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
  SetWindowLong hWnd, GWL_EXSTYLE, IStyle
  
  With Me
  '  .StartUpPosition = 0
  '  .Left = 500
  '  .Top = 200
    .Width = 378
    .Height = 228
  End With
  
End Sub

Private Sub LOGO_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button Then
    mx = X
    my = Y
  End If
End Sub

Private Sub LOGO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button Then
    Me.Left = Me.Left - mx + X
    Me.Top = Me.Top - my + Y
  End If
End Sub


Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Dim pos_x As Variant
  Dim pos_y As Variant
  pos_x = Array(307, 27)
  pos_y = Array(64, 126, 188, 200)

  If Abs(X - pos_x(0)) < 30 And Abs(Y - pos_y(0)) < 30 Then
    Call copy_shape_replace
  ElseIf Abs(X - pos_x(0)) < 30 And Abs(Y - pos_y(1)) < 30 Then
    Call copy_shape_replace_resize
  ElseIf Abs(X - pos_x(0)) < 30 And Abs(Y - pos_y(2)) < 30 Then
    Call image_replace
  ElseIf Abs(X - pos_x(1)) < 30 And Abs(Y - pos_y(3)) < 30 Then
    API.WebHelp "https://262235.xyz/index.php/tag/vba/"
  End If
  
  Replace_UI.Hide
End Sub


Private Sub image_replace()
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim image_path As String
  image_path = API.GetClipBoardString
  ActiveDocument.ReferencePoint = cdrCenter
  
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim X As Double, Y As Double
  Set shs = ActiveSelection.Shapes
  cnt = 0
  For Each sh In shs
    If cnt = 0 Then
      ActiveDocument.ClearSelection
      ActiveLayer.Import image_path
      Set sc = ActiveSelection
      cnt = 1
    Else
      sc.Duplicate 0, 0
    End If
    sh.GetPosition X, Y
    sc.SetPosition X, Y
    
    sh.GetSize X, Y
    sc.SetSize X, Y
    sh.Delete
    
  Next sh

ErrorHandler:
'//    MsgBox "请先复制图片的完整路径，本工具能自动替换图片!"
  API.EndOpt
End Sub

Private Sub copy_shape_replace_resize()
  On Error GoTo ErrorHandler
  API.BeginOpt

  ActiveDocument.ReferencePoint = cdrCenter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim X As Double, Y As Double
  Set shs = ActiveSelection.Shapes
  cnt = 0
  For Each sh In shs
    If cnt = 0 Then
      Set sc = ActiveDocument.ActiveLayer.Paste
      cnt = 1
    Else
      sc.Duplicate 0, 0
    End If
    sh.GetPosition X, Y
    sc.SetPosition X, Y
    
    sh.GetSize X, Y
    sc.SetSize X, Y
    sh.Delete
    
  Next sh

ErrorHandler:
'// MsgBox "请先复制Ctrl+C，然后选择要替换的物件运行本工具!"
  API.EndOpt
End Sub


Private Sub copy_shape_replace()
  On Error GoTo ErrorHandler
  API.BeginOpt

  ActiveDocument.ReferencePoint = cdrCenter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim X As Double, Y As Double
  Set shs = ActiveSelection.Shapes
  cnt = 0
  For Each sh In shs
    If cnt = 0 Then
      Set sc = ActiveDocument.ActiveLayer.Paste
      cnt = 1
    Else
      sc.Duplicate 0, 0
    End If
    sh.GetPosition X, Y
    sc.SetPosition X, Y
    sh.Delete
  Next sh

ErrorHandler:
'// MsgBox "请先复制Ctrl+C，然后选择要替换的物件运行本工具!"
  API.EndOpt
End Sub

