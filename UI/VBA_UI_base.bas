#If VBA7 Then
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
 
#Else
    Private Declare Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Option Explicit
Dim mX As Long, mY As Long

Private Sub UserForm_Initialize()
  Dim IStyle As Long
  Dim Hwnd As Long
  
  Hwnd = FindWindow("ThunderDFrame", Me.Caption)

  IStyle = GetWindowLong(Hwnd, GWL_STYLE)
  IStyle = IStyle And Not WS_CAPTION
  SetWindowLong Hwnd, GWL_STYLE, IStyle
  DrawMenuBar Hwnd
  IStyle = GetWindowLong(Hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
  SetWindowLong Hwnd, GWL_EXSTYLE, IStyle
  
With Me
  .StartUpPosition = 0
  .Left = 500
  .Top = 200
  .Height = 312
  .Width = 36
End With

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button Then
        mX = X
        mY = Y
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button Then
    Me.Left = Me.Left - mx + x
    Me.Top = Me.Top - my + y
    End If
End Sub

Private Sub UserForm_Click()
     '// 屏幕分辨率
    Dim X As Long, Y As Long
    X = GetSystemMetrics(SM_CXSCREEN)
    Y = GetSystemMetrics(SM_CYSCREEN)
  '  MsgBox "您的屏幕分辨率为：" & x & "*" & y
      With Me
        .Height = 30
        .Top = .Top + 10
      End With
  ' MsgBox "窗口定位点: 左" & Me.Left & " 上 " & Me.Top & vbNewLine & "您的屏幕分辨率为：" & X & "*" & Y
 
End Sub
