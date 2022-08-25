VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CQL_FIND_UI 
   Caption         =   "使剪贴板上的物件替换选择的目标物件"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   OleObjectBlob   =   "CQL_FIND_UI.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "CQL_FIND_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Close_Icon_Click()
  Unload Me    ' 关闭
End Sub


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
  '  .StartUpPosition = 0
  '  .Left = 500
  '  .Top = 200
    .Width = 378
    .Height = 228
  End With
  
End Sub

Private Sub LOGO_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button Then
    mx = x
    my = y

  End If
End Sub

Private Sub LOGO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button Then
  Debug.Print x, y
    Me.Left = Me.Left - mx + x
    Me.Top = Me.Top - my + y
  End If
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim pos_x As Variant
  Dim pos_y As Variant
  pos_x = Array(307, 27)
  pos_y = Array(64, 126, 188, 200)

  If Abs(x - pos_x(0)) < 30 And Abs(y - pos_y(0)) < 30 Then
    Call CQLSameUniformColor
  ElseIf Abs(x - pos_x(0)) < 30 And Abs(y - pos_y(1)) < 30 Then
    Call CQLSameOutlineColor
  ElseIf Abs(x - pos_x(0)) < 30 And Abs(y - pos_y(2)) < 30 Then
    Call CQLSameSize
  ElseIf Abs(x - pos_x(1)) < 30 And Abs(y - pos_y(3)) < 30 Then
    CorelVBA.WebHelp "https://262235.xyz/index.php/tag/vba/"
  End If
  
  '// 预置颜色轮廓选择
  If Abs(x - 178) < 30 And Abs(y - 118) < 30 Then
    Debug.Print "选择图标: " & x & "  , " & y
    CQL查找相同.CQLline_CM100
  End If
  
  CQL_FIND_UI.Hide
End Sub

Private Sub CQLSameSize()
  ActiveDocument.Unit = cdrMillimeter
  Dim s As Shape
  Set s = ActiveShape
  If s Is Nothing Then Exit Sub
    
  If OptBt.value = True Then
    ActiveDocument.ClearSelection
    OptBt.value = 0
    CQL_FIND_UI.Hide
    
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim Shift As Long
    Dim box As Boolean
    box = ActiveDocument.GetUserArea(x1, y1, x2, y2, Shift, 10, False, cdrCursorWeldSingle)
    If Not b Then
      ' MsgBox "选区范围: " & x1 & y1 & x2 & y2
      Set sh = ActivePage.SelectShapesFromRectangle(x1, y1, x2, y2, False)
      sh.Shapes.FindShapes(Query:="@width = {" & s.SizeWidth & " mm} and @height ={" & s.SizeHeight & "mm}").CreateSelection
    End If
  Else
    ActivePage.Shapes.FindShapes(Query:="@width = {" & s.SizeWidth & " mm} and @height ={" & s.SizeHeight & "mm}").CreateSelection
  End If
End Sub

Private Sub CQLSameOutlineColor()
  On Error GoTo err
  Dim colr As New Color, s As Shape
  Set s = ActiveShape
  If s Is Nothing Then Exit Sub
  colr.CopyAssign s.Outline.Color
  colr.ConvertToRGB
  ' 查找对象
  r = colr.RGBRed
  G = colr.RGBGreen
  b = colr.RGBBlue
  
  If OptBt.value = True Then
    ActiveDocument.ClearSelection
    OptBt.value = 0
    CQL_FIND_UI.Hide
    
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim Shift As Long
    Dim box As Boolean
    box = ActiveDocument.GetUserArea(x1, y1, x2, y2, Shift, 10, False, cdrCursorWeldSingle)
    If Not b Then
      ' MsgBox "选区范围: " & x1 & y1 & x2 & y2
      Set sh = ActivePage.SelectShapesFromRectangle(x1, y1, x2, y2, False)
      sh.Shapes.FindShapes(Query:="@Outline.Color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").CreateSelection
    End If
  Else
    ActivePage.Shapes.FindShapes(Query:="@Outline.Color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").CreateSelection
  End If
  
  Exit Sub
err:
    MsgBox "对象轮廓为空。"
End Sub

Private Sub CQLSameUniformColor()
  On Error GoTo err
  Dim colr As New Color, s As Shape
  Set s = ActiveShape
  If s Is Nothing Then Exit Sub
  If s.Fill.Type = cdrFountainFill Then MsgBox "不支持渐变色。": Exit Sub
  colr.CopyAssign s.Fill.UniformColor
  colr.ConvertToRGB
  ' 查找对象
  r = colr.RGBRed
  G = colr.RGBGreen
  b = colr.RGBBlue
  
  If OptBt.value = True Then
    ActiveDocument.ClearSelection
    OptBt.value = 0
    CQL_FIND_UI.Hide
    
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim Shift As Long
    Dim box As Boolean
    box = ActiveDocument.GetUserArea(x1, y1, x2, y2, Shift, 10, False, cdrCursorWeldSingle)
    If Not b Then
      ' MsgBox "选区范围: " & x1 & y1 & x2 & y2
      Set sh = ActivePage.SelectShapesFromRectangle(x1, y1, x2, y2, False)
      sh.Shapes.FindShapes(Query:="@fill.color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").CreateSelection
    End If
  Else
    ActivePage.Shapes.FindShapes(Query:="@fill.color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").CreateSelection
  End If
  Exit Sub
err:
  MsgBox "对象填充为空。"
End Sub
