VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CQL_FIND_UI 
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   OleObjectBlob   =   "CQL_FIND_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CQL_FIND_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

#If VBA7 Then
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    
#Else
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&


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
  '  .StartUpPosition = 0
  '  .Left = 500
  '  .Top = 200
    .width = 378
    .Height = 228
  End With
  
  LNG_CODE = API.GetLngCode
  Init_Translations Me, LNG_CODE
  
  If LNG_CODE = 1033 Then
    txtInfo.text = "Usage: A->Left B->Right C->Ctrl"
  Else
    txtInfo.text = "使用: A->左键 B->右键 C->Ctrl键"
  End If
End Sub

Private Sub LOGO_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button Then
    mx = X
    my = Y

  End If
End Sub

Private Sub LOGO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button Then
'//  Debug.Print X, Y
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
    Call CQLSameUniformColor
  ElseIf Abs(X - pos_x(0)) < 30 And Abs(Y - pos_y(1)) < 30 Then
    Call CQLSameOutlineColor
  ElseIf Abs(X - pos_x(0)) < 30 And Abs(Y - pos_y(2)) < 30 Then
    Call CQLSameSize
  ElseIf Abs(X - pos_x(1)) < 30 And Abs(Y - pos_y(3)) < 30 Then
'//   WebHelp "https://262235.xyz/index.php/tag/vba/"
  End If
  
    '// 预置颜色轮廓选择    和 '// 彩蛋功能
  If Abs(X - 178) < 30 And Abs(Y - 118) < 30 = True Then
    Image1.Visible = False
    Close_Icon.Visible = False
    X_EXIT.Visible = True
    
    With CQL_FIND_UI
      .StartUpPosition = 0
      .Left = Val(GetSetting("LYVBA", "Settings", "Left", "400")) + 318
      .Top = Val(GetSetting("LYVBA", "Settings", "Top", "55")) - 2
      .Height = 30
      .width = .width - 20
    End With
    
    If OptBt.value Then
      frmSelectSame.Show 0
    Else
      CQLFindSame.CQLline_CM100
    End If
    Exit Sub
  End If
  CQL_FIND_UI.Hide
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
      '// MsgBox "选区范围: " & x1 & y1 & x2 & y2
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

Private Sub X_EXIT_Click()
  Unload Me    '// 关闭
End Sub

Private Sub Close_Icon_Click()
  Unload Me    '// 关闭
End Sub
