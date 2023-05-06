VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhotoForm 
   Caption         =   "对象批量转位图 by filon [玉环]"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "PhotoForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "PhotoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Private Const GWL_STYLE = (-16) '设置窗口样式
Private Const WS_MINIMIZEBOX As Long = &H20000 '最小化

Private Sub CovPhotos_Click()
    On Error Resume Next
    ActiveDocument.BeginCommandGroup
    Dim a, Color As String
    Dim b As Boolean
    Dim i, dpi As Integer
    
    If ABox1.Value = False Then
        a = False
    Else
        a = True
    End If

    b = True
    If BBox2.Value = False Then b = False
    
    dpi = CInt(ComboBox2.text)
    
    Select Case ComboBox1.text
      Case Is = "灰度"
      Color = cdrGrayscaleImage
      Case Is = "CMYK颜色"
      Color = cdrCMYKColorImage
      Case Is = "RGB颜色"
      Color = cdrRGBColorImage
      Case Is = "黑白"
      Color = cdrBlackAndWhiteImage
    End Select
    
    Dim s As Shapes
    Set s = ActiveSelection.Shapes
    If s Is Nothing Then
        MsgBox "请先选中一个形状!"
        Exit Sub
    Else
        For i = 1 To s.Count
        Set s(i) = ActiveShape.ConvertToBitmapEx(Color, False, a, dpi, cdrNormalAntiAliasing, True, False, 95)
        Next i
    End If
    ActiveDocument.EndCommandGroup
End Sub


Private Sub UserForm_Initialize()
Dim hWndForm As Long
Dim IStyle As Long
hWndForm = FindWindow("ThunderDFrame", Me.Caption)  '获取窗口句柄
IStyle = GetWindowLong(hWndForm, GWL_STYLE) '获取当前标题栏样式
IStyle = IStyle Or WS_MINIMIZEBOX '设置最小化按钮
SetWindowLong hWndForm, GWL_STYLE, IStyle  '显示最小化按钮
    On Error Resume Next
    ComboBox1.AddItem "灰度"
    ComboBox1.AddItem "CMYK颜色"
    ComboBox1.AddItem "RGB颜色"
    ComboBox1.AddItem "黑白"
    
    ComboBox2.AddItem "300", 0
    ComboBox2.AddItem "450", 1
    ComboBox2.AddItem "600", 2
End Sub

