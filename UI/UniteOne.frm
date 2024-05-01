VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UniteOne 
   Caption         =   "Merge Multiple Pages Into One"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   OleObjectBlob   =   "UniteOne.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UniteOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Dim iHang, iLie, iPages As Integer     '// 定义行数(Y) 列数(X)
Dim iYouyi, iXiayi As Single           '// 右移(R) 下移(B)
                                       '// txtHang, txtLie, txtYouyi, txtXiayi ,txtInfo
Dim LogoFile As String                 '// Logo

Dim s(1 To 255) As Shape   '// 定义对象用于存放每页的群组
Dim P As Page              '// 定义多页
 
'// *********** 初始化程序 ***************
Private Sub UserForm_Initialize()
  Dim s As Shape
  ActiveDocument.Unit = cdrMillimeter '// 本文档单位为mm
  
  For Each P In ActiveDocument.Pages
    iPages = P.index
    If iPages = 1 Then
      P.Activate
      P.Shapes.all.CreateSelection
      
      Set s = ActiveDocument.Selection
      If s.Shapes.Count = 0 Then
       MsgBox i18n("The current document's first page is blank and has no objects.", LNG_CODE)
      Exit Sub
    End If
    
    End If
  Next P
 
 txtLie.text = 5
 txtHang.text = Int(iPages / CInt(txtLie.text) + 0.9)
 txtLie.text = Int(iPages / CInt(txtHang.text) + 0.9)
 
 iHang = CInt(txtHang.text)
 iLie = CInt(txtLie.text)
 
 
 iYouyi = Int(s.SizeWidth + 0.6)
 iXiayi = Int(s.SizeHeight + 0.6)
 
 txtYouyi.text = iYouyi
 txtXiayi.text = iXiayi
 
  LogoFile = path & "GMS\LYVBA\LOGO.jpg"
  If API.ExistsFile_UseFso(LogoFile) Then
    LogoPic.Picture = LoadPicture(LogoFile)   '// 换LOGO图
  End If

  LNG_CODE = API.GetLngCode
  Init_Translations Me, LNG_CODE
  Me.Caption = i18n("Merge Multiple Pages Into One", LNG_CODE)
  Me.Matrix.Caption = i18n("Matrix", LNG_CODE)
  Me.OffsetSelection.Caption = i18n("Offset Selection", LNG_CODE)
 
  txtInfo.text = i18n("Total Pages:", LNG_CODE) & iPages & "  " & i18n("Home Page Shape Size(mm):", LNG_CODE) & s.SizeWidth & "x" & s.SizeHeight
End Sub

'**** 主程序  执行
Private Sub cmdRun_Click()
  API.BeginOpt
  
  Dim x_M, y_M
  ActiveDocument.EditAcrossLayers = False    '// 跨图层编辑禁止
  
  For Each P In ActiveDocument.Pages
    P.Activate                              '// 激活每页
    P.Shapes.all.CreateSelection            '// 每页全选
    Set s(P.index) = ActiveSelection.Group  '// 存放每页的群组
  Next P
  
  ActiveDocument.EditAcrossLayers = True     '// 跨图层编辑开启
  
  x_M = y_M = 0
  
  For Each P In ActiveDocument.Pages
    P.Activate
      
    s(P.index).MoveToLayer ActivePage.DesktopLayer         '// 每页对象移动到桌面层
    s(P.index).Move (iYouyi * x_M), -(300 + iXiayi * y_M)  '// 排列对象  右偏移，下偏移
    
    y_M = y_M + 1
    
    If y_M = iLie Then
      x_M = x_M + 1
      y_M = 0
    End If
  
  Next P
  
  Unload Me
  API.EndOpt
End Sub

'**** 主程序 副本 横排序
Private Sub cmdRunX_Click()
  API.BeginOpt

  Dim x_M, y_M
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.EditAcrossLayers = False
  
  For Each P In ActiveDocument.Pages
    P.Activate
    P.Shapes.all.CreateSelection
    Set s(P.index) = ActiveSelection.Group
  Next P
  
  ActiveDocument.EditAcrossLayers = True
  
  x_M = y_M = 0
  
  For Each P In ActiveDocument.Pages
    P.Activate
    
    s(P.index).MoveToLayer ActivePage.DesktopLayer
    s(P.index).Move (iYouyi * y_M), -(600 + iXiayi * x_M)
    
    y_M = y_M + 1
    
    If y_M = iHang Then
      x_M = x_M + 1
      y_M = 0
    End If
  
  Next P
  
  Unload Me
  API.EndOpt
End Sub

'// 关闭
Private Sub cmdClose_Click()
  Unload Me
End Sub

'// VB限制文本框只能输入数字和小数点
Private Sub txtHang_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Numbers As String
Numbers = "1234567890"
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If
End Sub

Private Sub txtLie_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Numbers As String
Numbers = "1234567890"
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If
End Sub

Private Sub txtXiayi_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If
End Sub

Private Sub txtYouyi_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If
End Sub

Private Sub txtHang_Change()
    Dim n As Single
    n = Val(txtHang.text)
    If n > 0 And n < 1001 Then
        HangSpin.value = n
        iHang = n
    End If
 
 txtHang.text = iHang
 txtLie.text = Int(iPages / iHang + 0.9)
  
  iLie = CInt(txtLie.text)
End Sub

Private Sub HangSpin_Change()
    txtHang.text = CStr(HangSpin.value)
End Sub

Private Sub txtLie_Change()
    Dim n As Single
    n = Val(txtLie.text)
    If n > 0 And n < 1001 Then
        LieSpin.value = n
        iLie = n
    End If
    
    txtLie.text = iLie
    txtHang.text = Int(iPages / iLie + 0.9)
    
    iHang = CInt(txtHang.text)
End Sub

Private Sub LieSpin_Change()
    txtLie.text = CStr(LieSpin.value)
End Sub


Private Sub txtXiayi_Change()
    Dim n As Single
    n = Val(txtXiayi.text)
    If n > 0 And n < 1001 Then
        iXiayi = n
    End If
End Sub

Private Sub txtYouyi_Change()
    Dim n As Single
    n = Val(txtYouyi.text)
    If n > 0 And n < 1001 Then
        iYouyi = n
    End If
End Sub

