VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MakeSizePlus 
   Caption         =   "Batch Dimensions Plus"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   OleObjectBlob   =   "MakeSizePlus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MakeSizePlus"
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

'// 插件名称 VBA_UserForm
Private Const TOOLNAME As String = "LYVBA"
Private Const SECTION As String = "MakeSizePlus"
Private sreg As New ShapeRange

Private Sub UserForm_Initialize()
  With Me
    .StartUpPosition = 0
    .Left = Val(GetSetting(TOOLNAME, SECTION, "form_left", 900))
    .Top = Val(GetSetting(TOOLNAME, SECTION, "form_top", 200))
    .width = Val(GetSetting(TOOLNAME, SECTION, "form_width", 200))
    .Height = Val(GetSetting(TOOLNAME, SECTION, "form_Height", 105))
  End With

  LNG_CODE = API.GetLngCode
  Init_Translations Me, LNG_CODE
  Me.Caption = i18n("Batch Dimensions Plus", LNG_CODE)
  
   ' 读取线设置
  Bleed.text = API.GetSet("Bleed")
  Line_len.text = API.GetSet("Line_len")
  Outline_Width.text = GetSetting("LYVBA", "Settings", "Outline_Width", "0.2")
  Font_Size.text = GetSetting("LYVBA", "Settings", "Font_Size", "18")
End Sub

'// 关闭窗口时保存窗口位置
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    saveFormPos True
End Sub

'// 保存窗口位置和加载窗口位置
Sub saveFormPos(bDoSave As Boolean)
  If bDoSave Then 'save position
    SaveSetting TOOLNAME, SECTION, "form_left", Me.Left
    SaveSetting TOOLNAME, SECTION, "form_top", Me.Top
    SaveSetting TOOLNAME, SECTION, "form_width", Me.width
    SaveSetting TOOLNAME, SECTION, "form_Height", Me.Height
  End If
End Sub

Private Sub btn_ExpandForm_Click()
  With Me
    If .width = 200 Then
      .width = 260: .Height = 132
    ElseIf .Height = 132 Then
      .Height = 206
    Else
      .width = 200: .Height = 105
    End If
  End With
End Sub


'// Minimizes the window and retains dimensioning functionality   '// 最小化窗口并保留标注尺寸功能
Private Function MiniForm()

  Dim IStyle As Long
  Dim hwnd As Long
  
  hwnd = FindWindow("ThunderDFrame", MakeSizePlus.Caption)

  IStyle = GetWindowLong(hwnd, GWL_STYLE)
  IStyle = IStyle And Not WS_CAPTION
  SetWindowLong hwnd, GWL_STYLE, IStyle
  DrawMenuBar hwnd
  IStyle = GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
  SetWindowLong hwnd, GWL_EXSTYLE, IStyle

  Dim ctl As Variant  '// CorelDRAW 2020 定义成 Variant 才不会错误
  For Each ctl In MakeSizePlus.Controls
      ctl.Visible = False
      ctl.Top = 2
  Next ctl
  
  With Me
    .StartUpPosition = 0
    .BackColor = &H80000012
    .Left = Val(GetSetting("LYVBA", "Settings", "Left", "400")) + 318
    .Top = Val(GetSetting("LYVBA", "Settings", "Top", "55")) - 2
    .Height = 28
    .width = 98
    
    .MarkLines_Makesize.Visible = True
    .btn_Makesizes.Visible = True
    .Manual_Makesize.Visible = True
    .chkOpposite.Visible = True
    .X_EXIT.Visible = True
    
    .MarkLines_Makesize.Left = 1
    .btn_Makesizes.Left = 26
    .Manual_Makesize.Left = 50
    .chkOpposite.Left = 75: .chkOpposite.Top = 14
    .X_EXIT.Left = 85: .X_EXIT.Top = 0
  End With
End Function

Private Sub btn_MiniForm_Click()
  MiniForm
End Sub

Private Sub Settings_Click()
  If 0 < Val(Bleed.text) * Val(Line_len.text) < 100 Then
   SaveSetting "LYVBA", "Settings", "Bleed", Bleed.text
   SaveSetting "LYVBA", "Settings", "Line_len", Line_len.text
   SaveSetting "LYVBA", "Settings", "Outline_Width", Outline_Width.text
   SaveSetting "LYVBA", "Settings", "Font_Size", Font_Size.text
   Call API.Set_Space_Width  '// 设置空间间隙
  End If
End Sub


Private Sub btn_Makesizes_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim os As ShapeRange
  Dim s As Shape
  Dim sr As ShapeRange
  Set doc = ActiveDocument
  Set sr = ActiveSelectionRange
  sr.RemoveAll
    
  If Shift = 4 Then
    Set os = ActiveSelectionRange
    For Each s In os.Shapes
      If s.Type = cdrTextShape Then sr.Add s
    Next s
    
  sr.CreateSelection
  ElseIf Shift = 1 Then
    Set os = ActiveSelectionRange
    For Each s In os.Shapes
      If s.Type = cdrLinearDimensionShape Then sr.Add s
    Next s
    sr.CreateSelection
    
  ElseIf Shift = 2 Then
    Set os = ActiveSelectionRange
    For Each s In os.Shapes
      If s.Type = cdrLinearDimensionShape Then sr.Add s
    Next s
    sr.Delete
    If os.count > 0 Then
      os.Shapes.FindShapes(Query:="@name ='DMKLine'").CreateSelection
      ActiveSelectionRange.Delete
    End If
  Else
    make_sizes Shift
  End If
  
ErrorHandler:
  API.EndOpt
End Sub

Sub make_sizes_sep(dr, Optional shft = 0, Optional ByVal mirror As Boolean = False)
  On Error GoTo ErrorHandler
  API.BeginOpt "Make Size"
  Set doc = ActiveDocument
  Dim s As Shape, sh As Shape
  Dim pts As New SnapPoint, pte As New SnapPoint
  Dim os As ShapeRange
  
  Set os = ActiveSelectionRange
  
  Dim border As Variant
  Dim Line_len As Double
  Line_len = API.Set_Space_Width(True)  '// 读取间隔

  border = Array(cdrBottomRight, cdrBottomLeft, os.topY + Line_len, os.topY + 2 * Line_len, _
  cdrBottomRight, cdrTopRight, os.LeftX - Line_len, os.LeftX - 2 * Line_len)
  
  If mirror = True Then border = Array(cdrTopRight, cdrTopLeft, os.BottomY - Line_len, os.BottomY - 2 * Line_len, _
  cdrBottomLeft, cdrTopLeft, os.RightX + Line_len, os.RightX + 2 * Line_len)
  
  If dr = "upbx" Or dr = "upb" Or dr = "dnb" Or dr = "up" Or dr = "dn" Then Set os = X4_Sort_ShapeRange(os, stlx)
  If dr = "lfbx" Or dr = "lfb" Or dr = "rib" Or dr = "lf" Or dr = "ri" Then Set os = X4_Sort_ShapeRange(os, stty).ReverseRange

  If os.count > 0 Then
    If os.count > 1 And Len(dr) > 2 And os.Shapes.count > 1 Then
      For i = 1 To os.Shapes.count - 1
        Select Case dr
          Case "upbx"
#If VBA7 Then
            Set pts = os.Shapes(i).SnapPoints.BBox(border(0))
            Set pte = os.Shapes(i + 1).SnapPoints.BBox(border(1))
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, 0, border(2), cdrDimensionStyleEngineering)
            
            If shft > 0 And i = 1 Then
              Dimension_SetProperty sh, PresetProperty.value, mirror
              Set pts = os.FirstShape.SnapPoints.BBox(border(0))
              Set pte = os.LastShape.SnapPoints.BBox(border(1))
              Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, 0, border(3), cdrDimensionStyleEngineering)
            End If
          
          
          Case "lfbx"
            Set pts = os.Shapes(i).SnapPoints.BBox(border(4))
            Set pte = os.Shapes(i + 1).SnapPoints.BBox(border(5))
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, border(6), 0, cdrDimensionStyleEngineering)
            
            If shft > 0 And i = 1 Then
              Dimension_SetProperty sh, PresetProperty.value, mirror
              Set pts = os.FirstShape.SnapPoints.BBox(border(4))
              Set pte = os.LastShape.SnapPoints.BBox(border(5))
              Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, border(7), 0, cdrDimensionStyleEngineering)
            End If
#Else
' X4  There is a difference
            Set pts = CreateSnapPoint(os.Shapes(i).CenterX, os.Shapes(i).CenterY)
            Set pte = CreateSnapPoint(os.Shapes(i + 1).CenterX, os.Shapes(i + 1).CenterY)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, 0, border(2), Textsize:=18)
            
          Case "lfbx"
            Set pts = CreateSnapPoint(os.Shapes(i).CenterX, os.Shapes(i).CenterY)
            Set pte = CreateSnapPoint(os.Shapes(i + 1).CenterX, os.Shapes(i + 1).CenterY)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, border(6), 0, Textsize:=18)
#End If
          
          Case "upb"
            Set pts = os.Shapes(i).SnapPoints.BBox(cdrTopRight)
            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopLeft)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.topY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
            
          Case "dnb"
            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomRight)
            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrBottomLeft)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.BottomY - os.SizeHeight / 10, cdrDimensionStyleEngineering)
            
          Case "lfb"
            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomLeft)
            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopLeft)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, os.LeftX - os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
            
          Case "rib"
            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomRight)
            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopRight)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, os.RightX + os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
        End Select
        '// 尺寸标注设置属性
        Dimension_SetProperty sh, PresetProperty.value, mirror
        'ActiveDocument.ClearSelection
      Next i
    Else
      If shft > 0 Then
        Select Case dr
          Case "up"
            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopLeft)
            Set pte = os.LastShape.SnapPoints.BBox(cdrTopRight)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.topY + os.SizeHeight / 10, cdrDimensionStyleEngineering)

          Case "dn"
            Set pts = os.FirstShape.SnapPoints.BBox(cdrBottomLeft)
            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomRight)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.BottomY - os.SizeHeight / 10, cdrDimensionStyleEngineering)

          Case "lf"
            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopLeft)
            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomLeft)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, os.LeftX - os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
          
          Case "ri"
            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopRight)
            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomRight)
            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, os.RightX + os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
        End Select
        Dimension_SetProperty sh, PresetProperty.value, mirror
      Else
        For Each s In os.Shapes
          Select Case dr
            Case "up"
              Set pts = s.SnapPoints.BBox(cdrTopLeft)
              Set pte = s.SnapPoints.BBox(cdrTopRight)
              Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, s.LeftX + s.SizeWidth / 10, s.topY + s.SizeHeight / 10, cdrDimensionStyleEngineering)
            
            Case "dn"
              Set pts = s.SnapPoints.BBox(cdrBottomLeft)
              Set pte = s.SnapPoints.BBox(cdrBottomRight)
              Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, s.LeftX + s.SizeWidth / 10, s.BottomY - s.SizeHeight / 10, cdrDimensionStyleEngineering)
            
            Case "lf"
              Set pts = s.SnapPoints.BBox(cdrTopLeft)
              Set pte = s.SnapPoints.BBox(cdrBottomLeft)
              Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, s.LeftX - s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering)
            
            Case "ri"
              Set pts = s.SnapPoints.BBox(cdrTopRight)
              Set pte = s.SnapPoints.BBox(cdrBottomRight)
              Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, s.RightX + s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering)
          End Select
          Dimension_SetProperty sh, PresetProperty.value, mirror
        Next s
      End If
    End If
  End If
  os.CreateSelection
  
ErrorHandler:
  
  API.EndOpt
End Sub

Sub make_sizes(Optional shft = 0)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim s As Shape
  Dim pts As SnapPoint, pte As SnapPoint
  Dim os As ShapeRange
  Set os = ActiveSelectionRange
  If os.count > 0 Then
  For Each s In os.Shapes
#If VBA7 Then
      Set pts = s.SnapPoints.BBox(cdrTopLeft)
      Set pte = s.SnapPoints.BBox(cdrTopRight)
      Set ptle = s.SnapPoints.BBox(cdrBottomLeft)
      If shft <> 6 Then Dimension_SetProperty ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, ptle, True, _
                                              s.LeftX - s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering), PresetProperty.value
      If shft <> 3 Then Dimension_SetProperty ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, _
                                          s.LeftX + s.SizeWidth / 10, s.topY + s.SizeHeight / 10, cdrDimensionStyleEngineering), PresetProperty.value
#Else
' X4  There is a difference
      Set pts = s.SnapPoints(cdrTopLeft)
      Set pte = s.SnapPoints(cdrTopRight)
      Set ptle = s.SnapPoints(cdrBottomLeft)
      If shft <> 6 Then ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, ptle, True, _
                      s.LeftX - s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering, Textsize:=18
      If shft <> 3 Then ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, _
                      s.LeftX + s.SizeWidth / 10, s.topY + s.SizeHeight / 10, cdrDimensionStyleEngineering, Textsize:=18
#End If
  Next s
  End If

ErrorHandler:
  API.EndOpt
End Sub

'// 使用标记线批量建立尺寸标注:   左键上标注，右键右标注
Private Sub MarkLines_Makesize_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  SRMInst 3, "sw"
  '// 右键
  If Button = 2 Then
    If chkOpposite.value = True Then
        CutLines.Dimension_MarkLines cdrAlignTop, True
        make_sizes_sep "upbx", Shift, True
    Else
      CutLines.Dimension_MarkLines cdrAlignLeft, True
      make_sizes_sep "lfbx", Shift, True
    End If
  
  '// 左键
  ElseIf Button = 1 Then
    If chkOpposite.value = True Then
      CutLines.Dimension_MarkLines cdrAlignLeft, False
      make_sizes_sep "lfbx", Shift, False
    Else
        CutLines.Dimension_MarkLines cdrAlignTop, False
        make_sizes_sep "upbx", Shift, False
    End If
  End If
  SRMInst 3, "lw"
End Sub

'// 自动酷炫风格标注
Private Sub CoolStyle_Click()
  SRMInst 3, "sw"

  CutLines.Dimension_MarkLines cdrAlignTop, False
  make_sizes_sep "upbx", Shift, False

  SRMInst 3, "lw"
  CutLines.Dimension_MarkLines cdrAlignLeft, False
  make_sizes_sep "lfbx", Shift, False

  SRMInst 3, "lw"
  CutLines.Dimension_MarkLines cdrAlignTop, True
  make_sizes_sep "upbx", Shift, True

  SRMInst 3, "lw"
  CutLines.Dimension_MarkLines cdrAlignLeft, True
  make_sizes_sep "lfbx", Shift, True
  SRMInst 3, "lw"
  
End Sub

'// 快速标注尺寸样式
Private Sub QuickStyle_Click()
  Dim os As ShapeRange
  Set os = ActiveSelectionRange
  SRMInst 3, "sw"
  
  CutLines.Dimension_MarkLines cdrAlignTop, True
  make_sizes_sep "upbx", 2, True
  SRMInst 4, "sw"
  
  SRMInst 3, "lw"
  CutLines.Dimension_MarkLines cdrAlignLeft, False
  make_sizes_sep "lfbx", Shift, False

  SRMInst 4, "lw"
  
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  sr.Sort "@shape1.left<@shape2.left"
  If sr.count > 5 And IsAllSameSize(os) Then
    n = sr.count
    sr.Remove n: sr.Remove (n - 1)
    sr.Remove 3: sr.Remove 2: sr.Remove 1
    sr.Delete
  End If
  
  SRMInst 3, "lw"
End Sub

'// 标注文字红色，分离标注
Private Sub QuickRedText_Click()
  SRMInst 3, "sw"
  
  '// 选择文本，改成红色
  ModulePlus.Dimension_Select_or_Delete 4
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  sr.ApplyUniformFill CreateCMYKColor(0, 100, 100, 0)
  
  '// 解绑标注线
  SRMInst 3, "lw"
  ModulePlus.Untie_MarkLines
  
  SRMInst 3, "lw"
  
End Sub


'// 使用手工选节点建立尺寸标注，使用Ctrl分离尺寸标注
Private Sub Manual_Makesize_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
      '// 右键
  ElseIf Shift = fmCtrlMask Then
      Slanted_Makesize  '// 手动标注倾斜尺寸
  Else
      ModulePlus.Untie_MarkLines   '// 解绑尺寸，分离尺寸
  End If
End Sub

'// 手动标注倾斜尺寸
Private Function Slanted_Makesize()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim nr As NodeRange, cnt As Integer
  Dim sr As ShapeRange, sh As Shape
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  
  Set sr = ActiveSelectionRange
  Set nr = ActiveShape.Curve.Selection
  
  If chkOpposite.value = False Then
    Slanted_Sort_Make sr  '// 排序标注倾斜尺寸
    Exit Function
  End If
  If nr.count < 2 Then Exit Function

  cnt = nr.count
  While cnt > 1
    x1 = nr(cnt).PositionX
    y1 = nr(cnt).PositionY
    x2 = nr(cnt - 1).PositionX
    y2 = nr(cnt - 1).PositionY
    
    Set pts = CreateSnapPoint(x1, y1)
    Set pte = CreateSnapPoint(x2, y2)
    Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionSlanted, pts, pte, True, x1 - 20, y1 + 20, cdrDimensionStyleEngineering)
    
    Dimension_SetProperty sh, PresetProperty.value, mirror
    cnt = cnt - 1
  Wend

ErrorHandler:
  API.EndOpt
End Function

'// 排序标注倾斜尺寸
Private Function Slanted_Sort_Make(shs As ShapeRange)
  On Error GoTo ErrorHandler
  Dim sr As New ShapeRange
  Dim s As Shape, sh As Shape
  Dim nr As NodeRange
  For Each sh In shs
    Set nr = sh.Curve.Selection
    For Each n In nr
      Set s = ActiveLayer.CreateEllipse2(n.PositionX, n.PositionY, 0.5, 0.5)
      sr.Add s
    Next n
  Next sh
  
  CutLines.RemoveDuplicates sr  '// 简单删除重复算法
  Set sr = X4_Sort_ShapeRange(sr, stlx)

  For i = 1 To sr.count - 1
    x1 = sr(i + 1).CenterX
    y1 = sr(i + 1).CenterY
    x2 = sr(i).CenterX
    y2 = sr(i).CenterY
    
    Set pts = CreateSnapPoint(x1, y1)
    Set pte = CreateSnapPoint(x2, y2)
#If VBA7 Then
    Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionSlanted, pts, pte, True, x1 - 20, y1 + 20, cdrDimensionStyleEngineering)
#Else
' X4  There is a difference
    Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionSlanted, pts, pte, True, (x1 + x2) / 2, (y1 + y2) / 2, cdrDimensionStyleEngineering, Textsize:=18)
#End If
    Dimension_SetProperty sh, PresetProperty.value
  Next i
  sr.Delete

ErrorHandler:
  API.EndOpt
End Function

'// 尺寸标注设置属性
Private Function Dimension_SetProperty(sh_dim As Shape, Optional ByVal Preset As Boolean = False, Optional ByVal mirror As Boolean = False)
#If VBA7 Then
  plt = 0: If periphery.value And mirror Then plt = 1
  
  If Preset And sh_dim.Type = cdrLinearDimensionShape Then
    With sh_dim.Style.GetProperty("dimension")
      .SetProperty "precision", 0 '       小数位数
      .SetProperty "showUnits", 0 '       是否显示单位 0/1
      
      .SetProperty "textPlacement", plt  '   0、上方，1、下方，2、中间

    '  .SetProperty "dynamicText", 0 '    是否可以编辑尺寸0/1
    '  .SetProperty "overhang", 500000 '
    End With
  End If
  
  sh_dim.Outline.width = API.GetSet("Outline_Width")
  sh_dim.Dimension.TextShape.text.Story.size = Font_Size.value
  
#Else
' X4  There is a difference
#End If
End Function

'// 尺寸标注左边
Private Sub Makesize_Left_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  SRMInst 3, "sw"
  If Button = 2 Then
    CutLines.Dimension_MarkLines cdrAlignLeft, False
    make_sizes_sep "lfbx", Button, False
      
  ElseIf Shift = fmCtrlMask Then
    CutLines.Dimension_MarkLines cdrAlignLeft, False
    make_sizes_sep "lfbx", Shift, False
  Else
    '// Ctrl Key
    make_sizes_sep "lfb"
  End If
  SRMInst 3, "lw"
End Sub

'// 尺寸标注右边
Private Sub Makesize_Right_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  SRMInst 3, "sw"
  If Button = 2 Then
    CutLines.Dimension_MarkLines cdrAlignLeft, True
    make_sizes_sep "lfbx", Button, True
    
  ElseIf Shift = fmCtrlMask Then
    CutLines.Dimension_MarkLines cdrAlignLeft, True
    make_sizes_sep "lfbx", Shift, True
  Else
    '// Ctrl Key
    make_sizes_sep "rib"
  End If
  SRMInst 3, "lw"
End Sub

'// 尺寸标注向上
Private Sub Makesize_Up_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  SRMInst 3, "sw"
  If Button = 2 Then
    CutLines.Dimension_MarkLines cdrAlignTop, False
    make_sizes_sep "upbx", Button, False
      
  ElseIf Shift = fmCtrlMask Then
    CutLines.Dimension_MarkLines cdrAlignTop, False
    make_sizes_sep "upbx", Shift, False
  Else
   '// Ctrl Key
    make_sizes_sep "upb"
  End If
  SRMInst 3, "lw"
End Sub

'// 尺寸标注向下
Private Sub Makesize_Down_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  SRMInst 3, "sw"
  If Button = 2 Then
    CutLines.Dimension_MarkLines cdrAlignTop, True
    make_sizes_sep "upbx", Button, True
      
  ElseIf Shift = fmCtrlMask Then
    CutLines.Dimension_MarkLines cdrAlignTop, True
    make_sizes_sep "upbx", Shift, True
  Else
   '// Ctrl Key
    make_sizes_sep "dnb"
  End If
  SRMInst 3, "lw"
End Sub

Private Sub MakeRuler_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  On Error GoTo ErrorHandler
  API.BeginOpt
  Set sreg = Nothing
  
  If Button = 2 And Shift = 0 Then       '// 鼠标右键 标注右边
    Ruler_Align cdrAlignRight
    
  ElseIf Button = 2 And Shift = 2 Then  '// Ctrl+鼠标右键 标注左边
    Ruler_Align cdrAlignLeft
 
  ElseIf Shift = 0 Then    '// 鼠标左键，标注在上边
    Ruler_Align cdrAlignTop
    
  ElseIf Shift = 2 Then  '// Ctrl+鼠标左键，标注下边
    Ruler_Align cdrAlignBottom
  End If
  
  sreg.CreateSelection
ErrorHandler:
  API.EndOpt
End Sub

Private Sub MakeRuler_Align_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  On Error GoTo ErrorHandler
  API.BeginOpt
  Set sreg = Nothing
   
  Dim ra As cdrAlignType
  ra = cdrAlignTop
  ' 定义方向上下左右
  Dim pos_x As Variant, pos_y As Variant
  pos_x = Array(27, 27, 12, 44)
  pos_y = Array(12, 44, 27, 27)
  If Abs(X - pos_x(0)) < 14 And Abs(Y - pos_y(0)) < 14 Then
    ra = cdrAlignTop
  ElseIf Abs(X - pos_x(1)) < 14 And Abs(Y - pos_y(1)) < 14 Then
    ra = cdrAlignBottom
  ElseIf Abs(X - pos_x(2)) < 14 And Abs(Y - pos_y(2)) < 14 Then
    ra = cdrAlignLeft
  ElseIf Abs(X - pos_x(3)) < 14 And Abs(Y - pos_y(3)) < 14 Then
    ra = cdrAlignRight
  End If
  
  Ruler_Align ra
  sreg.CreateSelection
ErrorHandler:
  API.EndOpt
End Sub

Private Function Ruler_Align(ra As cdrAlignType)
  If ra = cdrAlignRight Then       '// 标注右边
    CutLines.Dimension_MarkLines cdrAlignLeft, True
    Add_Ruler_Text_Y True
  ElseIf ra = cdrAlignLeft Then  '// 标注左边
    CutLines.Dimension_MarkLines cdrAlignLeft, False
    Add_Ruler_Text_Y True
  ElseIf ra = cdrAlignTop Then    '// 标注上边
    CutLines.Dimension_MarkLines cdrAlignTop, False
    Add_Ruler_Text True
  ElseIf ra = cdrAlignBottom Then  '// 标注下边
    CutLines.Dimension_MarkLines cdrAlignTop, True
    Add_Ruler_Text True
  End If
End Function

  '// 标尺线转换成距离数字
Private Function Add_Ruler_Text(rm_lines As Boolean)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim s As Shape, t As Shape, sr As ShapeRange
  Dim text As String
  Set sr = ActiveSelectionRange
  Set sr = X4_Sort_ShapeRange(sr, stlx)
  For Each s In sr
    X = s.CenterX: Y = s.CenterY
    text = str(Int(X - sr.FirstShape.CenterX + 0.5))
    Set t = ActiveLayer.CreateArtisticText(X, Y, text, size:=Font_Size.value)
    t.CenterX = X: t.CenterY = Y
    sreg.Add t
  Next
  
  If rm_lines Then sr.Delete
ErrorHandler:
  API.EndOpt
End Function

  '// 标尺线转换成距离数字
Private Function Add_Ruler_Text_Y(rm_lines As Boolean)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim s As Shape, t As Shape, sr As ShapeRange
  Dim text As String
  Set sr = ActiveSelectionRange
  Set sr = X4_Sort_ShapeRange(sr, stty)
  For Each s In sr
    X = s.CenterX: Y = s.CenterY
    text = str(Int(Y - sr.FirstShape.CenterY + 0.5))
    Set t = ActiveLayer.CreateArtisticText(X, Y, text, size:=Font_Size.value)
    t.Rotate 90
    t.CenterX = X: t.CenterY = Y
    sreg.Add t
  Next
  
  If rm_lines Then sr.Delete
ErrorHandler:
  API.EndOpt
End Function

Private Sub X_EXIT_Click()
  Me.width = 200: Me.Height = 105
  Unload Me    '// EXIT
End Sub

Private Sub I18N_LNG_Click()
  LNG_CODE = API.GetLngCode
  If LNG_CODE = 1033 Then
    LNG_CODE = 2052
  Else
    LNG_CODE = 1033
  End If
  SaveSetting "LYVBA", "Settings", "I18N_LNG", LNG_CODE
  LNG_CODE = API.GetLngCode
  MsgBox i18n("Chinese And English Language Switching Is Completed, Please Restart The Plug-In.", LNG_CODE), vbOKOnly, i18n("Lanya Corelvba Plug-In", LNG_CODE)
End Sub


Private Sub Bt_SplitSegment_Click()
  ModulePlus.SplitSegment
End Sub

Private Sub btn_square_hi_Click()
  ModulePlus.square_hw "Height"
End Sub

Private Sub btn_square_wi_Click()
  ModulePlus.square_hw "Width"
End Sub

'// 节点连接合并
Private Sub btn_join_nodes_Click()
    ActiveSelection.CustomCommand "ConvertTo", "JoinCurves"
    Application.Refresh
End Sub

'// 节点优化减少
Private Sub btn_nodes_reduce_Click()
  ModulePlus.Nodes_Reduce
End Sub

'// 选择标注线 选择文字 删除或者解绑标准线
Private Sub SelectText_Click()
  ModulePlus.Dimension_Select_or_Delete 4
End Sub
Private Sub SelectLine_Click()
  ModulePlus.Dimension_Select_or_Delete 1
End Sub
Private Sub Delete_Dimension_Click()
  ModulePlus.Dimension_Select_or_Delete 2
End Sub
Private Sub bt_Untie_MarkLines_Click()
  ModulePlus.Untie_MarkLines
End Sub

'// Select_Range 工具组合按钮
Private Sub MADD_Click()
  SRMInst 1, "add"
End Sub
Private Sub MSUB_Click()
  SRMInst 1, "sub"
End Sub
Private Sub MRLW_Click()
  SRMInst 1, "lw"
End Sub
Private Sub MZERO_Click()
    SRMInst 1, "zero"
    MsgBox "Selection Range is Removed!"
End Sub

'''////  CorelDRAW 与 Adobe_Illustrator 剪贴板转换  ////'''
Private Sub Adobe_Illustrator_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Dim value As Integer
  If Button = 2 Then
    savePDFtoClip.AICopyToCdr
    Exit Sub
  End If
  
  If Button Then
    savePDFtoClip.CdrCopyToAI
    MsgBox "CorelDRAW 与 Adobe_Illustrator 剪贴板转换" & vbNewLine & "鼠标左键复制，鼠标右键粘贴"
  End If
End Sub

'// 修复圆角缺角到直角
Private Sub btn_corners_off_Click()
  Tools.corner_off
End Sub
