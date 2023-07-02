VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Woodman 
   Caption         =   "Batch Dimension Nodes"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   OleObjectBlob   =   "Woodman.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Woodman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

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

'// Minimizes the window and retains dimensioning functionality   '// 最小化窗口并保留标注尺寸功能
Private Function MiniForm()

  Dim IStyle As Long
  Dim hWnd As Long
  
  hWnd = FindWindow("ThunderDFrame", Woodman.Caption)

  IStyle = GetWindowLong(hWnd, GWL_STYLE)
  IStyle = IStyle And Not WS_CAPTION
  SetWindowLong hWnd, GWL_STYLE, IStyle
  DrawMenuBar hWnd
  IStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
  SetWindowLong hWnd, GWL_EXSTYLE, IStyle

' Dim ctl As Control  '// CorelDRAW 2020 需要注释，才不会错误
  For Each ctl In Woodman.Controls
      ctl.Visible = False
      ctl.Top = 2
  Next ctl
  
  With Me
    .StartUpPosition = 0
    .BackColor = &H80000012
    .Left = Val(GetSetting("LYVBA", "Settings", "Left", "400")) + 318
    .Top = Val(GetSetting("LYVBA", "Settings", "Top", "55")) - 2
    .Height = 28
    .Width = 98
    
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

Private Sub btn_square_hi_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ActiveDocument.BeginCommandGroup:  Application.Optimization = True
    Set os = ActiveSelectionRange
    Set ss = os.Shapes
    uc = 0
    For Each s In ss
        s.SizeWidth = s.SizeHeight
        uc = uc + 1
    Next s
    Application.Optimization = False
    ActiveWindow.Refresh:    Application.Refresh
End Sub


Private Sub btn_square_wi_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ActiveDocument.BeginCommandGroup:  Application.Optimization = True
    Set os = ActiveSelectionRange
    Set ss = os.Shapes
    uc = 0
    For Each s In ss
        s.SizeHeight = s.SizeWidth
        uc = uc + 1
    Next s
    Application.Optimization = False
    ActiveWindow.Refresh:    Application.Refresh
End Sub

Private Sub btn_Makesizes_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim os As ShapeRange
    Dim s As Shape
    Dim sr As ShapeRange
    Set doc = ActiveDocument
    
'rasm.Dimension.TextShape.Text.Story.size = CLng(fnt)
'rasm.Style.GetProperty("dimension").SetProperty "precision", 0
'rasm.Style.GetProperty("dimension").SetProperty "units", 3
    
    doc.BeginCommandGroup "delete sizes"
        Set sr = ActiveSelectionRange
        sr.RemoveAll
    If Shift = 4 Then
        On Error Resume Next
        Set os = ActiveSelectionRange
        For Each s In os.Shapes
            If s.Type = cdrLinearDimensionShape Then s.Delete
        Next s
        On Error GoTo 0
    ElseIf Shift = 1 Then
        Set os = ActiveSelectionRange
        For Each s In os.Shapes
            If s.Type = cdrLinearDimensionShape Then sr.Add s
        Next s
        sr.CreateSelection
        On Error GoTo 0
    ElseIf Shift = 2 Then
        On Error Resume Next
        Set os = ActiveSelectionRange
        For Each s In os.Shapes
            If s.Type = cdrLinearDimensionShape Then s.Delete
        Next s
          If os.Count > 0 Then
            os.Shapes.FindShapes(Query:="@name ='DMKLine'").CreateSelection
            ActiveSelectionRange.Delete
          End If
        On Error GoTo 0
    Else
        make_sizes Shift
    End If
    doc.EndCommandGroup
    Application.Refresh
End Sub

Private Sub btn_sizes_up_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    make_sizes_sep "up", Shift
End Sub
Private Sub btn_sizes_dn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    make_sizes_sep "dn", Shift
End Sub
Private Sub btn_sizes_lf_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    make_sizes_sep "lf", Shift
End Sub
Private Sub btn_sizes_ri_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    make_sizes_sep "ri", Shift
End Sub

Private Sub btn_sizes_btw_up_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    make_sizes_sep "upb", Shift
End Sub
Private Sub btn_sizes_btw_dn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    make_sizes_sep "dnb", Shift
End Sub
Private Sub btn_sizes_btw_lf_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    make_sizes_sep "lfb", Shift
End Sub
Private Sub btn_sizes_btw_ri_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    make_sizes_sep "rib", Shift
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
    If shft > 1 Then
        Line_len = API.Set_Space_Width   '// 设置文字空间间隙
    Else
        Line_len = API.Set_Space_Width(True)    '// 只读文字空间间隙
    End If
    
    border = Array(cdrBottomRight, cdrBottomLeft, os.TopY + Line_len, os.TopY + 2 * Line_len, _
                    cdrBottomRight, cdrTopRight, os.LeftX - Line_len, os.LeftX - 2 * Line_len)
                    
    If mirror = True Then border = Array(cdrTopRight, cdrTopLeft, os.BottomY - Line_len, os.BottomY - 2 * Line_len, _
                            cdrBottomLeft, cdrTopLeft, os.RightX + Line_len, os.RightX + 2 * Line_len)
   
        
    If dr = "upbx" Or dr = "upb" Or dr = "dnb" Or dr = "up" Or dr = "dn" Then os.Sort "@shape1.left < @shape2.left"
    If dr = "lfbx" Or dr = "lfb" Or dr = "rib" Or dr = "lf" Or dr = "ri" Then os.Sort "@shape1.top > @shape2.top"
    
    If os.Count > 0 Then
        If os.Count > 1 And Len(dr) > 2 And os.Shapes.Count > 1 Then
            For i = 1 To os.Shapes.Count - 1
                Select Case dr
                    Case "upbx":
                          Set pts = os.Shapes(i).SnapPoints.BBox(border(0))
                          Set pte = os.Shapes(i + 1).SnapPoints.BBox(border(1))
                          Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, 0, border(2), cdrDimensionStyleEngineering)

                          If shft > 0 And i = 1 Then
                            Dimension_SetProperty sh, PresetProperty.value
                            Set pts = os.FirstShape.SnapPoints.BBox(border(0))
                            Set pte = os.LastShape.SnapPoints.BBox(border(1))
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, 0, border(3), cdrDimensionStyleEngineering)
                          End If

                    Case "lfbx":
                          Set pts = os.Shapes(i).SnapPoints.BBox(border(4))
                          Set pte = os.Shapes(i + 1).SnapPoints.BBox(border(5))
                          Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, border(6), 0, cdrDimensionStyleEngineering)
  
                          If shft > 0 And i = 1 Then
                            Dimension_SetProperty sh, PresetProperty.value
                            Set pts = os.FirstShape.SnapPoints.BBox(border(4))
                            Set pte = os.LastShape.SnapPoints.BBox(border(5))
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, border(7), 0, cdrDimensionStyleEngineering)
                          End If
                          
                    Case "upb":
                            Set pts = os.Shapes(i).SnapPoints.BBox(cdrTopRight)
                            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopLeft)
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.TopY + os.SizeHeight / 10, cdrDimensionStyleEngineering)

                    Case "dnb":
                            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomRight)
                            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrBottomLeft)
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.BottomY - os.SizeHeight / 10, cdrDimensionStyleEngineering)
                    
                    Case "lfb":
                            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomLeft)
                            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopLeft)
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, os.LeftX - os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
                    
                    Case "rib":
                            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomRight)
                            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopRight)
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, os.RightX + os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
                End Select
                '// 尺寸标注设置属性
                Dimension_SetProperty sh, PresetProperty.value
                'ActiveDocument.ClearSelection
            Next i
        Else
            If shft > 0 Then
                Select Case dr
                    Case "up":
                            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopLeft)
                            Set pte = os.LastShape.SnapPoints.BBox(cdrTopRight)
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.TopY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
                    
                    Case "dn":
                            Set pts = os.FirstShape.SnapPoints.BBox(cdrBottomLeft)
                            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomRight)
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.BottomY - os.SizeHeight / 10, cdrDimensionStyleEngineering)
                    Case "lf":
                            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopLeft)
                            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomLeft)
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, os.LeftX - os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
                    Case "ri":
                            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopRight)
                            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomRight)
                            Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, os.RightX + os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering)
                End Select
                Dimension_SetProperty sh, PresetProperty.value
            Else
                For Each s In os.Shapes
                    Select Case dr
                        Case "up":
                                Set pts = s.SnapPoints.BBox(cdrTopLeft)
                                Set pte = s.SnapPoints.BBox(cdrTopRight)
                                Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, s.LeftX + s.SizeWidth / 10, s.TopY + s.SizeHeight / 10, cdrDimensionStyleEngineering)
                        
                        Case "dn":
                                Set pts = s.SnapPoints.BBox(cdrBottomLeft)
                                Set pte = s.SnapPoints.BBox(cdrBottomRight)
                                Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, s.LeftX + s.SizeWidth / 10, s.BottomY - s.SizeHeight / 10, cdrDimensionStyleEngineering)
                        Case "lf":
                                Set pts = s.SnapPoints.BBox(cdrTopLeft)
                                Set pte = s.SnapPoints.BBox(cdrBottomLeft)
                                Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, s.LeftX - s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering)
                        Case "ri":
                                Set pts = s.SnapPoints.BBox(cdrTopRight)
                                Set pte = s.SnapPoints.BBox(cdrBottomRight)
                                Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, pte, True, s.RightX + s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering)
                    End Select
                    Dimension_SetProperty sh, PresetProperty.value
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
  If os.Count > 0 Then
  For Each s In os.Shapes
      Set pts = s.SnapPoints.BBox(cdrTopLeft)
      Set pte = s.SnapPoints.BBox(cdrTopRight)
      Set ptle = s.SnapPoints.BBox(cdrBottomLeft)
      If shft <> 6 Then Dimension_SetProperty ActiveLayer.CreateLinearDimension(cdrDimensionVertical, pts, ptle, True, _
                                              s.LeftX - s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering), PresetProperty.value
      If shft <> 3 Then Dimension_SetProperty ActiveLayer.CreateLinearDimension(cdrDimensionHorizontal, pts, pte, True, _
                                              s.LeftX + s.SizeWidth / 10, s.TopY + s.SizeHeight / 10, cdrDimensionStyleEngineering), PresetProperty.value
  Next s
  End If

ErrorHandler:
  API.EndOpt
End Sub

Private Sub btn_join_nodes_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ActiveSelection.CustomCommand "ConvertTo", "JoinCurves"
    Application.Refresh
End Sub

Private Sub btn_nodes_reduce_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    Set doc = ActiveDocument
    Dim s As Shape
    ps = Array(1)
    doc.Unit = cdrTenthMicron
    Set os = ActivePage.Shapes
    If os.Count > 0 Then
        doc.BeginCommandGroup "reduce nodes"
        For Each s In os
            s.ConvertToCurves
            If Not s.DisplayCurve Is Nothing Then
                s.Curve.AutoReduceNodes 50
            End If
        Next s
        doc.EndCommandGroup
    End If
    Application.Refresh
ErrorHandler:
  MsgBox "s.Curve.AutoReduceNodes 只有高版本才支持本API"
End Sub

'// 使用标记线批量建立尺寸标注:   左键上标注，右键右标注
Private Sub MarkLines_Makesize_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  
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
  
  sr.CreateSelection
End Sub

'// 使用手工选节点建立尺寸标注，使用Ctrl分离尺寸标注
Private Sub Manual_Makesize_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
      '// 右键
  ElseIf Shift = fmCtrlMask Then
      Slanted_Makesize  '// 手动标注倾斜尺寸
  Else
      Untie_MarkLines   '// 解绑尺寸，分离尺寸
  End If
End Sub



'// 解绑尺寸，分离尺寸
Private Function Untie_MarkLines()
  Dim os As ShapeRange, dss As New ShapeRange
  Set os = ActiveSelectionRange
  For Each s In os.Shapes
      If s.Type = cdrLinearDimensionShape Then
        dss.Add s
      End If
  Next s
  If dss.Count > 0 Then
    dss.BreakApartEx
    os.Shapes.FindShapes(Query:="@name ='DMKLine'").CreateSelection
    ActiveSelectionRange.Delete
  End If
End Function

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
  If nr.Count < 2 Then Exit Function

  cnt = nr.Count
  While cnt > 1
    x1 = nr(cnt).PositionX
    y1 = nr(cnt).PositionY
    x2 = nr(cnt - 1).PositionX
    y2 = nr(cnt - 1).PositionY
    
    Set pts = CreateSnapPoint(x1, y1)
    Set pte = CreateSnapPoint(x2, y2)
    Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionSlanted, pts, pte, True, x1 - 20, y1 + 20, cdrDimensionStyleEngineering)
    
    Dimension_SetProperty sh, PresetProperty.value
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
  
  sr.Sort "@shape1.left < @shape2.left"

  For i = 1 To sr.Count - 1
    x1 = sr(i + 1).CenterX
    y1 = sr(i + 1).CenterY
    x2 = sr(i).CenterX
    y2 = sr(i).CenterY
    
    Set pts = CreateSnapPoint(x1, y1)
    Set pte = CreateSnapPoint(x2, y2)
    Set sh = ActiveLayer.CreateLinearDimension(cdrDimensionSlanted, pts, pte, True, x1 - 20, y1 + 20, cdrDimensionStyleEngineering)
    
    Dimension_SetProperty sh, PresetProperty.value
  Next i
  sr.Delete

ErrorHandler:
  API.EndOpt
End Function

'// 尺寸标注设置属性
Private Function Dimension_SetProperty(sh_dim As Shape, Optional ByVal Preset As Boolean = False)
  If Preset And sh_dim.Type = cdrLinearDimensionShape Then
    With sh_dim.Style.GetProperty("dimension")
      .SetProperty "precision", 0 '       小数位数
      .SetProperty "showUnits", 0 '       是否显示单位 0/1
      .SetProperty "textPlacement", 0 '   0、上方，1、下方，2、中间
    '  .SetProperty "dynamicText", 0 '    是否可以编辑尺寸0/1
    '  .SetProperty "overhang", 500000 '
    End With
  End If
End Function


Private Sub X_EXIT_Click()
  Unload Me    '// EXIT
End Sub
