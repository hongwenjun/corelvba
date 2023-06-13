VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Woodman 
   Caption         =   "批量标注尺寸节点"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   OleObjectBlob   =   "Woodman.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Woodman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_square_hi_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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


Private Sub btn_square_wi_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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

Private Sub btn_makesizes_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
        On Error GoTo 0
    Else
        make_sizes Shift
    End If
    doc.EndCommandGroup
    Application.Refresh
End Sub

Private Sub btn_sizes_up_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    make_sizes_sep "up", Shift
End Sub
Private Sub btn_sizes_dn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    make_sizes_sep "dn", Shift
End Sub
Private Sub btn_sizes_lf_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    make_sizes_sep "lf", Shift
End Sub
Private Sub btn_sizes_ri_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    make_sizes_sep "ri", Shift
End Sub

Private Sub btn_sizes_btw_up_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    make_sizes_sep "upb", Shift
End Sub
Private Sub btn_sizes_btw_dn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    make_sizes_sep "dnb", Shift
End Sub
Private Sub btn_sizes_btw_lf_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    make_sizes_sep "lfb", Shift
End Sub
Private Sub btn_sizes_btw_ri_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    make_sizes_sep "rib", Shift
End Sub

Sub make_sizes_sep(dr, Optional shft = 0)
    Set doc = ActiveDocument
    Dim s As Shape
    Dim pts As New SnapPoint, pte As New SnapPoint
    Dim os As ShapeRange
    un = doc.Unit
    doc.Unit = cdrMillimeter
    doc.BeginCommandGroup "make sizes"
    
    Set os = ActiveSelectionRange
    
    Dim border As Variant
    Dim Line_len As Double
    Line_len = API.GetSet("Line_len")
    
    border = Array(cdrBottomRight, cdrBottomLeft, os.TopY + 10, os.TopY + 20 + Line_len, _
                    cdrBottomRight, cdrTopRight, os.LeftX - 10, os.LeftX - 20 - Line_len)
                    
    If chkOpposite.value Then border = Array(cdrTopRight, cdrTopLeft, os.BottomY - 10, os.BottomY - 20 - Line_len, _
                            cdrBottomLeft, cdrTopLeft, os.RightX + 10, os.RightX + 20 + Line_len)
   
        
    If dr = "upbx" Or dr = "upb" Or dr = "dnb" Or dr = "up" Or dr = "dn" Then os.Sort "@shape1.left < @shape2.left"
    If dr = "lfbx" Or dr = "lfb" Or dr = "rib" Or dr = "lf" Or dr = "ri" Then os.Sort "@shape1.top > @shape2.top"
    
    If os.Count > 0 Then
        If os.Count > 1 And Len(dr) > 2 Then
            For i = 1 To os.Shapes.Count - 1
                Select Case dr
                    Case "upbx":
                          Set pts = os.Shapes(i).SnapPoints.BBox(border(0))
                          Set pte = os.Shapes(i + 1).SnapPoints.BBox(border(1))
                          ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, 0, border(2), cdrDimensionStyleEngineering
                          If shft > 0 And i = 1 Then
                            Set pts = os.FirstShape.SnapPoints.BBox(border(0))
                            Set pte = os.LastShape.SnapPoints.BBox(border(1))
                            ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, 0, border(3), cdrDimensionStyleEngineering
                          End If
                          
                    Case "lfbx":
                          Set pts = os.Shapes(i).SnapPoints.BBox(border(4))
                          Set pte = os.Shapes(i + 1).SnapPoints.BBox(border(5))
                          ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, border(6), 0, cdrDimensionStyleEngineering
                          If shft > 0 And i = 1 Then
                            Set pts = os.FirstShape.SnapPoints.BBox(border(4))
                            Set pte = os.LastShape.SnapPoints.BBox(border(5))
                           ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, border(7), 0, cdrDimensionStyleEngineering
                          End If
                          
                    Case "upb":
                            Set pts = os.Shapes(i).SnapPoints.BBox(cdrTopRight)
                            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopLeft)
                            ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.TopY + os.SizeHeight / 10, cdrDimensionStyleEngineering

                    Case "dnb":
                            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomRight)
                            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrBottomLeft)
                            ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.BottomY - os.SizeHeight / 10, cdrDimensionStyleEngineering
                    
                    Case "lfb":
                            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomLeft)
                            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopLeft)
                            ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, os.LeftX - os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering
                    
                    Case "rib":
                            Set pts = os.Shapes(i).SnapPoints.BBox(cdrBottomRight)
                            Set pte = os.Shapes(i + 1).SnapPoints.BBox(cdrTopRight)
                            ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, os.RightX + os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering
                End Select
                'ActiveDocument.ClearSelection
            Next i
        Else
            If shft > 0 Then
                Select Case dr
                    Case "up":
                            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopLeft)
                            Set pte = os.LastShape.SnapPoints.BBox(cdrTopRight)
                            ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.TopY + os.SizeHeight / 10, cdrDimensionStyleEngineering
                    
                    Case "dn":
                            Set pts = os.FirstShape.SnapPoints.BBox(cdrBottomLeft)
                            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomRight)
                            ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.BottomY - os.SizeHeight / 10, cdrDimensionStyleEngineering
                    Case "lf":
                            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopLeft)
                            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomLeft)
                            ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, os.LeftX - os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering
                    Case "ri":
                            Set pts = os.FirstShape.SnapPoints.BBox(cdrTopRight)
                            Set pte = os.LastShape.SnapPoints.BBox(cdrBottomRight)
                            ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, os.RightX + os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering
                End Select
            Else
                For Each s In os.Shapes
                    Select Case dr
                        Case "up":
                                Set pts = s.SnapPoints.BBox(cdrTopLeft)
                                Set pte = s.SnapPoints.BBox(cdrTopRight)
                                ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, s.LeftX + s.SizeWidth / 10, s.TopY + s.SizeHeight / 10, cdrDimensionStyleEngineering
                        
                        Case "dn":
                                Set pts = s.SnapPoints.BBox(cdrBottomLeft)
                                Set pte = s.SnapPoints.BBox(cdrBottomRight)
                                ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, s.LeftX + s.SizeWidth / 10, s.BottomY - s.SizeHeight / 10, cdrDimensionStyleEngineering
                        Case "lf":
                                Set pts = s.SnapPoints.BBox(cdrTopLeft)
                                Set pte = s.SnapPoints.BBox(cdrBottomLeft)
                                ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, s.LeftX - s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering
                        Case "ri":
                                Set pts = s.SnapPoints.BBox(cdrTopRight)
                                Set pte = s.SnapPoints.BBox(cdrBottomRight)
                                ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, s.RightX + s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering
                    End Select
                Next s
            End If
        End If
    End If
    os.CreateSelection
    doc.EndCommandGroup
    doc.Unit = un
End Sub

Sub make_sizes(Optional shft = 0)
    Set doc = ActiveDocument
    Dim s As Shape
    Dim pts As SnapPoint, pte As SnapPoint
    Dim os As ShapeRange
    un = doc.Unit
    doc.Unit = cdrMillimeter
    doc.BeginCommandGroup "make sizes"
    Set os = ActiveSelectionRange
    If os.Count > 0 Then
    For Each s In os.Shapes
        Set pts = s.SnapPoints.BBox(cdrTopLeft)
        Set pte = s.SnapPoints.BBox(cdrTopRight)
        Set ptle = s.SnapPoints.BBox(cdrBottomLeft)
        If shft <> 6 Then ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, ptle, True, s.LeftX - s.SizeWidth / 10, s.BottomY + s.SizeHeight / 10, cdrDimensionStyleEngineering
        If shft <> 3 Then ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, s.LeftX + s.SizeWidth / 10, s.TopY + s.SizeHeight / 10, cdrDimensionStyleEngineering
    Next s
    End If
    doc.EndCommandGroup
    doc.Unit = un
End Sub

Public Function make_selection(Optional mode = "fcolor", Optional sel = True, Optional OSS As ShapeRange = Nothing, Optional colr = Nothing) As ShapeRange
    Dim s As Shape, lst As Shape
    Dim sr As ShapeRange
    'Dim os As ShapeRange
    Set doc = ActiveDocument
    doc.Unit = cdrTenthMicron
    
    If OSS Is Nothing Then
        If toolspanel.num_list.value Or mode = "locked" Then
            Set os = ActivePage
        Else
            Set os = ActiveSelectionRange
        End If
    Else
        Set os = OSS
    End If
    Set sr = ActiveSelectionRange
    sr.RemoveAll
    If sel Then ActiveDocument.ClearSelection
    Set lst = os.Shapes.First
    For Each s In os.Shapes
        Select Case mode
            Case "ocolor": If s.Outline.Type <> cdrNoOutline And s.Shapes.Count = 0 And s.Outline.Color.HexValue = colr.HexValue Then sr.Add s
            Case "fcolor": If s.Fill.Type <> cdrNoFill And s.Shapes.Count = 0 And s.Fill.UniformColor.HexValue = colr.HexValue Then sr.Add s
            Case "nofil": If s.Fill.Type = cdrNoFill And s.Shapes.Count = 0 Then sr.Add s
            Case "fil": If s.Fill.Type <> cdrNoFill And s.Shapes.Count = 0 Then sr.Add s
            Case "abr": If s.Outline.Type <> cdrNoOutline And s.Shapes.Count = 0 Then sr.Add s
            Case "noabr": If s.Outline.Type = cdrNoOutline And s.Shapes.Count = 0 Then sr.Add s
            Case "open": If Not s.DisplayCurve Is Nothing Then If Not s.DisplayCurve.Closed Then sr.Add s
            Case "closed": If Not s.DisplayCurve Is Nothing Then If s.DisplayCurve.Closed Then sr.Add s
            Case "single": If s.Shapes.Count = 0 Then sr.Add s
            Case "dashed": If s.Outline.Style.DashCount > 0 Then sr.Add s
            Case "groups": If s.Shapes.Count > 0 And s.Effect Is Nothing Then sr.Add s
            Case "text": If s.Shapes.Count = 0 And s.Type = cdrTextShape Then sr.Add s
            Case "notext": If s.Shapes.Count = 0 And s.Type <> cdrTextShape Then sr.Add s
            Case "images": If s.Type = cdrBitmapShape Then sr.Add s
            Case "locked": If s.Locked Then sr.Add s
            Case "effects": If s.Effects.Count > 0 Or Not s.Effect Is Nothing Then sr.Add s
            Case "noeffects": If s.Effects.Count = 0 And s.Effect Is Nothing Then sr.Add s
            Case "bigger":
                arelst = lst.SizeHeight * lst.SizeWidth
                ares = s.SizeHeight * s.SizeWidth
                If ares >= arelst Then
                    are = one_shape_area(lst)
                    If one_shape_area(s) >= are Then sr.Add s
                End If
            Case "smaller":
                arelst = lst.SizeHeight * lst.SizeWidth
                ares = s.SizeHeight * s.SizeWidth
                If ares <= arelst Then
                    are = one_shape_area(lst)
                    If one_shape_area(s) <= are Then sr.Add s
                End If
            Case "last":
                If lst.Fill.Type = cdrNoFill Then
                    's.CreateSelection
                    If s.Outline.Type <> cdrNoOutline Then If s.Outline.Color.HexValue = lst.Outline.Color.HexValue Then sr.Add s
                Else
                    If s.Fill.UniformColor.HexValue = lst.Fill.UniformColor.HexValue Then sr.Add s
                End If
        End Select
    Next s
    
    If sr.Shapes.Count > 0 And sel Then sr.CreateSelection
    Set make_selection = sr
    
    Application.Refresh
    ActiveWindow.Activate
End Function

Public Function get_events(btn As String, Optional shft = 0, Optional click = 1)
    out = "ok"
    get_events = out
End Function

Private Sub btn_join_nodes_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ActiveSelection.CustomCommand "ConvertTo", "JoinCurves"
    Application.Refresh
End Sub

Private Sub btn_nodes_reduce_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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


Private Sub MarkLines_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = 2 Then
    CutLines.Dimension_MarkLines cdrAlignLeft, chkOpposite.value
    make_sizes_sep "lfbx", Shift
  Else
    CutLines.Dimension_MarkLines cdrAlignTop, chkOpposite.value
    Label_Makesizes.Caption = "试试右键"
    make_sizes_sep "upbx", Shift
  End If
End Sub

Private Sub chkOpposite_Click()
'  Debug.Print chkOpposite.value
End Sub

Private Sub manual_makesize_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
  ActiveDocument.Unit = cdrMillimeter
  Dim nr As NodeRange, cnt As Integer
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  Set nr = ActiveShape.Curve.Selection
  If nr.Count < 2 Then Exit Function
  cnt = nr.Count
  While cnt > 1
    x1 = nr(cnt).PositionX
    y1 = nr(cnt).PositionY
    x2 = nr(cnt - 1).PositionX
    y2 = nr(cnt - 1).PositionY
    
    Set pts = CreateSnapPoint(x1, y1)
    Set pte = CreateSnapPoint(x2, y2)
    ActiveLayer.CreateLinearDimension cdrDimensionSlanted, pts, pte, True, x1 - 5, y1 + 5, cdrDimensionStyleEngineering
    cnt = cnt - 1
  Wend
ErrorHandler:
End Function

