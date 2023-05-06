VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZCOPY 
   Caption         =   "UserForm1"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   OleObjectBlob   =   "ZCOPY.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ZCOPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_square_hi_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_square_hi", Shift, Button) = "exit" Then Exit Sub
    Set os = ActiveSelectionRange
    Set ss = os.Shapes
    uc = 0
    For Each s In ss
        s.SizeWidth = s.SizeHeight
        uc = uc + 1
    Next s
    Application.Refresh
    If ch_main_switch Then ActiveWindow.Activate
End Sub


Private Sub btn_square_wi_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_square_wi", Shift, Button) = "exit" Then Exit Sub
    Set os = ActiveSelectionRange
    Set ss = os.Shapes
    uc = 0
    For Each s In ss
        s.SizeHeight = s.SizeWidth
        uc = uc + 1
    Next s
    Application.Refresh
    If ch_main_switch Then ActiveWindow.Activate
End Sub

Private Sub btn_makesizes_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_makesizes", Shift, Button) = "exit" Then Exit Sub
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
        Make_Sizes Shift
    End If
    doc.EndCommandGroup
    Application.Refresh
End Sub

Private Sub btn_sizes_up_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_sizes_up", Shift, Button) = "exit" Then Exit Sub
    make_sizes_sep "up", Shift
End Sub
Private Sub btn_sizes_dn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_sizes_dn", Shift, Button) = "exit" Then Exit Sub
    make_sizes_sep "dn", Shift
End Sub
Private Sub btn_sizes_lf_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_sizes_lf", Shift, Button) = "exit" Then Exit Sub
    make_sizes_sep "lf", Shift
End Sub
Private Sub btn_sizes_ri_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_sizes_ri", Shift, Button) = "exit" Then Exit Sub
    make_sizes_sep "ri", Shift
End Sub

Private Sub btn_sizes_btw_up_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_sizes_btw_up", Shift, Button) = "exit" Then Exit Sub
    make_sizes_sep "upb", Shift
End Sub
Private Sub btn_sizes_btw_dn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_sizes_btw_dn", Shift, Button) = "exit" Then Exit Sub
    make_sizes_sep "dnb", Shift
End Sub
Private Sub btn_sizes_btw_lf_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_sizes_btw_lf", Shift, Button) = "exit" Then Exit Sub
    make_sizes_sep "lfb", Shift
End Sub
Private Sub btn_sizes_btw_ri_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If get_events("btn_sizes_btw_ri", Shift, Button) = "exit" Then Exit Sub
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
        
    If dr = "upb" Or dr = "dnb" Or dr = "up" Or dr = "dn" Then os.Sort "@shape1.left < @shape2.left"
    If dr = "lfb" Or dr = "rib" Or dr = "lf" Or dr = "ri" Then os.Sort "@shape1.top > @shape2.top"
    
    If os.Count > 0 Then
        If os.Count > 1 And Len(dr) > 2 Then
            For i = 1 To os.Shapes.Count - 1
                Select Case dr
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

Sub Make_Sizes(Optional shft = 0)
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
        If toolspanel.num_list.Value Or mode = "locked" Then
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
