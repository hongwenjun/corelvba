Attribute VB_Name = "Tools"
#If VBA7 Then
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
  Private Declare Sub Sleep Lib "kernel32" (ByValdwMilliseconds As Long)
#End If

Public Function wait()
  Sleep 3000
End Function


Public Sub �����������(str)
  Dim s As Shape
  Dim x As Double, Y As Double, Shift As Long
  Dim b As Boolean
  b = ActiveDocument.GetUserClick(x, Y, Shift, 10, False, cdrCursorIntersectSingle)
  
  str = VBA.Replace(str, vbNewLine, Chr(10))
  str = VBA.Replace(str, Chr(10), vbNewLine)
  Set s = ActiveLayer.CreateArtisticText(0, 0, str)
  s.CenterX = x
  s.CenterY = Y
End Sub

Public Sub �ߴ��ע()
  ActiveDocument.Unit = cdrMillimeter
  Set s = ActiveSelection
  x = s.CenterX: Y = s.TopY
  sw = s.SizeWidth: sh = s.SizeHeight
        
  text = Int(sw) & "x" & Int(sh) & "mm"
  Set s = ActiveLayer.CreateArtisticText(0, 0, text)
  s.CenterX = x: s.BottomY = Y + 5
End Sub

Public Sub ������������(str)
  Dim s As Shape, sr As ShapeRange
  Set sr = ActiveSelectionRange
  
  For Each s In sr.Shapes
    x = s.CenterX: Y = s.CenterY
    
    Set s = ActiveLayer.CreateArtisticText(0, 0, str)
    s.CenterX = x: s.CenterY = Y
  Next
End Sub

Public Sub ������ע()
  ActiveDocument.Unit = cdrMillimeter
  Set sr = ActiveSelectionRange
  
  For Each s In sr.Shapes
    x = s.CenterX: Y = s.TopY
    sw = s.SizeWidth: sh = s.SizeHeight
          
    text = Int(sw + 0.5) & "x" & Int(sh + 0.5) & "mm"
    Set s = ActiveLayer.CreateArtisticText(0, 0, text)
    s.CenterX = x: s.BottomY = Y + 5
  Next
End Sub

Public Sub ����Ⱥ��()
  Set s1 = ActiveSelectionRange.CustomCommand("Boundary", "CreateBoundary")
  Set brk1 = s1.BreakApartEx

  For Each s In brk1
    Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, True)
    sh.Shapes.All.Group
    s.Delete
  Next
End Sub


' ʵ��Ӧ��: ѡ�����Ⱥ��,ҳ�����������С,���ҳ�����
Public Function Ⱥ�����ҳ��()
  ActiveDocument.Unit = cdrMillimeter
  Dim OrigSelection As ShapeRange, sh As Shape
  Set OrigSelection = ActiveSelectionRange
  Set sh = OrigSelection.Group
  
  ' MsgBox "ѡ������ߴ�: " & sh.SizeWidth & "x" & sh.SizeHeight
  ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
  
#If VBA7 Then
  ActiveDocument.ClearSelection
  sh.AddToSelection
  ActiveSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
#Else
  sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
#End If

End Function


Public Function ������ҳ����()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True

  ActiveDocument.Unit = cdrMillimeter
  Set sr = ActiveSelectionRange
  total = sr.Count

  '// ������ҳ��
  Set doc = ActiveDocument
  doc.AddPages (total - 1)

  Dim sh As Shape
  
  '// ����������������������ҳ��
  For i = 1 To sr.Count
    doc.Pages(i).Activate
    Set sh = sr.Shapes(i)
    ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
 
   '// �������ҳ��
#If VBA7 Then
  ActiveDocument.ClearSelection
  sh.AddToSelection
  ActiveSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
#Else
  sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
#End If

  Next i

  ActiveDocument.EndCommandGroup: Application.Optimization = False
  ActiveWindow.Refresh:   Application.Refresh
Exit Function

ErrorHandler:
  Application.Optimization = False
  MsgBox "����ѡ��һЩ���"
  On Error Resume Next
End Function


'// ��ȫ��: ���һ�ν��������ߣ��ٵ�������ο���
Public Function guideangle(actnumber As ShapeRange, cardblood As Integer)
  Dim sr As ShapeRange
  Set sr = ActiveDocument.MasterPage.GuidesLayer.FindShapes(Type:=cdrGuidelineShape)
  If sr.Count <> 0 Then
    sr.Delete
    Exit Function
  End If
  
  If 0 = ActiveSelectionRange.Count Then Exit Function
  ActiveDocument.Unit = cdrMillimeter

  With actnumber
    Set s1 = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(0, .TopY - cardblood, 0#)
    Set s1 = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(0, .BottomY + cardblood, 0#)
    Set s1 = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(.LeftX + cardblood, 0, 90#)
    Set s1 = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(.RightX - cardblood, 0, 90#)
  End With
  
End Function

Public Function splash_cnt()
  splash.Show 0
  splash.text1 = splash.text1 & ">"
  Sleep 100
End Function


Public Function vba_cnt()
  ' VBA_FORM.text1 = VBA_FORM.text1 & ">"
  Sleep 100
End Function

Public Function ���������(space_width As Double)
  If 0 = ActiveSelectionRange.Count Then Exit Function
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.ReferencePoint = cdrCenter
  
  Set ssr = ActiveSelectionRange
  cnt = 1

#If VBA7 Then
  ssr.Sort "@shape1.width * @shape1.height < @shape2.width * @shape2.height"
#Else
' X4 ��֧�� ShapeRange.sort
#End If

  Dim str As String, size As String
  For Each sh In ssr
    size = Int(sh.SizeWidth + 0.5) & "x" & Int(sh.SizeHeight + 0.5) & "mm"
    sh.SetSize Int(sh.SizeWidth + 0.5), Int(sh.SizeHeight + 0.5)
    str = str & size & vbNewLine
  Next sh

  ActiveDocument.ReferencePoint = cdrTopLeft
  
  For Each s In ssr
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).LeftX, ssr(cnt - 1).BottomY - space_width
    cnt = cnt + 1
    
    vba_cnt

  Next s


'  д�ļ�������EXCEL��ͳ��
'  Set fs = CreateObject("Scripting.FileSystemObject")
'  Set f = fs.CreateTextFile("D:\size.txt", True)
'  f.WriteLine str: f.Close

  str = �������(str)
  Debug.Print str

  Dim s1 As Shape
' Set s1 = ActiveLayer.CreateParagraphText(0, 0, 100, 150, Str, Font:="��������")
  x = ssr.FirstShape.LeftX - 100
  Y = ssr.FirstShape.TopY
  Set s1 = ActiveLayer.CreateParagraphText(x, Y, x + 90, Y - 150, str, Font:="��������")
End Function
 
'// ʵ��Excel�������ܹ���
Private Function �������(str As String) As String
  Dim a, b, d, arr
  str = VBA.Replace(str, vbNewLine, " ")
  Do While InStr(str, "  ")
      str = VBA.Replace(str, "  ", " ")
  Loop
  arr = Split(str)

  Set d = CreateObject("Scripting.dictionary")

  For i = 0 To UBound(arr) - 1
    If d.Exists(arr(i)) = True Then
      d.Item(arr(i)) = d.Item(arr(i)) + 1
    Else
       d.Add arr(i), 1
    End If
  Next

  str = "   ��   ��" & vbTab & vbTab & vbTab & "����" & vbNewLine

  a = d.keys: b = d.items
  For i = 0 To d.Count - 1
    ' Debug.Print a(i), b(i)
    str = str & a(i) & vbTab & vbTab & b(i) & "��" & vbNewLine
  Next

  ������� = str & "�ϼ�����:" & vbTab & vbTab & vbTab & UBound(arr) & "��" & vbNewLine
End Function


' �����˵������,Ϊ(x1,y1)��(x2,y2) ��ô��Ƕ�a��tanֵ: tana=(y2-y1)/(x2-x1)
' ���Լ���arctan(y2-y1)/(x2-x1), �õ���Ƕ�ֵa
' VB����atn(), ����ֵ�ǻ��ȣ���Ҫ ���� PI /180
Private Function lineangle(x1, y1, x2, y2) As Double
  pi = 4 * VBA.Atn(1) ' ����Բ����
  If x2 = x1 Then
    lineangle = 90: Exit Function
  End If
  lineangle = VBA.Atn((y2 - y1) / (x2 - x1)) / pi * 180
End Function

Public Function �Ƕ�תƽ()
  On Error GoTo ErrorHandler
'  ActiveDocument.ReferencePoint = cdrCenter
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.All

  If nr.Count = 2 Then
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2): sr.Rotate -a
    ' sr.LastShape.Delete   '// ɾ���ο���
  End If
ErrorHandler:
End Function

Public Function �Զ���ת�Ƕ�()
  On Error GoTo ErrorHandler
'  ActiveDocument.ReferencePoint = cdrCenter
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.All

  If nr.Count = 2 Then
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2): sr.Rotate 90 + a
    sr.LastShape.Delete   '// ɾ���ο���
  End If
ErrorHandler:
End Function


Public Function ��������()
  Set sr = ActiveSelectionRange
  If sr.Count = 2 Then
    x = sr.LastShape.CenterX: Y = sr.LastShape.CenterY
    sr.LastShape.CenterX = sr.FirstShape.CenterX: sr.LastShape.CenterY = sr.FirstShape.CenterY
    sr.FirstShape.CenterX = x: sr.FirstShape.CenterY = Y
  End If
End Function

Public Function �ο��߾���()
  On Error GoTo ErrorHandler
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.All

  If nr.Count = 2 Then
    ActiveDocument.BeginCommandGroup "Mirror"
    byshape = False
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2)  '// �ο��ߺ�ˮƽ�ļн� a
    sr.Remove sr.Count
    
    ang = 90 - a  ' �������ת�Ƕ�
    For Each s In sr
      With s
        .Duplicate   ' // �������������Ȼ�� x1,y1 �� ��ת
        .RotationCenterX = x1
        .RotationCenterY = y1
        .Rotate ang
        If Not byshape Then
            lx = .LeftX
            .Stretch -1#, 1#    ' // ͨ��������ɾ���
            .LeftX = lx
            .Move (x1 - .LeftX) * 2 - .SizeWidth, 0
            .RotationCenterX = x1   '// ֮ǰ��Ϊ������ת���ĵ㷴�ˣ����û���
            .RotationCenterY = y1
            .Rotate -ang
        End If
        .RotationCenterX = .CenterX   '// ���û���ת���ĵ�Ϊ�������
        .RotationCenterY = .CenterY
      End With
    Next s
    ActiveDocument.EndCommandGroup
  End If
ErrorHandler:
End Function


Public Function autogroup(Optional Group As String = "group", Optional shft = 0, Optional sss As Shapes = Nothing, Optional undogroup = True) As ShapeRange
  Dim sr As ShapeRange, sr_all As ShapeRange, os As ShapeRange
  Dim sp As SubPaths
  Dim arr()
  Dim s As Shape
  If sss Is Nothing Then Set os = ActiveSelectionRange Else Set os = sss.All
  On Error GoTo errn
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  
  If ActiveSelection.Shapes.Count > 0 Then
    gcnt = os.Shapes.Count
    ReDim arr(1 To gcnt, 1 To gcnt)
    Set sr_all = ActiveSelectionRange
    sr_all.RemoveAll
    ReDim arr(1 To gcnt, 1 To gcnt)
    ActiveDocument.Unit = cdrTenthMicron
    sgap = 10
    If shft = 2 Or shft = 3 Or shft = 6 Or shft = 7 Then
      os.RemoveAll
      For Each s In ActiveSelectionRange.Shapes
          os.Add ActivePage.SelectShapesFromRectangle(s.LeftX - sgap, s.BottomY - sgap, s.RightX + sgap, s.TopY + sgap, True)
      Next s
    End If
    
    For i = 1 To os.Shapes.Count
      Set s1 = os.Shapes(i)
      arr(i, i) = i
      For j = 1 To os.Shapes.Count
        Set s2 = os.Shapes(j)
        If s2.LeftX < s1.RightX + sgap And s2.RightX > s1.LeftX - sgap And s2.BottomY < s1.TopY + sgap And s2.TopY > s1.BottomY - sgap Then
          If shft = 1 Or shft = 3 Or shft = 5 Or shft = 7 Then
            Set isec = s1.Intersect(s2)
            If Not isec Is Nothing Then
              arr(i, j) = j
              isec.CreateSelection
              isec.Delete
            End If
          Else
            arr(i, j) = j
          End If
        End If
      Next j
    Next i
    
    For i = 1 To gcnt
      arr = collect_arr(arr, i, i)
    Next i
    
    Set sr = ActiveSelectionRange

    For i = 1 To gcnt
      sr.RemoveAll
      inar = 0
      For j = 1 To gcnt
        If arr(i, j) > 0 Then
          sr.Add os.Shapes(j)
          inar = inar + 1
        End If
      Next j
      If inar > 1 Then
        If Group = "group" Then
          If shft < 4 Then sr_all.Add sr.Group
        End If
      Else
        If sr.Shapes.Count > 0 Then sr_all.AddRange sr
      End If
    Next i
  Set autogroup = sr_all
  End If

  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
  Exit Function
errn:
  Application.Optimization = False
End Function

Public Function collect_arr(arr, ci, ki)
    lim = UBound(arr)
    For k = 1 To lim
        If arr(ki, k) > 0 Then
            arr(ci, k) = k
            If ki <> ci Then arr(ki, k) = Empty
            If ci <> k And ki <> k Then arr = collect_arr(arr, ci, k)
        End If
    Next k
    'If ki <> ci Then arr(ki, ki) = Empty
    collect_arr = arr
End Function



Sub Make_Sizes()
    ActiveDocument.Unit = cdrMillimeter
    Set os = ActiveSelectionRange
    If os.Count > 0 Then
    For Each s In os.Shapes
      Set pts = os.FirstShape.SnapPoints.BBox(cdrTopLeft)
      Set pte = os.LastShape.SnapPoints.BBox(cdrTopRight)
      ActiveLayer.CreateLinearDimension cdrDimensionHorizontal, pts, pte, True, os.LeftX + os.SizeWidth / 10, os.TopY + os.SizeHeight / 10, cdrDimensionStyleEngineering
      
      Set pte = os.LastShape.SnapPoints.BBox(cdrBottomLeft)
      ActiveLayer.CreateLinearDimension cdrDimensionVertical, pts, pte, True, os.LeftX - os.SizeWidth / 10, os.BottomY + os.SizeHeight / 10, cdrDimensionStyleEngineering
      
    Next s
    End If
End Sub

'''////  ѡ�����������Ȼ�����߶Σ�Ϊ��������׼��  ////'''
Public Function Split_Segment()
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  
  Dim ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  Dim s As Shape
  Dim nr As NodeRange
  Dim nd As Node
  
  Set s = ssr.UngroupAllEx.Combine
  Set nr = s.Curve.Nodes.All
  
  nr.BreakApart
  s.BreakApartEx
'  For Each nd In nr
'    nd.BreakApart
'  Next nd
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
Exit Function
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function


'// �޸�Բ��ȱ�ǵ�ֱ��
Public Sub corner_off()
    Dim os As ShapeRange
    Dim s As Shape, fir As Shape, ci As Shape
    Dim nd As Node, nds As Node, nde As Node

    Set os = ActiveSelectionRange
    ud = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
On Error GoTo errn
    ActiveDocument.BeginCommandGroup "corners off"
    'Application.Optimization = True
    selec = False
    If os.Shapes.Count = 1 Then
        Set s = os.FirstShape
        If Not s.Curve Is Nothing Then
            For Each nd In s.Curve.Nodes
                If nd.Selected Then
                    selec = True
                    Exit For
                End If
            Next nd
        End If
    End If
    
    If os.Shapes.Count > 1 Or Not selec Then
        os.ConvertToCurves
        For Each s In os.Shapes
            Set nds = Nothing
            Set nde = Nothing
            For k = 1 To 3
            For i = 1 To s.Curve.Nodes.Count
                If i <= s.Curve.Nodes.Count Then
                    Set nd = s.Curve.Nodes(i)
                    If Not nd.NextSegment Is Nothing And Not nd.PrevSegment Is Nothing Then
                        If Abs(nd.PrevSegment.Length - nd.NextSegment.Length) < (nd.PrevSegment.Length + nd.NextSegment.Length) / 30 And nd.PrevSegment.Type = cdrCurveSegment And nd.NextSegment.Type = cdrCurveSegment Then
                            corner_off_make s, nd.Previous, nd.Next
                        ElseIf Not nd.Next.NextSegment Is Nothing Then
                            If (nd.PrevSegment.Type = cdrLineSegment Or Abs(Abs(nd.PrevSegment.StartingControlPointAngle - nd.PrevSegment.EndingControlPointAngle) - 180) < 1) _
                                And (nd.Next.NextSegment.Type = cdrLineSegment Or Abs(Abs(nd.Next.NextSegment.StartingControlPointAngle - nd.Next.NextSegment.EndingControlPointAngle) - 180) < 1) _
                                And nd.NextSegment.Type = cdrCurveSegment Then
                                    corner_off_make s, nd, nd.Next
                            End If
                       End If
                    End If
                End If
            Next i
            Next k
            
             
        Next s
    ElseIf os.Shapes.Count = 1 And selec Then
        Set nds = Nothing
        Set nde = Nothing
        For Each nd In s.Curve.Nodes
            If Not nd.Selected And Not nd.Next.Selected Then Exit For
        Next nd
        If Not nd Is s.Curve.Nodes.Last Then
            For i = 1 To s.Curve.Nodes.Count
                Set nd = nd.Next
                If Not nde Is Nothing And Not nds Is Nothing And Not nd.Selected Then Exit For
                If Not nds Is Nothing And nd.Selected Then Set nde = nd
                If nde Is Nothing And nds Is Nothing And nd.Selected Then Set nds = nd
            Next i
            
            If Not nds Is Nothing And Not nde Is Nothing Then
                'ActiveLayer.CreateEllipse2 nds.PositionX, nds.PositionY, nde.PrevSegment.Length / 4
                'ActiveLayer.CreateEllipse2 nde.PositionX, nde.PositionY, nde.PrevSegment.Length / 4
                corner_off_make s, nds, nde
            End If
        End If
    End If
errn:
    Application.Optimization = False
    ActiveDocument.EndCommandGroup
    Application.Refresh
    ActiveDocument.Unit = ud
End Sub

Private Sub corner_off_make(s As Shape, nds As Node, nde As Node)
    Dim l1 As Shape, l2 As Shape
    Dim os As ShapeRange
    Dim ss As Shape
    ud = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter

    Set l1 = ActiveLayer.CreateLineSegment(nds.PositionX, nds.PositionY, nds.PositionX + s.SizeWidth * 3, nds.PositionY)
    l1.RotationCenterX = nds.PositionX
    l1.RotationAngle = nds.PrevSegment.EndingControlPointAngle + 180
    
    Set l2 = ActiveLayer.CreateLineSegment(nde.PositionX, nde.PositionY, nde.PositionX + s.SizeWidth * 3, nde.PositionY)
    l2.RotationCenterX = nde.PositionX
    l2.RotationAngle = nde.NextSegment.StartingControlPointAngle + 180
    
    Set lcross = l2.Curve.Segments.First.GetIntersections(l1.Curve.Segments.First)
    If lcross.Count > 0 Then
        cx = lcross(1).PositionX
        cy = lcross(1).PositionY
        sx = nds.PositionX
        sy = nds.PositionY
        ex = nde.PositionX
        ey = nde.PositionY
        
        l1.Curve.Nodes.Last.PositionX = cx
        l1.Curve.Nodes.Last.PositionY = cy
        l2.Curve.Nodes.Last.PositionX = cx
        l2.Curve.Nodes.Last.PositionY = cy
        
        s.Curve.Nodes.Range(Array(nds.AbsoluteIndex, nde.AbsoluteIndex)).BreakApart
        Set os = s.BreakApartEx
        oscnt = os.Shapes.Count
        For Each ss In os.Shapes
            If ss.Curve.Nodes.First.PositionX = ex And ss.Curve.Nodes.First.PositionY = ey Then Set s2 = ss
            If ss.Curve.Nodes.Last.PositionX = sx And ss.Curve.Nodes.Last.PositionY = sy Then Set s1 = ss
            If ss.Curve.Nodes.First.PositionX = sx And ss.Curve.Nodes.First.PositionY = sy Then ss.Delete
        Next ss
        
        If s1.Curve.Segments.Last.Type = cdrLineSegment Or Abs(Abs(s1.Curve.Segments.Last.StartingControlPointAngle - s1.Curve.Segments.Last.EndingControlPointAngle) - 180) < 1 Then
            s1.Curve.Nodes.Last.PositionX = lcross(1).PositionX
            s1.Curve.Nodes.Last.PositionY = lcross(1).PositionY
            l1.Delete
        Else
            Set s1 = l1.Weld(s1)
        End If
        If oscnt = 2 Then Set s2 = s1
        If s2.Curve.Segments.First.Type = cdrLineSegment Or Abs(Abs(s2.Curve.Segments.First.StartingControlPointAngle - s2.Curve.Segments.First.EndingControlPointAngle) - 180) < 1 Then
            s2.Curve.Nodes.First.PositionX = lcross(1).PositionX
            s2.Curve.Nodes.First.PositionY = lcross(1).PositionY
            l2.Delete
        Else
            Set s2 = l2.Weld(s2)
        End If
        If oscnt > 2 Then Set s2 = s1.Weld(s2)
        s2.CustomCommand "ConvertTo", "JoinCurves", 0.1
        Set s = s2
    Else
        l1.Delete
        l2.Delete
    End If
    ActiveDocument.Unit = ud
End Sub

Sub ExportNodePositions()
    Dim s As Shape, n As Node
    Dim srActiveLayer As ShapeRange
    Dim x As Double, Y As Double
    Dim strNodePositions As String
    
    ActiveDocument.Unit = cdrMillimeter
    
    'Get all the curve shapes on the Active Layer
    '��ȡActive Layer�ϵ�����������״
    Set srActiveLayer = ActiveLayer.Shapes.FindShapes(Query:="@type='curve'")
    'This is another way you can get only the curve shapes
    '������һ����ֻ�ܵõ�������״�ķ���
    'Set srActiveLayer = ActiveLayer.Shapes.FindShapes.FindAnyOfType(cdrCurveShape)
    
    'Loop through each curve
    '����ÿ������
    For Each s In srActiveLayer.Shapes
        'Loop though each node in the curve and get the position
        '���������е�ÿ���ڵ㲢��ȡλ��
        For Each n In s.Curve.Nodes
            n.GetPosition x, Y
            strNodePositions = strNodePositions & "x: " & x & " y: " & Y & vbCrLf
        Next n
    Next s
    
    'Save the node positions to a file
    '���ڵ�λ�ñ��浽�ļ�
    Open "C:\Temp\NodePositions.txt" For Output As #1
        Print #1, strNodePositions
    Close #1
End Sub

Sub ������T()
   Dim mark As Shape
   Dim sr As ShapeRange
   
    Set sr = ActiveSelectionRange
        If (Shift And 1) <> 0 Then ActivePage.Shapes.FindShapes(Query:="@type ='rectangle'or @type ='curve'or @type ='Ellipse'or @type ='Polygon'").CreateSelection
        sr.Shapes.FindShapes(Query:="@type ='rectangle'or @type ='curve'or @type ='Ellipse'or @type ='Polygon'").ConvertToCurves
   If sr.Count = 0 Then Exit Sub
   
    ' CorelDRAW����ԭ���ǵ���DXFʹ��
    
    ' ����ԭ���ǣ��������������������
    Dim MarkPos_Array() As Double
    MarkPos_Array = Get_MarkPosition
    AtOrigin MarkPos_Array(0), MarkPos_Array(1)
    
    sr.Add ActiveDocument.ActiveShape
     Set mark = ActiveDocument.ActiveShape
   ActiveDocument.ClearSelection
   sr.CreateSelection
 '    Set mark = ActiveDocument.ActiveShape
 '  If FileExists("d:\mytempdxf.dxf") Then
 '   DeleteFile "d:\mytempdxf.dxf"
 '  End If
    
 SaveDXF "d:\mytempdxf.dxf"
 
 '  Do While FileExists("d:\mytempdxf.dxf") = False
 '       DoEvents
 '       Delay 1
 '   Loop
 Shell Application.GMSManager.GMSPath & "tuznr.exe d:/mytempdxf.dxf", 1
    
 mark.Delete
End Sub

Sub SaveDXF(FileName As String)
    Dim expopt As StructExportOptions
    Set expopt = CreateStructExportOptions
    expopt.UseColorProfile = False
    Dim expflt As ExportFilter
    Set expflt = ActiveDocument.ExportEx(FileName, cdrDXF, cdrSelection, expopt)
    With expflt
        .BitmapType = 0 ' FilterDXFLib.dxfBitmapJPEG
        .TextAsCurves = True
        .Version = 3 ' FilterDXFLib.dxfVersion13
        .Units = 3 ' FilterDXFLib.dxfMillimeters
        .FillUnmapped = True
        .Finish
    End With
End Sub

' ����ԭ���Ǻ������������������������
Sub AtOrigin(Optional px As Double = 0#, Optional py As Double = 0#)
  Dim doc As Document: Set doc = ActiveDocument
  doc.Unit = cdrMillimeter

  '// ����ԭ���Ǳ���ļ� OriginMark.cdr ��ɢȺ��
  doc.ActiveLayer.Import path & "GMS\OriginMark.cdr"
  doc.ReferencePoint = cdrCenter
  doc.Selection.Ungroup

  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  '// �� MarkName ���Ʋ��� ������
  For Each sh In shs
    If "AtOrigin" = sh.ObjectData("MarkName").Value Then
      sh.SetPosition px, py
    Else
      sh.Delete   ' ����Ҫ�ı��ɾ��
    End If
  Next sh
End Sub

' ʹ�� GlobalUserData ���󱣴� Mark��������ı������ú����������ı�
Public Function Mark_SetPosition() As String
  Dim text As String
  If GlobalUserData.Exists("MarkPosition", 1) Then
    text = GlobalUserData("MarkPosition", 1)
  End If
  text = InputBox("������Mark�������(x,y),�ո�򶺺ż��", "����Mark�������(x,y),��λ(mm)", text)
  If text = "" Then Exit Function
  GlobalUserData("MarkPosition", 1) = text
  Mark_SetPosition = text
End Function

' ��������Mark������깦�ܣ����� ����(x,y)
Public Function Get_MarkPosition() As Double()
  Dim MarkPos_Array(0 To 1) As Double
  Dim str, arr
  
  str = Mark_SetPosition

  ' �滻 ���� Ϊ�ո�
  str = VBA.Replace(str, ",", " ")
  Do While InStr(str, "  ") '����ո񻻳�һ���ո�
      str = VBA.Replace(str, "  ", " ")
  Loop
  arr = Split(str)
  
  MarkPos_Array(0) = Val(arr(0))
  MarkPos_Array(1) = Val(arr(1))
  
  Debug.Print MarkPos_Array(0), MarkPos_Array(1)  ' ��ͼ->�������ڣ�������ʾ
  
  Get_MarkPosition = MarkPos_Array
  
End Function

Public Function SetNames()
  Dim ssr As ShapeRange
  Set ssr = ActiveSelectionRange

#If VBA7 Then
  ssr.Sort " @shape1.left<@shape2.left"
#Else
' X4 ��֧�� ShapeRange.sort
#End If

  Dim text As String
  Dim lines() As String
  ' ��ȡ�ı���Ϣ���и��ı�
  If ssr(1).Type = cdrTextShape Then
    If ssr(1).text.Type = cdrArtistic Then
      text = ssr(1).text.Story.text
      lines = Split(text, vbCr)
      ssr.Remove 1
  #If VBA7 Then
      ssr.Sort " @shape1.top>@shape2.top"
  #Else
  ' X4 ��֧�� ShapeRange.sort
  #End If
    End If
  Else
      MsgBox "��Ѷ����ı��������"
      Exit Function
  End If
    
' Debug.Print ssr.Count, UBound(lines), LBound(lines)
' ������������ƣ��ô�:��������������һ������
  i = 0
  If ssr.Count <= UBound(lines) + 1 Then
    For Each s In ssr
      s.name = lines(i)
      i = i + 1
    Next s
  End If
  
  If ssr.Count <> UBound(lines) + 1 Then MsgBox "�ı���:" & (UBound(lines) + 1) & vbNewLine & "�ұ����:" & ssr.Count
    
End Function

Sub Nodes_TO_TSP()
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile("C:\TSP\CDR_TO_TSP", True)
    ActiveDocument.Unit = cdrMillimeter

    Dim s As Shape, ssr As ShapeRange
    Set ssr = ActiveSelectionRange

    Dim TSP As String
    TSP = (ssr.Count * 4) & " " & 0 & vbNewLine

    For Each s In ssr
        lx = s.LeftX:   rx = s.RightX
        By = s.BottomY: ty = s.TopY
        TSP = TSP & lx & " " & By & vbNewLine
        TSP = TSP & lx & " " & ty & vbNewLine
        TSP = TSP & rx & " " & By & vbNewLine
        TSP = TSP & rx & " " & ty & vbNewLine
    Next s
    f.WriteLine TSP
    f.Close
End Sub


