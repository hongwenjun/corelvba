Attribute VB_Name = "Tools"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// 简易火车排列
Public Function Simple_Train_Arrangement(Space_Width As Double)
  API.BeginOpt
  Dim ssr As ShapeRange, s As Shape
  Dim cnt As Integer
  Set ssr = ActiveSelectionRange
  cnt = 1

#If VBA7 Then
'  ssr.sort " @shape1.top>@shape2.top"
  ssr.Sort " @shape1.left<@shape2.left"
#Else
' X4 不支持 ShapeRange.sort
#End If

  ActiveDocument.ReferencePoint = cdrTopLeft
  For Each s In ssr
    '// 底对齐 If cnt > 1 Then s.SetPosition ssr(cnt - 1).RightX, ssr(cnt - 1).BottomY
    '// 改成顶对齐 2022-08-10
    ActiveDocument.ReferencePoint = cdrTopLeft + cdrBottomTop
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).RightX + Space_Width, ssr(cnt - 1).TopY
    cnt = cnt + 1
  Next s

  API.EndOpt
End Function

'// 简易阶梯排列
Public Function Simple_Ladder_Arrangement(Space_Width As Double)
  API.BeginOpt
  Dim ssr As ShapeRange, s As Shape
  Dim cnt As Integer
  Set ssr = ActiveSelectionRange
  cnt = 1

#If VBA7 Then
  ssr.Sort " @shape1.top>@shape2.top"
'  ssr.sort " @shape1.left<@shape2.left"
#Else
' X4 不支持 ShapeRange.sort
#End If

  ActiveDocument.ReferencePoint = cdrTopLeft
  For Each s In ssr
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).LeftX, ssr(cnt - 1).BottomY - Space_Width
    cnt = cnt + 1
  Next s

  API.EndOpt
End Function

'// 文本转曲线   默认使用简单转曲，参数 all=1 ，支持框选和图框剪裁内的文本
Public Function TextShape_ConvertToCurves(Optional all = 0)
  API.BeginOpt
  On Error GoTo ErrorHandler
  Dim s As Shape, cnt As Long
  
  If all = 1 Then
    For Each s In API.FindAllShapes.Shapes.FindShapes(, cdrTextShape)
      s.ConvertToCurves
      cnt = cnt + 1
    Next s
  Else
  
    For Each s In ActivePage.FindShapes(, cdrTextShape)
      s.ConvertToCurves
      cnt = cnt + 1
    Next s
  End If
ErrorHandler:
  API.EndOpt
End Function

'// 复制物件
Public Function copy_shape()
  Dim OrigSelection As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  OrigSelection.Copy

End Function

'// 旋转物件角度
Public Function Rotate_Shapes(n As Double)
  API.BeginOpt
  
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String, size As String
  For Each sh In shs
    sh.Rotate n
  Next sh
  
  API.EndOpt
End Function

'// 得到物件尺寸
Public Function get_shape_size(ByRef sx As Double, ByRef sy As Double)
  ActiveDocument.Unit = cdrMillimeter
  Dim sh As ShapeRange
  Set sh = ActiveSelectionRange
  sx = sh.SizeWidth
  sy = sh.SizeHeight
  sx = Int(sx * 100 + 0.5) / 100
  sy = Int(sy * 100 + 0.5) / 100
End Function

'// 批量设置物件尺寸
Public Function Set_Shapes_size(ByRef sx As Double, ByRef sy As Double)
  API.BeginOpt
  ActiveDocument.ReferencePoint = cdrCenter
  
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String, size As String
  For Each sh In shs
     sh.SizeWidth = sx
     sh.SizeHeight = sy
  Next sh
  
  API.EndOpt
End Function

'// 批量设置物件尺寸整数
Public Function Size_to_Integer()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  API.BeginOpt
  '// 修改变形尺寸基准
  ActiveDocument.ReferencePoint = cdrCenter
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String, size As String
  For Each sh In shs
    size = Int(sh.SizeWidth + 0.5) & "x" & Int(sh.SizeHeight + 0.5) & "mm"
    sh.SetSize Int(sh.SizeWidth + 0.5), Int(sh.SizeHeight + 0.5)
    
    s = s & size & vbNewLine
  Next sh

  API.WriteClipBoard s
  API.EndOpt

  MsgBox "Object Size Information To Clipboard:" & vbNewLine & s & vbNewLine
End Function

'// 设置物件页面居中
Public Function Align_Page_Center()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  '// 实践应用: 选择物件群组,页面设置物件大小,物件页面居中
  API.BeginOpt
  
  Dim OrigSelection As ShapeRange, sh As Shape
  Set OrigSelection = ActiveSelectionRange
  Set sh = OrigSelection.Group
  ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
  
#If VBA7 Then
  ActiveDocument.ClearSelection
  sh.AddToSelection
  ActiveSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
#Else
  sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
#End If

  API.EndOpt
End Function


'''///  使用Python脚本 整理尺寸 提取条码数字 建立二维码 位图转文本 ///'''
Public Function Python_Organize_Size()
  On Error GoTo ErrorHandler
  mypy = Path & "GMS\LYVBA\Organize_Size.py"
  cmd_line = "pythonw " & Chr(34) & mypy & Chr(34)
  Shell cmd_line
ErrorHandler:
End Function

Public Function Python_Get_Barcode_Number()
  On Error GoTo ErrorHandler
  mypy = Path & "GMS\LYVBA\Get_Barcode_Number.py"
  cmd_line = "pythonw " & Chr(34) & mypy & Chr(34)
  Shell cmd_line
ErrorHandler:
End Function

Public Function Python_BITMAP()
  On Error GoTo ErrorHandler
  mypy = Path & "GMS\LYVBA\BITMAP.py"
  cmd_line = "pythonw " & Chr(34) & mypy & Chr(34)
  Shell cmd_line
ErrorHandler:
End Function

Public Function Python_BITMAP2()
  On Error GoTo ErrorHandler
  Bitmap = "C:\TSP\BITMAP.exe"
  Shell Bitmap
ErrorHandler:
End Function


Public Function Python_Make_QRCode()
  On Error GoTo ErrorHandler
  mypy = Path & "GMS\LYVBA\Make_QRCode.py"
  cmd_line = "pythonw " & Chr(34) & mypy & Chr(34)
  Shell cmd_line
ErrorHandler:
End Function

'// QRCode二维码制作
Public Function QRCode_replace()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim image_path As String
  image_path = API.GetClipBoardString
  ActiveDocument.ReferencePoint = cdrCenter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim X As Double, Y As Double
  Set shs = ActiveSelection.Shapes
  cnt = 0
  For Each sh In shs
    If cnt = 0 Then
      ActiveDocument.ClearSelection
      ActiveLayer.Import image_path
      Set sc = ActiveSelection
      cnt = 1
    Else
      sc.Duplicate 0, 0
    End If
    sh.GetPosition X, Y
    sc.SetPosition X, Y
    
    sh.GetSize X, Y
    sc.SetSize X, Y
    sh.Delete
    
  Next sh
  
ErrorHandler:
  API.EndOpt
End Function

'// QRCode二维码转矢量图
Public Function QRCode_to_Vector()
  On Error GoTo ErrorHandler
  
  Set sr = ActiveSelectionRange
  With sr(1).Bitmap.Trace(cdrTraceHighQualityImage)
    .TraceType = cdrTraceHighQualityImage
    .Smoothing = 50 '数值小则平滑，数值大则细节多
    .RemoveBackground = False
    .DeleteOriginalObject = True
    .Finish
  End With
 
Exit Function
ErrorHandler:
    On Error Resume Next
End Function

'''////  选择多物件，组合然后拆分线段，为角线爬虫准备  ////'''
Public Function Split_Segment()
  On Error GoTo ErrorHandler
  API.BeginOpt

  Dim ssr As ShapeRange, s As Shape
  Dim nr As NodeRange, nd As Node
  
  Set ssr = ActiveSelectionRange
  
  Set s = ssr.UngroupAllEx.Combine
  Set nr = s.Curve.Nodes.all
  
  nr.BreakApart
  s.BreakApartEx
'  For Each nd In nr
'    nd.BreakApart
'  Next nd
  
ErrorHandler:
  API.EndOpt
End Function


'''////  标记画框 支持容差  ////'''
Public Function Mark_CreateRectangle(expand As Boolean)
  On Error GoTo ErrorHandler
  API.BeginOpt
  ActiveDocument.ReferencePoint = cdrBottomLeft
  Dim ssr As ShapeRange
  Dim sh As Shape, tr As Double
  Set ssr = ActiveSelectionRange
  
  tr = 0
  If GlobalUserData.Exists("Tolerance", 1) Then
    tr = Val(GlobalUserData("Tolerance", 1))
  End If

  For Each sh In ssr
    If expand = False Then
      mark_shape sh
    Else
      mark_shape_expand sh, tr
    End If
  Next sh
  
ErrorHandler:
  API.EndOpt
End Function

Private Function mark_shape_expand(sh As Shape, tr As Double)
    Dim s As Shape
    Dim X As Double, Y As Double, w As Double, h As Double, r As Double
    sh.GetBoundingBox X, Y, w, h
    X = X - tr: Y = Y - tr:   w = w + 2 * tr: h = h + 2 * tr
    
    r = Max(w, h) / Min(w, h) / 30 * Math.Sqr(w * h)
    If w < h Then
      Set s = ActiveLayer.CreateRectangle2(X - r, Y, w + 2 * r, h)
    Else
      Set s = ActiveLayer.CreateRectangle2(X, Y - r, w, h + 2 * r)
    End If
    s.Outline.SetProperties Color:=CreateRGBColor(0, 255, 0)
End Function

Private Function mark_shape(sh As Shape)
  Dim s As Shape
  Dim X As Double, Y As Double, w As Double, h As Double
  sh.GetBoundingBox X, Y, w, h, True
  Set s = ActiveLayer.CreateRectangle2(X, Y, w, h)
  s.Outline.SetProperties Color:=CreateRGBColor(0, 255, 0)
End Function

Private Function Max(ByVal a, ByVal b)
  If a < b Then
    a = b
  End If
    Max = a
End Function

Private Function Min(ByVal a, ByVal b)
  If a > b Then
    a = b
  End If
    Min = a
End Function


'''////  批量组合合并  ////'''
Public Function Batch_Combine()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim ssr As ShapeRange, sh As Shape
  Set ssr = ActiveSelectionRange
  
  For Each sh In ssr
    sh.UngroupAllEx.Combine
  Next sh
  
ErrorHandler:
  API.EndOpt
End Function

'''////  一键拆开多行组合的文字字符   ////'''   ''' 本功能由群友半缘君赞助发行 '''
Public Function Take_Apart_Character()
  On Error GoTo ErrorHandler
  API.BeginOpt
  ActiveDocument.ReferencePoint = cdrBottomLeft
  
  Dim ssr As ShapeRange
  Dim s1 As Shape, sh As Shape, s As Shape
  Dim tr As Double
  Set ssr = ActiveSelectionRange
  
  '// 记忆选择范围
  Dim X As Double, Y As Double, w As Double, h As Double
  ssr.GetBoundingBox X, Y, w, h
  Set s1 = ActiveLayer.CreateRectangle2(X, Y, w, h)
  
  '// 解散群组，先组合，再散开
  Set s = ssr.UngroupAllEx.Combine
  Set ssr = s.BreakApartEx

  '// 读取容差值
  tr = 0
  If GlobalUserData.Exists("Tolerance", 1) Then
    tr = Val(GlobalUserData("Tolerance", 1))
  End If

  '// 标记画框，选择标记框
  For Each sh In ssr
    mark_shape_expand sh, tr
  Next sh
  
  Set ssr = ActivePage.Shapes.FindShapes(Query:="@colors.find(RGB(0, 255, 0))")
  ActiveDocument.ClearSelection
  ssr.AddToSelection
  
  '// 调用 智能群组 后删除标记画框
  SmartGroup.Smart_Group
  
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  ssr.Delete
  
  Set sh = ActivePage.SelectShapesFromRectangle(s1.LeftX, s1.TopY, s1.RightX, s1.BottomY, False)
' sh.Shapes.All.Group
  s1.Delete
  
  '// 通过s1矩形范围选择群组后合并组合
  For Each s In sh.Shapes
    s.UngroupAllEx.Combine
  Next s

ErrorHandler:
  API.EndOpt
End Function


'''//// 简单一刀切 识别群组 ////''' ''' 本功能由群友宏瑞广告赞助发行 '''
Public Function Single_Line()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim cm(2)  As Color
  Set cm(0) = CreateRGBColor(0, 255, 0) ' RGB 绿
  Set cm(1) = CreateRGBColor(255, 0, 0) ' RGB 红

  Dim ssr As ShapeRange
  Dim SrNew As New ShapeRange
  Dim s As Shape, s1 As Shape, line As Shape, line2 As Shape
  Dim cnt As Integer
  cnt = 1
  
  If 1 = ActiveSelectionRange.Count Then
    Set ssr = ActiveSelectionRange(1).UngroupAllEx
  Else
    Set ssr = ActiveSelectionRange
  End If
    
  '// 记忆选择范围
  Dim X As Double, Y As Double, w As Double, h As Double

  ssr.GetBoundingBox X, Y, w, h
  Set s1 = ActiveLayer.CreateRectangle2(X, Y, w, h)
  s1.Outline.SetProperties Color:=cm(0)
  SrNew.Add s1
  
#If VBA7 Then
'  ssr.sort " @shape1.top>@shape2.top"
  ssr.Sort " @shape1.left<@shape2.left"
#Else
' X4 不支持 ShapeRange.sort
#End If

'//  相交   Set line2 = line.Intersect(s, True, True)
'//  判断相交  line.Curve.IntersectsWith(s.Curve)

  For Each s In ssr
    If cnt > 1 Then
      s.ConvertToCurves
      Set line = ActiveLayer.CreateLineSegment(s.LeftX, s.TopY, s.LeftX, s.TopY - s.SizeHeight)
      line.Outline.SetProperties Color:=cm(1)
      SrNew.Add line
    End If
    cnt = cnt + 1
  Next s
  
  SrNew.Group
  
ErrorHandler:
  API.EndOpt
End Function

Public Function Single_Line_Vertical()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim cm(2)  As Color
  Set cm(0) = CreateRGBColor(0, 255, 0) ' RGB 绿
  Set cm(1) = CreateRGBColor(255, 0, 0) ' RGB 红

  Dim ssr As ShapeRange
  Dim SrNew As New ShapeRange
  Dim s As Shape, s1 As Shape, line As Shape, line2 As Shape
  Dim cnt As Integer
  cnt = 1
  
  If 1 = ActiveSelectionRange.Count Then
    Set ssr = ActiveSelectionRange(1).UngroupAllEx
  Else
    Set ssr = ActiveSelectionRange
  End If
    
  '// 记忆选择范围
  Dim X As Double, Y As Double, w As Double, h As Double

  ssr.GetBoundingBox X, Y, w, h
  Set s1 = ActiveLayer.CreateRectangle2(X, Y, w, h)
  s1.Outline.SetProperties Color:=cm(0)
  SrNew.Add s1
  
#If VBA7 Then
  ssr.Sort " @shape1.top>@shape2.top"
#Else
' X4 不支持 ShapeRange.sort
#End If

  For Each s In ssr
    If cnt > 1 Then
      s.ConvertToCurves
      Set line = ActiveLayer.CreateLineSegment(s.LeftX, s.TopY, s.RightX, s.TopY)
      line.Outline.SetProperties Color:=cm(1)
      SrNew.Add line
    End If
    cnt = cnt + 1
  Next s
  
  SrNew.Group
  
ErrorHandler:
  API.EndOpt
End Function

Public Function Single_Line_LastNode()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim cm(2)  As Color
  Set cm(0) = CreateRGBColor(0, 255, 0) ' RGB 绿
  Set cm(1) = CreateRGBColor(255, 0, 0) ' RGB 红

  Dim ssr As ShapeRange
  Dim SrNew As New ShapeRange
  Dim s As Shape, s1 As Shape, line As Shape, line2 As Shape
  Dim cnt As Integer
  cnt = 1
  

  If 1 = ActiveSelectionRange.Count Then
    Set ssr = ActiveSelectionRange(1).UngroupAllEx
  Else
    Set ssr = ActiveSelectionRange
  End If
    
  ' 记忆选择范围
  Dim X As Double, Y As Double, w As Double, h As Double

  ssr.GetBoundingBox X, Y, w, h
  Set s1 = ActiveLayer.CreateRectangle2(X, Y, w, h)
  s1.Outline.SetProperties Color:=cm(0)
  SrNew.Add s1
  
#If VBA7 Then
  ssr.Sort " @shape1.left<@shape2.left"
#Else
' X4 不支持 ShapeRange.sort
#End If

  Dim nr As NodeRange
  For Each s In ssr
    If cnt > 1 Then
      Set nr = s.DisplayCurve.Nodes.all
      Set line = ActiveLayer.CreateLineSegment(nr.FirstNode.PositionX, nr.FirstNode.PositionY, nr.LastNode.PositionX, nr.LastNode.PositionY)
      line.Outline.SetProperties Color:=cm(1)
      SrNew.Add line
    End If
    cnt = cnt + 1
  Next s
  
  SrNew.Group
  
ErrorHandler:
  API.EndOpt
End Function


'''//// 选择范围画框 ////'''
Public Function Mark_Range_Box()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  ActiveDocument.Unit = cdrMillimeter
  Dim s1 As Shape, ssr As ShapeRange
  
  Set ssr = ActiveSelectionRange
  Dim X As Double, Y As Double, w As Double, h As Double

  ssr.GetBoundingBox X, Y, w, h
  Set s1 = ActiveLayer.CreateRectangle2(X, Y, w, h)
  s1.Outline.SetProperties Color:=CreateRGBColor(0, 255, 0)  '// RGB 绿
End Function


'''//// 快速颜色选择 ////'''
Function quickColorSelect()
    Dim X As Double, Y As Double
    Dim s As Shape, s1 As Shape
    Dim sr As ShapeRange, sr2 As ShapeRange
    Dim Shift As Long, bClick As Boolean
    Dim c As New Color, c2 As New Color

    EventsEnabled = False
    
    Set sr = ActivePage.Shapes.FindShapes(Query:="@fill.type = 'uniform'")
    ActiveDocument.ClearSelection
    bClick = False
    While Not bClick
    On Error Resume Next
        bClick = ActiveDocument.GetUserClick(X, Y, Shift, 10, False, cdrCursorPickNone)
        If Not bClick Then
            Set s = ActivePage.SelectShapesAtPoint(X, Y, False)
            Set s = s.Shapes.Last
            c2.CopyAssign s.Fill.UniformColor
            Set sr2 = New ShapeRange
            For Each s1 In sr.Shapes
                c.CopyAssign s1.Fill.UniformColor
                If c.IsSame(c2) Then
                    sr2.Add s1
                End If
            Next s1
            sr2.CreateSelection
            ActiveWindow.Refresh
        End If
    Wend
    
    EventsEnabled = True
End Function


'''//// 切割图形-垂直分割-水平分割 ////'''
Function divideVertically()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  
  cutInHalf 1
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
  
Exit Function
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function

Function divideHorizontally()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  
  cutInHalf 2
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
  
Exit Function
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function

Private Function cutInHalf(Optional method As Integer)
    Dim s As Shape, rect As Shape, rect2 As Shape
    Dim trimmed1 As Shape, trimmed2 As Shape
    Dim X As Double, Y As Double, w As Double, h As Double
    Dim vBool As Boolean
    Dim leeway As Double
    Dim sr As ShapeRange, sr2 As New ShapeRange
    
    vBool = True
    If method = 2 Then
        vBool = False
    End If
    leeway = 0.1
    Set sr = ActiveSelectionRange
    ActiveDocument.BeginCommandGroup "Cut in half"
    For Each s In sr
        s.GetBoundingBox X, Y, w, h
        
        If (vBool) Then
            'vertical slice
            Set rect = ActiveLayer.CreateRectangle2(X - leeway, Y - leeway, (w / 2) + leeway, h + (leeway * 2))
            Set rect2 = ActiveLayer.CreateRectangle2(X + (w / 2), Y - leeway, (w / 2) + leeway, h + (leeway * 2))
        Else
            Set rect = ActiveLayer.CreateRectangle2(X - leeway, Y - leeway, w + (leeway * 2), (h / 2) + leeway)
            Set rect2 = ActiveLayer.CreateRectangle2(X - leeway, Y + (h / 2), w + (leeway * 2), (h / 2) + leeway)
        End If
        
        Set trimmed1 = rect.Intersect(s, True, True)
        rect.Delete
        Set trimmed2 = rect2.Intersect(s, True, True)
        s.Delete
        rect2.Delete
        sr2.Add trimmed1
        sr2.Add trimmed2
    Next s
    ActiveDocument.EndCommandGroup
    
    sr2.CreateSelection
End Function


'// 批量多页居中-遍历批量物件，放置物件到页面
Public Function Batch_Align_Page_Center()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Set sr = ActiveSelectionRange
  total = sr.Count

  '// 建立多页面
  Set doc = ActiveDocument
  doc.AddPages (total - 1)


#If VBA7 Then
  sr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
#Else
' X4 不支持 ShapeRange.sort
#End If


  Dim sh As Shape

  '// 遍历批量物件，放置物件到页面
  For i = 1 To sr.Count
    doc.Pages(i).Activate
    Set sh = sr.Shapes(i)
    ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
 
   '// 物件居中页面
#If VBA7 Then
  ActiveDocument.ClearSelection
  sh.AddToSelection
  ActiveSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
#Else
  sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
#End If

  Next i
ErrorHandler:
  API.EndOpt
End Function


'// 安全线: 点击一次建立辅助线，再调用清除参考线
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

'// 标注尺寸 批量简单标注数字
Public Function Simple_Label_Numbers()
  API.BeginOpt
  Set sr = ActiveSelectionRange
  
  For Each s In sr.Shapes
    X = s.CenterX: Y = s.TopY
    sw = s.SizeWidth: sh = s.SizeHeight
          
    text = Int(sw + 0.5) & "x" & Int(sh + 0.5) & "mm"
    Set s = ActiveLayer.CreateArtisticText(0, 0, text)
    s.CenterX = X: s.BottomY = Y + 5
  Next
  API.EndOpt
End Function

'// 修复圆角缺角到直角
Public Function corner_off()
  API.BeginOpt
    Dim os As ShapeRange
    Dim s As Shape, fir As Shape, ci As Shape
    Dim nd As Node, nds As Node, nde As Node
    
    Set os = ActiveSelectionRange

On Error GoTo errn
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
  API.EndOpt
End Function

Private Function corner_off_make(s As Shape, nds As Node, nde As Node)
    Dim l1 As Shape, l2 As Shape
    Dim os As ShapeRange
    Dim ss As Shape

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
End Function

Public Function autogroup(Optional Group As String = "group", Optional shft = 0, Optional sss As Shapes = Nothing, Optional undogroup = True) As ShapeRange
  Dim sr As ShapeRange, sr_all As ShapeRange, os As ShapeRange
  Dim sp As SubPaths
  Dim arr()
  Dim s As Shape
  If sss Is Nothing Then Set os = ActiveSelectionRange Else Set os = sss.all
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

'// 两个端点的坐标,为(x1,y1)和(x2,y2) 那么其角度a的tan值: tana=(y2-y1)/(x2-x1)
'// 所以计算arctan(y2-y1)/(x2-x1), 得到其角度值a
'// VB中用atn(), 返回值是弧度，需要 乘以 PI /180
Private Function lineangle(x1, y1, x2, y2) As Double
  pi = 4 * VBA.Atn(1)    '// 计算圆周率
  If x2 = x1 Then
    lineangle = 90: Exit Function
  End If
  lineangle = VBA.Atn((y2 - y1) / (x2 - x1)) / pi * 180
End Function

'// 角度转平
Public Function Angle_to_Horizon()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.all

  If nr.Count = 2 Then
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2): sr.Rotate -a
    sr.LastShape.Delete   '// 删除参考线
  End If
ErrorHandler:
  API.EndOpt
End Function

'// 自动旋转角度
Public Function Auto_Rotation_Angle()
  On Error GoTo ErrorHandler
  API.BeginOpt
  
'  ActiveDocument.ReferencePoint = cdrCenter
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.all

  If nr.Count = 2 Then
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2): sr.Rotate 90 + a
    sr.LastShape.Delete   '// 删除参考线
  End If
ErrorHandler:
  API.EndOpt
End Function

'// 交换对象
Public Function Exchange_Object()
  Set sr = ActiveSelectionRange
  If sr.Count = 2 Then
    X = sr.LastShape.CenterX: Y = sr.LastShape.CenterY
    sr.LastShape.CenterX = sr.FirstShape.CenterX: sr.LastShape.CenterY = sr.FirstShape.CenterY
    sr.FirstShape.CenterX = X: sr.FirstShape.CenterY = Y
  End If
End Function

'// 参考线镜像
Public Function Mirror_ByGuide()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.all

  If nr.Count = 2 Then
    byshape = False
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2)  '// 参考线和水平的夹角 a
    sr.Remove sr.Count
    
    ang = 90 - a    '// 镜像的旋转角度
    For Each s In sr
      With s
        .Duplicate   '// 复制物件保留，然后按 x1,y1 点 旋转
        .RotationCenterX = x1
        .RotationCenterY = y1
        .Rotate ang
        If Not byshape Then
            lx = .LeftX
            .Stretch -1#, 1#    '// 通过拉伸完成镜像
            .LeftX = lx
            .Move (x1 - .LeftX) * 2 - .SizeWidth, 0
            .RotationCenterX = x1     '// 之前因为镜像，旋转中心点反了，重置回来
            .RotationCenterY = y1
            .Rotate -ang
        End If
        .RotationCenterX = .CenterX   '// 重置回旋转中心点为物件中心
        .RotationCenterY = .CenterY
      End With
    Next s

  End If

ErrorHandler:
  API.EndOpt
End Function

'// 按面积排列计数
Public Function Count_byArea(Space_Width As Double)
  If 0 = ActiveSelectionRange.Count Then Exit Function
  API.BeginOpt
  ActiveDocument.ReferencePoint = cdrCenter
  
  Set ssr = ActiveSelectionRange
  cnt = 1

#If VBA7 Then
  ssr.Sort "@shape1.width * @shape1.height < @shape2.width * @shape2.height"
#Else
' X4 不支持 ShapeRange.sort
#End If

  Dim Str As String, size As String
  For Each sh In ssr
    size = Int(sh.SizeWidth + 0.5) & "x" & Int(sh.SizeHeight + 0.5) & "mm"
    sh.SetSize Int(sh.SizeWidth + 0.5), Int(sh.SizeHeight + 0.5)
    Str = Str & size & vbNewLine
  Next sh

  ActiveDocument.ReferencePoint = cdrTopLeft
  For Each s In ssr
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).LeftX, ssr(cnt - 1).BottomY - Space_Width
    cnt = cnt + 1
  Next s

'  写文件，可以EXCEL里统计
'  Set fs = CreateObject("Scripting.FileSystemObject")
'  Set f = fs.CreateTextFile("D:\size.txt", True)
'  f.WriteLine str: f.Close

  Str = Subtotals(Str)
  Debug.Print Str

  Dim s1 As Shape
' Set s1 = ActiveLayer.CreateParagraphText(0, 0, 100, 150, Str, Font:="华文中宋")
  X = ssr.FirstShape.LeftX - 100
  Y = ssr.FirstShape.TopY
  Set s1 = ActiveLayer.CreateParagraphText(X, Y, X + 90, Y - 150, Str, Font:="华文中宋")

  API.EndOpt
End Function
 
'// 实现Excel里分类汇总功能
Private Function Subtotals(Str As String) As String
  Dim a, b, d, arr
  Str = VBA.Replace(Str, vbNewLine, " ")
  Do While InStr(Str, "  ")
      Str = VBA.Replace(Str, "  ", " ")
  Loop
  arr = Split(Str)

  Set d = CreateObject("Scripting.dictionary")

  For i = 0 To UBound(arr) - 1
    If d.Exists(arr(i)) = True Then
      d.Item(arr(i)) = d.Item(arr(i)) + 1
    Else
       d.Add arr(i), 1
    End If
  Next

  Str = "   规   格" & vbTab & vbTab & vbTab & "数量" & vbNewLine

  a = d.keys: b = d.items
  For i = 0 To d.Count - 1
    ' Debug.Print a(i), b(i)
    Str = Str & a(i) & vbTab & vbTab & b(i) & "条" & vbNewLine
  Next

  Subtotals = Str & "合计总量:" & vbTab & vbTab & vbTab & UBound(arr) & "条" & vbNewLine
End Function
