Attribute VB_Name = "Tools"
Public Function 分分合合()
  拼版裁切线.arrange
  
  CQL查找相同.CQLline_CM100
  
  拼版裁切线.Cut_lines

  ' 记忆选择范围
  Dim x As Double, y As Double, w As Double, h As Double
  ActiveSelectionRange.GetBoundingBox x, y, w, h
  Set s = ActivePage.SelectShapesFromRectangle(x, y, w, h, True)
  
  自动中线色阶条.Auto_ColorMark

End Function


Public Function 傻瓜火车排列()
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
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
    '' 底对齐 If cnt > 1 Then s.SetPosition ssr(cnt - 1).RightX, ssr(cnt - 1).BottomY
    '' 改成顶对齐 2022-08-10
    ActiveDocument.ReferencePoint = cdrTopLeft + cdrBottomTop
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).RightX, ssr(cnt - 1).TopY
    cnt = cnt + 1
  Next s

  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
End Function


Public Function 傻瓜阶梯排列()
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
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
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).LeftX, ssr(cnt - 1).BottomY
    cnt = cnt + 1
  Next s

  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
End Function

'// 文本转曲线
Public Function TextShape_ConvertToCurves()
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  Dim s As Shape, cnt As Long
  For Each s In API.FindAllShapes.Shapes.FindShapes(, cdrTextShape)
    s.ConvertToCurves
    cnt = cnt + 1
  Next s
  MsgBox "转曲物件统计: " & cnt, , "文本转曲线"
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
End Function

'' 复制物件
Public Function copy_shape()
  Dim OrigSelection As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  OrigSelection.Copy

End Function

'' 旋转物件角度
Public Function Rotate_Shapes(n As Double)
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  ActiveDocument.Unit = cdrMillimeter
  
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String, size As String
  For Each sh In shs
    sh.Rotate n
  Next sh
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
End Function

'' 得到物件尺寸
Public Function get_shape_size(ByRef sx As Double, ByRef sy As Double)
  ActiveDocument.Unit = cdrMillimeter
  Dim sh As ShapeRange
  Set sh = ActiveSelectionRange
  sx = sh.SizeWidth
  sy = sh.SizeHeight
  sx = Int(sx * 100 + 0.5) / 100
  sy = Int(sy * 100 + 0.5) / 100
End Function

'' 批量设置物件尺寸
Public Function Set_Shapes_size(ByRef sx As Double, ByRef sy As Double)
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.ReferencePoint = cdrCenter
  
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String, size As String
  For Each sh In shs
     sh.SizeWidth = sx
     sh.SizeHeight = sy
  Next sh
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
End Function

Public Function 尺寸取整()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  ActiveDocument.Unit = cdrMillimeter
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String, size As String
  For Each sh In shs
    size = Int(sh.SizeWidth + 0.5) & "x" & Int(sh.SizeHeight + 0.5) & "mm"
    sh.SetSize Int(sh.SizeWidth + 0.5), Int(sh.SizeHeight + 0.5)
    
    s = s & size & vbNewLine
  Next sh

  MsgBox "物件尺寸信息到剪贴板" & vbNewLine & s & vbNewLine
  API.WriteClipBoard s

End Function

Public Function 居中页面()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  ' 实践应用: 选择物件群组,页面设置物件大小,物件页面居中
  ActiveDocument.Unit = cdrMillimeter
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
End Function


Public Function Python脚本整理尺寸()
    mypy = Path & "GMS\262235.xyz\整理尺寸.py"
    cmd_line = "pythonw " & Chr(34) & mypy & Chr(34)
    Shell cmd_line
End Function

Public Function Python提取条码数字()
    mypy = Path & "GMS\262235.xyz\提取条码数字.py"
    cmd_line = "pythonw " & Chr(34) & mypy & Chr(34)
    Shell cmd_line
End Function

Public Function Python二维码QRCode()
    mypy = Path & "GMS\262235.xyz\二维码QRCode.py"
    cmd_line = "pythonw " & Chr(34) & mypy & Chr(34)
    Shell cmd_line
End Function

'' QRCode二维码制作
Public Function QRCode_replace()
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  Dim image_path As String
  image_path = API.GetClipBoardString
  ActiveDocument.ReferencePoint = cdrCenter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim x As Double, y As Double
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
    sh.GetPosition x, y
    sc.SetPosition x, y
    
    sh.GetSize x, y
    sc.SetSize x, y
    sh.Delete
    
  Next sh
  
    '// 代码操作结束恢复窗口刷新
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh:    Application.Refresh
Exit Function
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function

'' QRCode二维码转矢量图
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


'''////  标记画框 支持容差  ////'''
Public Function Mark_CreateRectangle(expand As Boolean)
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.ReferencePoint = cdrBottomLeft
  Dim ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  Dim sh As Shape
  Dim tr As Double
  
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
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
Exit Function
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function

Private Function mark_shape_expand(sh As Shape, tr As Double)
    Dim s As Shape
    Dim x As Double, y As Double, w As Double, h As Double, r As Double
    sh.GetBoundingBox x, y, w, h
    x = x - tr: y = y - tr:   w = w + 2 * tr: h = h + 2 * tr
    
    r = Max(w, h) / Min(w, h) / 30 * Math.Sqr(w * h)
    If w < h Then
      Set s = ActiveLayer.CreateRectangle2(x - r, y, w + 2 * r, h)
    Else
      Set s = ActiveLayer.CreateRectangle2(x, y - r, w, h + 2 * r)
    End If
    s.Outline.SetProperties Color:=CreateRGBColor(0, 255, 0)
End Function

Public Function Create_Tolerance()
  Dim text As String
  If GlobalUserData.Exists("Tolerance", 1) Then
    text = GlobalUserData("Tolerance", 1)
  End If
  text = InputBox("请输入容差值 0 --> 99", "容差值(mm)", text)
  If text = "" Then Exit Function
  GlobalUserData("Tolerance", 1) = text
End Function

Private Function mark_shape(sh As Shape)
  Dim s As Shape
  Dim x As Double, y As Double, w As Double, h As Double
  sh.GetBoundingBox x, y, w, h
  Set s = ActiveLayer.CreateRectangle2(x, y, w, h)
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
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  
  Dim ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  Dim sh As Shape
  For Each sh In ssr
    sh.UngroupAllEx.Combine
  Next sh
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
  
Exit Function
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function

'''////  一键拆开多行组合的文字字符   ////'''   ''' 本功能由群友半缘君赞助发行 '''
Public Function Take_Apart_Character()
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.ReferencePoint = cdrBottomLeft
  
  Dim ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  Dim s1 As Shape, sh As Shape, s As Shape
  Dim tr As Double
  
  ' 记忆选择范围
  Dim x As Double, y As Double, w As Double, h As Double
  ssr.GetBoundingBox x, y, w, h
  Set s1 = ActiveLayer.CreateRectangle2(x, y, w, h)
  
  ' 解散群组，先组合，再散开
  Set s = ssr.UngroupAllEx.Combine
  Set ssr = s.BreakApartEx

  ' 读取容差值
  tr = 0
  If GlobalUserData.Exists("Tolerance", 1) Then
    tr = Val(GlobalUserData("Tolerance", 1))
  End If

  ' 标记画框，选择标记框
  For Each sh In ssr
    mark_shape_expand sh, tr
  Next sh
  
  Set ssr = ActivePage.Shapes.FindShapes(Query:="@colors.find(RGB(0, 255, 0))")
  ActiveDocument.ClearSelection
  ssr.AddToSelection
  
  ' 调用 智能群组 后删除标记画框
  智能群组和查找.智能群组
  
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  ssr.Delete
  
  Set sh = ActivePage.SelectShapesFromRectangle(s1.LeftX, s1.TopY, s1.RightX, s1.BottomY, False)
' sh.Shapes.All.Group
  s1.Delete
  
  ' 通过s1矩形范围选择群组后合并组合
  For Each s In sh.Shapes
    s.UngroupAllEx.Combine
  Next s

  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
  
Exit Function
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function


'''////  简单一刀切  识别群组 ////'''   ''' 本功能由群友宏瑞广告赞助发行 '''
Public Function Single_Line()
If 0 = ActiveSelectionRange.Count Then Exit Function
'  On Error GoTo ErrorHandler
'  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  ActiveDocument.Unit = cdrMillimeter
  
  Dim cm(2)  As Color
  Set cm(0) = CreateRGBColor(0, 255, 0) ' RGB 绿
  Set cm(1) = CreateRGBColor(255, 0, 0) ' RGB 红

  Dim ssr As ShapeRange
  Dim SrNew As New ShapeRange
  Dim s As Shape, s1 As Shape, line As Shape
  Dim cnt As Integer
  cnt = 1
  
  
  If 1 = ActiveSelectionRange.Count Then
    Set ssr = ActiveSelectionRange(1).UngroupAllEx
  Else
    Set ssr = ActiveSelectionRange
  End If
    
  ' 记忆选择范围
  Dim x As Double, y As Double, w As Double, h As Double

  ssr.GetBoundingBox x, y, w, h
  Set s1 = ActiveLayer.CreateRectangle2(x, y, w, h)
  s1.Outline.SetProperties Color:=cm(0)
  SrNew.Add s1
  
#If VBA7 Then
'  ssr.sort " @shape1.top>@shape2.top"
  ssr.Sort " @shape1.left<@shape2.left"
#Else
' X4 不支持 ShapeRange.sort
#End If

  For Each s In ssr
    If cnt > 1 Then
      Set line = ActiveLayer.CreateLineSegment(s.LeftX, s.TopY, s.LeftX, s.TopY - h)
      line.Outline.SetProperties Color:=cm(1)
      SrNew.Add line
    End If
    cnt = cnt + 1
  Next s
  
  SrNew.Group
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh:    Application.Refresh
  
Exit Function
ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function
