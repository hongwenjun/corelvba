Attribute VB_Name = "Tools"
Public Function 分分合合()
  拼版裁切线.arrange
  
  CQL查找相同.CQLline_CM100
  
  拼版裁切线.Cut_lines

  Dim s As Shape
  Set s = ActivePage.SelectShapesFromRectangle(ActivePage.LeftX, ActivePage.TopY, ActivePage.RightX, ActivePage.BottomY, True)
  
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

  MsgBox "物件尺寸信息到剪贴板" & vbNewLine & s
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

'' 选择多物件，组合然后拆分线段，为角线爬虫准备
Public Function Split_Segment()
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  
  Dim ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  Dim s As Shape
  Dim nr As NodeRange
  Dim nd As Node
  
  Set s = ssr.Combine
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
