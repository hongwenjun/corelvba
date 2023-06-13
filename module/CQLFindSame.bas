Attribute VB_Name = "CQLFindSame"
Sub 属性选择()
  CQL_FIND_UI.Show 0
End Sub

Public Function CQLline_CM100()
  On Error GoTo err
  Dim cm(5) As Color, i As Long
  Set cm(0) = CreateCMYKColor(100, 0, 100, 0)  '绿
  Set cm(1) = CreateCMYKColor(0, 100, 0, 0)  '洋红
  Set cm(2) = CreateCMYKColor(100, 100, 0, 0) '红
  Set cm(3) = CreateRGBColor(0, 255, 0) ' RGB 绿
  Set cm(4) = CreateRGBColor(255, 0, 0) ' RGB 红

  ActiveDocument.ClearSelection
  For i = 0 To 4
    cm(i).ConvertToRGB
    r = cm(i).RGBRed
    G = cm(i).RGBGreen
    b = cm(i).RGBBlue
    ActivePage.Shapes.FindShapes(Query:="@Outline.Color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").AddToSelection
  Next i

Exit Function
err:
  MsgBox "Function CQLline_CM100 错误!"
End Function


Sub 一键加点工具()
  Dim OrigSelection As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  If OrigSelection.Count <> 0 Then
    OrigSelection.Copy
  Else
    MsgBox "选择水洗标要加点部分，然后点击【加点工具】按钮!"
    Exit Sub
  End If
  
  ' 新建文件粘贴
  Dim doc1 As Document
  Set doc1 = CreateDocument
  ActiveLayer.Paste
  
  ' 转曲线，一键加粗小红点
  ActiveSelection.ConvertToCurves
  Call get_little_points
End Sub


Private Sub get_little_points()
  On Error GoTo ErrorHandler
  '// 代码运行时关闭窗口刷新
  Application.Optimization = True
  ActiveDocument.BeginCommandGroup  '一步撤消'
  
  red_point_Size = 0.3
  ActiveDocument.Unit = cdrMillimeter
  Dim OrigSelection As ShapeRange, grp1 As ShapeRange, sh As Shape
  Set OrigSelection = ActiveSelectionRange
  Set grp1 = OrigSelection.UngroupAllEx
  grp1.ApplyUniformFill CreateCMYKColor(50, 0, 0, 0)
  
  For Each sh In grp1
    sh.BreakApartEx
  Next sh
  
  ActiveDocument.ClearSelection
  Dim sr As ShapeRange
  Set sr = ActivePage.Shapes.FindShapes(Query:="@width < {" & red_point_Size & " mm} and @width > {0.1 mm} and @height <{" & red_point_Size & " mm} and @height >{0.1 mm}")
  If sr.Count <> 0 Then
    sr.CreateSelection
    Set sh = ActiveSelection.group
    sh.Outline.SetProperties 0.03, Color:=CreateCMYKColor(0, 100, 100, 0)
    sr.ApplyUniformFill CreateCMYKColor(0, 100, 100, 0)
    sh.Move 0, 0.015
    sh.Copy
  Else
    MsgBox "文件中小圆点足够大，不需要加粗!"
  End If

  ActivePage.Shapes.FindShapes(Query:="@colors.find(CMYK(50, 0, 0, 0))").CreateSelection
  ActiveSelection.group
  
  ActiveDocument.EndCommandGroup
  Application.Optimization = False
  ActiveWindow.Refresh
  Application.Refresh
  Exit Sub
ErrorHandler:
  MsgBox "选择水洗标要加点部分，然后点击【加点工具】按钮!"
  Application.Optimization = False
  On Error Resume Next
End Sub

Sub 文字转曲()
  Tools.TextShape_ConvertToCurves
End Sub

