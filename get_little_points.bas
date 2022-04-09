Sub 一键加点工具()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.Copy
    
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
    Set sh = ActiveSelection.Group
    sh.Outline.SetProperties 0.03, Color:=CreateCMYKColor(0, 100, 100, 0)
    sr.ApplyUniformFill CreateCMYKColor(0, 100, 100, 0)
    sh.Move 0, 0.015
  Else
    MsgBox "文件中小圆点足够大，不需要加粗!"
  End If

  ActivePage.Shapes.FindShapes(Query:="@colors.find(CMYK(50, 0, 0, 0))").CreateSelection
  ActiveSelection.Group
  
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
