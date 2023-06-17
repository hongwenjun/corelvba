Attribute VB_Name = "TSP"
'// 导出节点信息到数据文件
Public Function CDR_TO_TSP()
  API.BeginOpt
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.CreateTextFile("C:\TSP\CDR_TO_TSP", True)

  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim X As Double, Y As Double
  Set shs = ActiveSelection.Shapes
  
  Dim TSP As String
  TSP = shs.Count & " " & 0 & vbNewLine
  For Each sh In shs
    X = sh.CenterX
    Y = sh.CenterY
    TSP = TSP & X & " " & Y & vbNewLine
  Next sh
  
  f.WriteLine TSP
  f.Close
'//  MsgBox "小圆点导出节点信息到数据文件!" & vbNewLine
  API.EndOpt
End Function

'// 导出节点信息到数据文件
Public Function Nodes_To_TSP()
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.CreateTextFile("C:\TSP\CDR_TO_TSP", True)
  ActiveDocument.Unit = cdrMillimeter
  
  Dim ssr As ShapeRange
  Set ssr = ActiveSelectionRange.Duplicate
  Dim s As Shape
  Dim nr As NodeRange
  Dim nd As Node
  
  Dim X As String, Y As String
  Dim TSP As String
  
  Set s = ssr.UngroupAllEx.Combine
  Set nr = s.Curve.Nodes.all
  
  TSP = nr.Count & " " & 0 & vbNewLine
  For Each n In nr
      X = round(n.PositionX, 3) & " "
      Y = round(n.PositionY, 3) & vbNewLine
      TSP = TSP & X & Y
  Next n
  
  f.WriteLine TSP
  f.Close
  s.Delete
'//   MsgBox "选择物件导出节点信息到数据文件!" & vbNewLine
  
ErrorHandler:
  API.EndOpt
End Function

'// 运行CDR2TSP.exe
Public Function START_TSP()
  On Error GoTo ErrorHandler
  cmd_line = "C:\TSP\CDR2TSP.exe C:\TSP\CDR_TO_TSP"
  Shell cmd_line
  
ErrorHandler:
End Function

'//  TSP功能画线-连贯线
Public Function TSP_TO_DRAW_LINE()
  On Error GoTo ErrorHandler
  API.BeginOpt

  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.OpenTextFile("C:\TSP\TSP.txt", 1, False)
  Dim Str, arr, n
  Str = f.ReadAll()
  
  Str = API.Newline_to_Space(Str)
  arr = Split(Str)
  total = Val(arr(0))
  
  ReDim ce(total) As CurveElement
  Dim crv As Curve
  
  ce(0).ElementType = cdrElementStart
  ce(0).PositionX = Val(arr(2)) - 3    '// 线条起始坐标，偏移3mm方向指示
  ce(0).PositionY = Val(arr(3)) - 3
  
  Dim X As Double
  Dim Y As Double
  For n = 2 To UBound(arr) - 1 Step 2
    X = Val(arr(n))
    Y = Val(arr(n + 1))
  
    ce(n / 2).ElementType = cdrElementLine
    ce(n / 2).PositionX = X
    ce(n / 2).PositionY = Y
  
  Next
  
  Set crv = CreateCurve(ActiveDocument)
  crv.CreateSubPathFromArray ce
  ActiveLayer.CreateCurve crv
  
ErrorHandler:
  API.EndOpt
End Function

'// 设置线条标记(颜色)
Private Function set_line_color(line As Shape)
  line.Outline.SetProperties Color:=CreateRGBColor(26, 22, 35)
End Function


'//  TSP功能画线-多线段
Public Function TSP_TO_DRAW_LINES()
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.OpenTextFile("C:\TSP\TSP2.txt", 1, False)
  Dim Str, arr, n
  Dim line As Shape
  Str = f.ReadAll()
  
  Str = API.Newline_to_Space(Str)
  
  arr = Split(Str)
  For n = 2 To UBound(arr) - 1 Step 4
    X = Val(arr(n))
    Y = Val(arr(n + 1))
    x1 = Val(arr(n + 2))
    y1 = Val(arr(n + 3))

    Set line = ActiveLayer.CreateLineSegment(X, Y, x1, y1)
    set_line_color line
  Next
  
  ActivePage.Shapes.FindShapes(Query:="@colors.find(RGB(26, 22, 35))").CreateSelection
  ActiveSelection.Group
  ActiveSelection.Outline.SetProperties 0.2, Color:=CreateCMYKColor(0, 100, 100, 0)
  
ErrorHandler:
  API.EndOpt
End Function

'// 运行 TSP.exe
Public Function MAKE_TSP()
  On Error GoTo ErrorHandler
  cmd_line = "C:\TSP\TSP.exe"
  Shell cmd_line
ErrorHandler:
End Function

'// 位图制作小圆点
Public Function BITMAP_MAKE_DOTS()
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim line, art, n, h, w
  Dim X As Double
  Dim Y As Double
  Dim s As Shape
  flag = 0
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.OpenTextFile("C:\TSP\BITMAP", 1, False)

  line = f.ReadLine()
  Debug.Print line

  ' 读取第一行，位图 h高度 和 w宽度
  arr = Split(line)
  h = Val(arr(0)): w = Val(arr(1))
  
  If h * w > 20000 Then
'//      MsgBox "位图转换后的小圆点数量比较多:" & vbNewLine & h & " x " & w & " = " & h * w
      flag = 1
  End If

  For i = 1 To h
    line = f.ReadLine()
    arr = Split(line)
    For n = LBound(arr) To UBound(arr)
      If arr(n) > 0 Then
        X = n: Y = -i
        If flag = 1 Then
          Set s = ActiveLayer.CreateRectangle2(X, Y, 0.6, 0.6)
        Else
          make_dots X, Y
        End If
      End If
    Next n
  Next i

ErrorHandler:
  API.EndOpt
End Function

'// 坐标绘制圆点
Private Function make_dots(X As Double, Y As Double)
  Dim s As Shape, c As Variant
  c = Array(0, 255, 0)
  Set s = ActiveLayer.CreateEllipse2(X, Y, 0.5, 0.5)
  s.Fill.UniformColor.RGBAssign c(Int(Rnd() * 2)), c(Int(Rnd() * 2)), c(Int(Rnd() * 2))
  s.Outline.Width = 0#
End Function
