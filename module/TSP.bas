Attribute VB_Name = "TSP"
Public Function CDR_TO_TSP()
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.CreateTextFile("C:\TSP\CDR_TO_TSP", True)
  
  ActiveDocument.Unit = cdrMillimeter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim x As Double, y As Double
  Set shs = ActiveSelection.Shapes
  
  Dim TSP As String
  TSP = shs.Count & " " & 0 & vbNewLine
  For Each sh In shs
    x = sh.CenterX
    y = sh.CenterY
    TSP = TSP & x & " " & y & vbNewLine
  Next sh
  
  f.WriteLine TSP
  f.Close
  MsgBox "小圆点导出节点信息到数据文件!" & vbNewLine
End Function


Public Function PATH_TO_TSP()
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.CreateTextFile("C:\TSP\CDR_TO_TSP", True)
  
  ActiveDocument.Unit = cdrMillimeter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim x As Double, y As Double
  Set shs = ActiveSelection.Shapes
  
  Dim TSP As String
  TSP = shs.Count & " " & 0 & vbNewLine
  For Each sh In shs
    x = sh.CenterX
    y = sh.CenterY
    TSP = TSP & x & " " & y & vbNewLine
  Next sh
  
  f.WriteLine TSP
  f.Close
  MsgBox "选择曲线导出节点信息到数据文件!" & vbNewLine
End Function


Public Function START_TSP()
    cmd_line = "C:\TSP\CDR2TSP.exe C:\TSP\CDR_TO_TSP"
    Shell cmd_line
End Function

Public Function TSP_TO_DRAW_LINE()
 ' On Error GoTo ErrorHandler
  ActiveDocument.Unit = cdrMillimeter
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.OpenTextFile("C:\TSP\TSP.txt", 1, False)
  Dim Str, arr, n
  Str = f.ReadAll()
  
  Str = VBA.replace(Str, vbNewLine, " ")
  Do While InStr(Str, "  ")
      Str = VBA.replace(Str, "  ", " ")
  Loop
  
  arr = Split(Str)
  total = Val(arr(0))
  
  ReDim ce(total) As CurveElement
  Dim crv As Curve
  
  ce(0).ElementType = cdrElementStart
  ce(0).PositionX = 0
  ce(0).PositionY = 0
  
  Dim x As Double
  Dim y As Double
  For n = 2 To UBound(arr) - 1 Step 2
    x = Val(arr(n))
    y = Val(arr(n + 1))
  
    ce(n / 2).ElementType = cdrElementLine
    ce(n / 2).PositionX = x
    ce(n / 2).PositionY = y
  
  Next
  
  Set crv = CreateCurve(ActiveDocument)
  crv.CreateSubPathFromArray ce
  ActiveLayer.CreateCurve crv
  
ErrorHandler:
  On Error Resume Next
End Function

Public Function TSP_TO_DRAW_LINE_BAK()
  On Error GoTo ErrorHandler
  ActiveDocument.Unit = cdrMillimeter
  
  Dim Str, arr, n
  Str = API.GetClipBoardString
  Str = VBA.replace(Str, vbNewLine, " ")
  Do While InStr(Str, "  ")
      Str = VBA.replace(Str, "  ", " ")
  Loop
  
  arr = Split(Str)
  total = Val(arr(0))
  
  ReDim ce(total) As CurveElement
  Dim crv As Curve
  
  ce(0).ElementType = cdrElementStart
  ce(0).PositionX = 0
  ce(0).PositionY = 0
  
  Dim x As Double
  Dim y As Double
  For n = 2 To UBound(arr) - 1 Step 2
    x = Val(arr(n))
    y = Val(arr(n + 1))
  
    ce(n / 2).ElementType = cdrElementLine
    ce(n / 2).PositionX = x
    ce(n / 2).PositionY = y
  
  Next
  
  Set crv = CreateCurve(ActiveDocument)
  crv.CreateSubPathFromArray ce
  ActiveLayer.CreateCurve crv
  
ErrorHandler:
  On Error Resume Next
End Function


Public Function MAKE_TSP()
    cmd_line = "C:\TSP\TSP.exe"
    Shell cmd_line
End Function

' 位图制作小圆点
Public Function BITMAP_MAKE_DOTS()
 ' On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup: Application.Optimization = True
  ActiveDocument.Unit = cdrMillimeter
  Dim line, art, n, h, w
  Dim x As Double
  Dim y As Double
  Dim s As Shape
  flag = 0
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.OpenTextFile("C:\TSP\BITMAP", 1, False)

  line = f.ReadLine()
  Debug.Print line

  ' 读取第一行，位图 h高度 和 w宽度
  arr = Split(line)
  h = Val(arr(0)): w = Val(arr(1))
  
  If h * w > 40000 Then
      MsgBox "位图转换后的小圆点数量比较多:" & vbNewLine & h & " x " & w & " = " & h * w
      flag = 1
  End If

  For i = 1 To h
    line = f.ReadLine()
    arr = Split(line)
    For n = LBound(arr) To UBound(arr)
      If arr(n) > 0 Then
        x = n: y = -i
        If flag = 1 Then
          Set s = ActiveLayer.CreateRectangle2(x, y, 0.6, 0.6)
        Else
          make_dots x, y
        End If
      End If
    Next n
  Next i

  ActiveDocument.EndCommandGroup: Application.Optimization = False
  ActiveWindow.Refresh: Application.Refresh
Exit Function
ErrorHandler:
    Application.Optimization = False
    On Error Resume Next
End Function

Private Function make_dots(x As Double, y As Double)
  Dim s As Shape
  Dim c As Variant
  c = Array(0, 255, 0)
  Set s = ActiveLayer.CreateEllipse2(x, y, 0.5, 0.5)
  s.Fill.UniformColor.RGBAssign c(Int(Rnd() * 2)), c(Int(Rnd() * 2)), c(Int(Rnd() * 2))
  s.Outline.Width = 0#
End Function
