Attribute VB_Name = "ClipbRectangle"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "剪贴板尺寸建立矩形"  Clipboard Size Build Rectangle  2023.6.11

Type Coordinate
    X As Double
    Y As Double
End Type
Public O_O As Coordinate

Sub Start()
    '// 坐标原点
    O_O.X = 0:   O_O.Y = 0
    Dim ost As ShapeRange
    Set ost = ActiveSelectionRange

    O_O.X = ost.LeftX
    O_O.Y = ost.BottomY - 50    '选择物件 下移动 50mm

    '// 建立矩形 Width  x Height 单位 mm
    Dim Str, arr, n
    Str = API.GetClipBoardString

    ' 替换 mm x * 换行 TAB 为空格
    Str = VBA.Replace(Str, "m", " ")
    Str = VBA.Replace(Str, "x", " ")
    Str = VBA.Replace(Str, "X", " ")
    Str = VBA.Replace(Str, "*", " ")
    Str = VBA.Replace(Str, vbNewLine, " ")

    Do While InStr(Str, "  ") '多个空格换成一个空格
        Str = VBA.Replace(Str, "  ", " ")
    Loop
    arr = Split(Str)
    
    ActiveDocument.BeginCommandGroup  '一步撤消'
    Dim X As Double
    Dim Y As Double
    For n = LBound(arr) To UBound(arr) - 1 Step 2
        ' MsgBox arr(n)
        X = Val(arr(n))
        Y = Val(arr(n + 1))
        
        If X > 0 And Y > 0 Then
            Rectangle X, Y
            O_O.X = O_O.X + X + 30
        End If
    Next
    ActiveDocument.EndCommandGroup
End Sub

'// 建立矩形 Width  x Height 单位 mm
Private Function Rectangle(Width As Double, Height As Double)
  ActiveDocument.Unit = cdrMillimeter
  Dim size As Shape
  Dim d As Document
  Dim s1 As Shape

  '// 建立矩形 Width  x Height 单位 mm
  Set s1 = ActiveLayer.CreateRectangle(O_O.X, O_O.Y, O_O.X + Width, O_O.Y - Height)
  
  '// 填充颜色无，轮廓颜色 K100，线条粗细0.3mm
  s1.Fill.ApplyNoFill
  s1.Outline.SetProperties 0.3, OutlineStyles(0), CreateCMYKColor(0, 100, 0, 0), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
      
  sw = s1.SizeWidth
  sh = s1.SizeHeight

  text = Trim(Str(sw)) + "x" + Trim(Str(sh)) + "mm"
  Set d = ActiveDocument
  Set size = d.ActiveLayer.CreateArtisticText(O_O.X + sw / 2 - 25, O_O.Y + 10, text, Font:="Tahoma")  '// O_O.y + 10  标注尺寸上移 10mm
  size.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
End Function

' 测试矩形变形
Private Function setRectangle(Width As Double, Height As Double)

  Dim s1 As Shape
  Set s1 = ActiveSelection
  ActiveDocument.Unit = cdrMillimeter
  '// 物件中心基准, 先把宽度设定为
  ActiveDocument.ReferencePoint = cdrCenter
  s1.SetSize Height, Height

  '// 物件旋转 30度，轮廓线1mm ,轮廓颜色 M100Y100
  s1.Rotate 30#
  s1.Outline.SetProperties 1#
  s1.Outline.SetProperties Color:=CreateCMYKColor(0, 100, 100, 0)

End Function


'// 获得选择物件大小信息
Sub get_all_size()
  ActiveDocument.Unit = cdrMillimeter
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.CreateTextFile("R:\size.txt", True)
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String
  For Each sh In shs
    size = Trim(Str(Int(sh.SizeWidth + 0.5))) + "x" + Trim(Str(Int(sh.SizeHeight + 0.5))) + "mm"
    f.WriteLine (size)
    s = s + size + vbNewLine
  Next sh
  f.Close
  MsgBox "输出物件尺寸信息到文件" & "R:\size.txt" & vbNewLine & s
  API.WriteClipBoard s
End Sub


