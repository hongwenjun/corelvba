'// Attribute VB_Name = "剪贴板尺寸建立矩形"
Type Coordinate
    x As Double
    y As Double
End Type
Public O_O As Coordinate

Sub start()
    '// 坐标原点
    O_O.x = 0:   O_O.y = 0
    Dim ost As ShapeRange
    Set ost = ActiveSelectionRange

    O_O.x = ost.LeftX
    O_O.y = ost.BottomY - 50    '选择物件 下移动 50mm

    '// 建立矩形 Width  x Height 单位 mm
    ' Rectangle 101, 151
    Dim Str, arr, n
    Str = GetClipBoardString

    ' 替换 mm x * 换行 TAB 为空格
    Str = VBA.Replace(Str, "mm", " ")
    Str = VBA.Replace(Str, "x", " ")
    Str = VBA.Replace(Str, "*", " ")
    Str = VBA.Replace(Str, Chr(13), " ")
    Str = VBA.Replace(Str, Chr(9), " ")
    
    Do While InStr(Str, "  ") '多个空格换成一个空格
        Str = VBA.Replace(Str, "  ", " ")
    Loop
    arr = Split(Str)
    
    ActiveDocument.BeginCommandGroup  '一步撤消'
    Dim x As Double
    Dim y As Double
    For n = LBound(arr) To UBound(arr) - 1 Step 2
        ' MsgBox arr(n)
        x = Val(arr(n))
        y = Val(arr(n + 1))
        
        If x > 0 And y > 0 Then
            Rectangle x, y
            O_O.x = O_O.x + x + 30
        End If
    Next
    ActiveDocument.EndCommandGroup
End Sub

Private Function Rectangle(Width As Double, Height As Double)
    ActiveDocument.Unit = cdrMillimeter
    Dim size As Shape
    Dim d As Document
    Dim s1 As Shape

    '// 建立矩形 Width  x Height 单位 mm
    Set s1 = ActiveLayer.CreateRectangle(O_O.x, O_O.y, O_O.x + Width, O_O.y - Height)
    
    '// 填充颜色无，轮廓颜色 K100，线条粗细0.3mm
    s1.Fill.ApplyNoFill
    s1.Outline.SetProperties 0.3, OutlineStyles(0), CreateCMYKColor(0, 100, 0, 0), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
        
    sw = s1.SizeWidth
    sh = s1.SizeHeight

    Text = Trim(Str(sw)) + "x" + Trim(Str(sh)) + "mm"
    Set d = ActiveDocument
    Set size = d.ActiveLayer.CreateArtisticText(O_O.x + sw / 2 - 25, O_O.y + 10, Text)  '// O_O.y + 10  标注尺寸上移 10mm
    size.Fill.UniformColor.CMYKAssign 0, 100, 100, 0
End Function

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

Private Function GetClipBoardString() As String
    On Error Resume Next
    Dim MyData As New DataObject
    GetClipBoardString = ""
    MyData.GetFromClipboard
    GetClipBoardString = MyData.GetText
    Set MyData = Nothing
End Function
