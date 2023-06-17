Attribute VB_Name = "Arrange"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "物件排列拼版"   Arrange  2023.6.11

'// CorelDRAW 物件排列拼版简单代码
Public Function Arrange()
  On Error GoTo ErrorHandler
  API.BeginOpt
  ActiveDocument.Unit = cdrMillimeter
  row = 3     ' 拼版 3 x 4
  List = 4
  sp = 0       '间隔 0mm

  Dim Str, arr, n
  Str = API.GetClipBoardString

  ' 替换 mm x * 换行 TAB 为空格
  Str = VBA.Replace(Str, "mm", " ")
  Str = VBA.Replace(Str, "x", " ")
  Str = VBA.Replace(Str, "X", " ")
  Str = VBA.Replace(Str, "*", " ")

  '// 换行转空格 多个空格换成一个空格
  Str = API.Newline_to_Space(Str)
  
  arr = Split(Str)

  Dim s1 As Shape
  Dim X As Double, Y As Double
  
  If 0 = ActiveSelectionRange.Count Then
    X = Val(arr(0)):    Y = Val(arr(1))
    row = Int(ActiveDocument.Pages.First.SizeWidth / X)
    List = Int(ActiveDocument.Pages.First.SizeHeight / Y)

    If UBound(arr) > 2 Then
    row = Val(arr(2)):  List = Val(arr(3))
      If row * List > 8000 Then
        GoTo ErrorHandler
      ElseIf UBound(arr) > 3 Then
          sp = Val(arr(4))       '间隔
      End If
    End If
     
    '// 建立矩形 Width  x Height 单位 mm
    Set s1 = ActiveLayer.CreateRectangle(0, 0, X, Y)
    
    '// 填充颜色无，轮廓颜色 K100，线条粗细0.3mm
    s1.Fill.ApplyNoFill
    s1.Outline.SetProperties 0.3, OutlineStyles(0), CreateCMYKColor(0, 100, 0, 0), ArrowHeads(0), _
      ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#

  '// 如果当前选择物件，按当前物件拼版
  ElseIf 1 = ActiveSelectionRange.Count Then
    Set s1 = ActiveSelection
    X = s1.SizeWidth:    Y = s1.SizeHeight
    row = Int(ActiveDocument.Pages.First.SizeWidth / X)
    List = Int(ActiveDocument.Pages.First.SizeHeight / Y)
  End If
  
  sw = X:  sh = Y

  '// StepAndRepeat 方法在范围内创建多个形状副本
  Dim dup1 As ShapeRange
  Set dup1 = s1.StepAndRepeat(row - 1, sw + sp, 0#)
  Dim dup2 As ShapeRange
  Set dup2 = ActiveDocument.CreateShapeRangeFromArray(dup1, s1).StepAndRepeat(List - 1, 0#, (sh + sp))
       
ErrorHandler:
  API.EndOpt
End Function


