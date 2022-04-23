Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Dim pos_x As Variant
  Dim pos_Y As Variant
  pos_x = Array(307, 27)
  pos_Y = Array(64, 126, 188, 200)

  If Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(0)) < 30 Then
    Call CQLSameUniformColor
  ElseIf Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(1)) < 30 Then
    Call CQLSameOutlineColor
  ElseIf Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(2)) < 30 Then
    Call CQLSameSize
  ElseIf Abs(X - pos_x(1)) < 30 And Abs(Y - pos_Y(3)) < 30 Then
    CorelVBA.WebHelp "https://262235.xyz/index.php/tag/vba/"
  End If
  
  CQL_FIND_UI.Hide   ' show
End Sub

Private Sub CQLSameSize()
  ActiveDocument.Unit = cdrMillimeter
  Dim s As Shape
  Set s = ActiveShape
  If s Is Nothing Then Exit Sub
    
  If OptBt.Value = True Then
    ActiveDocument.ClearSelection
    OptBt.Value = 0
    CQL_FIND_UI.Hide
    
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim Shift As Long
    Dim box As Boolean
    box = ActiveDocument.GetUserArea(x1, y1, x2, y2, Shift, 10, False, cdrCursorWeldSingle)
    If Not b Then
      ' MsgBox "选区范围: " & x1 & y1 & x2 & y2
      Set sh = ActivePage.SelectShapesFromRectangle(x1, y1, x2, y2, False)
      sh.Shapes.FindShapes(Query:="@width = {" & s.SizeWidth & " mm} and @height ={" & s.SizeHeight & "mm}").CreateSelection
    End If
  Else
    ActivePage.Shapes.FindShapes(Query:="@width = {" & s.SizeWidth & " mm} and @height ={" & s.SizeHeight & "mm}").CreateSelection
  End If
End Sub

Private Sub CQLSameOutlineColor()
  On Error GoTo err
  Dim colr As New Color, s As Shape
  Set s = ActiveShape
  If s Is Nothing Then Exit Sub
  colr.CopyAssign s.Outline.Color
  colr.ConvertToRGB
  ' 查找对象
  r = colr.RGBRed
  G = colr.RGBGreen
  b = colr.RGBBlue
  
  If OptBt.Value = True Then
    ActiveDocument.ClearSelection
    OptBt.Value = 0
    CQL_FIND_UI.Hide
    
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim Shift As Long
    Dim box As Boolean
    box = ActiveDocument.GetUserArea(x1, y1, x2, y2, Shift, 10, False, cdrCursorWeldSingle)
    If Not b Then
      ' MsgBox "选区范围: " & x1 & y1 & x2 & y2
      Set sh = ActivePage.SelectShapesFromRectangle(x1, y1, x2, y2, False)
      sh.Shapes.FindShapes(Query:="@Outline.Color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").CreateSelection
    End If
  Else
    ActivePage.Shapes.FindShapes(Query:="@Outline.Color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").CreateSelection
  End If
  
  Exit Sub
err:
    MsgBox "对象轮廓为空。"
End Sub

Private Sub CQLSameUniformColor()
  On Error GoTo err
  Dim colr As New Color, s As Shape
  Set s = ActiveShape
  If s Is Nothing Then Exit Sub
  If s.Fill.Type = cdrFountainFill Then MsgBox "不支持渐变色。": Exit Sub
  colr.CopyAssign s.Fill.UniformColor
  colr.ConvertToRGB
  ' 查找对象
  r = colr.RGBRed
  G = colr.RGBGreen
  b = colr.RGBBlue
  
  If OptBt.Value = True Then
    ActiveDocument.ClearSelection
    OptBt.Value = 0
    CQL_FIND_UI.Hide
    
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim Shift As Long
    Dim box As Boolean
    box = ActiveDocument.GetUserArea(x1, y1, x2, y2, Shift, 10, False, cdrCursorWeldSingle)
    If Not b Then
      ' MsgBox "选区范围: " & x1 & y1 & x2 & y2
      Set sh = ActivePage.SelectShapesFromRectangle(x1, y1, x2, y2, False)
      sh.Shapes.FindShapes(Query:="@fill.color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").CreateSelection
    End If
  Else
    ActivePage.Shapes.FindShapes(Query:="@fill.color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").CreateSelection
  End If
  Exit Sub
err:
  MsgBox "对象填充为空。"
End Sub
