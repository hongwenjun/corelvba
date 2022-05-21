Attribute VB_Name = "CQL查找相同"
Sub 属性选择()
  CQL_FIND_UI.Show 0
End Sub

Public Function CQLline_CM100()
  On Error GoTo err
  Dim cm(3) As Color, i As Long
  Set cm(0) = CreateCMYKColor(100, 0, 0, 0)  '青
  Set cm(1) = CreateCMYKColor(0, 100, 0, 0)  '洋红
  Set cm(2) = CreateCMYKColor(100, 100, 0, 0) '洋红

ActiveDocument.ClearSelection
For i = 0 To 2
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
