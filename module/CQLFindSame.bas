Attribute VB_Name = "CQL查找相同"
Sub 属性选择()
  CQL_FIND_UI.Show 0
End Sub

Public Function CQLline_CM100()
  On Error GoTo err
  Dim cm(5) As Color, I As Long
  Set cm(0) = CreateCMYKColor(100, 0, 100, 0)  '绿
  Set cm(1) = CreateCMYKColor(0, 100, 0, 0)  '洋红
  Set cm(2) = CreateCMYKColor(100, 100, 0, 0) '红
  Set cm(3) = CreateRGBColor(0, 255, 0) ' RGB 绿
  Set cm(4) = CreateRGBColor(255, 0, 0) ' RGB 红

  ActiveDocument.ClearSelection
  For I = 0 To 4
    cm(I).ConvertToRGB
    r = cm(I).RGBRed
    G = cm(I).RGBGreen
    B = cm(I).RGBBlue
    ActivePage.Shapes.FindShapes(Query:="@Outline.Color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & B & "']").AddToSelection
  Next I

Exit Function
err:
  MsgBox "Function CQLline_CM100 错误!"
End Function
