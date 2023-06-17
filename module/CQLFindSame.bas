Attribute VB_Name = "CQLFindSame"
Public Function CQLline_CM100()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim cm(5) As Color, i As Long
  Set cm(0) = CreateCMYKColor(100, 0, 100, 0)  '// ÂÌ
  Set cm(1) = CreateCMYKColor(0, 100, 0, 0)    '// Ñóºì
  Set cm(2) = CreateCMYKColor(100, 100, 0, 0)  '// ºì
  Set cm(3) = CreateRGBColor(0, 255, 0)        '// RGB ÂÌ
  Set cm(4) = CreateRGBColor(255, 0, 0)        '// RGB ºì

  ActiveDocument.ClearSelection
  For i = 0 To 4
    cm(i).ConvertToRGB
    r = cm(i).RGBRed
    G = cm(i).RGBGreen
    b = cm(i).RGBBlue
    ActivePage.Shapes.FindShapes(Query:="@Outline.Color.rgb[.r='" & r & "' And .g='" & G & "' And .b='" & b & "']").AddToSelection
  Next i

ErrorHandler:
  API.EndOpt
End Function

