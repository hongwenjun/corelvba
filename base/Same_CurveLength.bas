Sub CurveLength()
 Dim s As Shape
 Set s = ActiveSelection.Shapes(1)
 If s.Type = cdrCurveShape Then
  MsgBox s.Curve.Length
 End If
 
 ActivePage.Shapes.FindShapes(Query:="@type ='curve' and @com.curve.length=3").CreateSelection
End Sub


Sub Same_CurveLength()
 Dim s As Shape
 Dim cl As Double
 Dim cql As String
   
 Set s = ActiveSelection.Shapes(1)
 If s.Type = cdrCurveShape Then
  cl = s.Curve.Length
  cql = "@type ='curve' and (@com.curve.length - " & cl & ").abs() < 0.1"
  ActivePage.Shapes.FindShapes(Query:=cql).CreateSelection
 End If

End Sub
