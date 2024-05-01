Attribute VB_Name = "arrowtool"
Public Sub SetArrow()
  Dim s As Shape
  Set s = ActiveShape
  s.name = "arrow"
End Sub

Public Sub turn_over()
  Dim sr As ShapeRange, s As Shape
  Set sr = ActiveSelectionRange
  
  For Each s In sr
    s.RotationAngle = s.RotationAngle + 180
  Next s
End Sub


Sub arrow_Batch_repalce()
  Dim old As Shape, src As Shape, arrow_set As ShapeRange
  Dim nr As NodeRange
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  
  For Each old In sr
    Set nr = old.DisplayCurve.Nodes.All
    x1 = nr(1).PositionX
    y1 = nr(1).PositionY
    x2 = nr(2).PositionX
    y2 = nr(2).PositionY
    Angle = lineangle(x1, y1, x2, y2)
    
    Set src = old.Duplicate(0, 0)
    src.Rotate -Angle
    
    Set arrow_set = ActivePage.Shapes.FindShapes(Query:="@name ='arrow'")
    
    arrow_repalce arrow_set(1), src, Angle
    src.Delete: old.Delete
  Next old
End Sub


Sub arrow_repalce(arrow As Shape, src As Shape, ByVal Angle As Double)
  ActiveDocument.Unit = cdrMillimeter
  Set s = arrow.Duplicate(0, 0)
  s.name = "new_arrow"
  s.SizeWidth = src.SizeWidth
  s.SizeHeight = src.SizeHeight
  s.RotationAngle = Angle
  s.CenterX = src.CenterX: s.CenterY = src.CenterY
  
 ' If Angle > 180 Then s.RotationAngle = s.RotationAngle + 180
End Sub


 Sub arrow_manual_tool()
 Dim old As Shape, src As Shape, arrow_set As ShapeRange
 Dim nr As NodeRange
 Dim x1 As Double, y1 As Double
 Dim x2 As Double, y2 As Double
 Set nr = ActiveShape.Curve.Selection
 Set old = ActiveShape
 x1 = nr(1).PositionX
 y1 = nr(1).PositionY
 x2 = nr(2).PositionX
 y2 = nr(2).PositionY
 Angle = lineangle(x1, y1, x2, y2)

 Set src = old.Duplicate(0, 0)
' MsgBox Angle
 src.Rotate -Angle
 
 Set arrow_set = ActivePage.Shapes.FindShapes(Query:="@name ='arrow'")
 
 arrow_repalce arrow_set(1), src, Angle
 
 src.Delete: old.Delete
End Sub


' �����˵������,Ϊ(x1,y1)��(x2,y2) ��ô��Ƕ�a��tanֵ: tana=(y2-y1)/(x2-x1)
' ���Լ���arctan(y2-y1)/(x2-x1), �õ���Ƕ�ֵa
' VB����atn(), ����ֵ�ǻ��ȣ���Ҫ ���� PI /180
Private Function old_lineangle(x1, y1, x2, y2) As Double
  pi = 4 * VBA.Atn(1) ' ����Բ����
  If x2 = x1 Then
    lineangle = 90: Exit Function
  End If
  lineangle = VBA.Atn((y2 - y1) / (x2 - x1)) / pi * 180
End Function

Private Function lineangle(x1, y1, x2, y2) As Double
  If x2 = x1 Then lineangle = 90: Exit Function
  pi = 4 * VBA.Atn(1)

  k = (y2 - y1) / (x2 - x1)
  Angle = VBA.Atn(k) * 180 / pi
  
  If k >= 0 Then
    lineangle = Angle
  Else
    lineangle = Angle + 180
  End If
End Function
