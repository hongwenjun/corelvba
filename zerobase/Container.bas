Attribute VB_Name = "Container"
' �� �����������
Public Function SetBoxName()
  API.BeginOpt "�����������"
  
  Dim box As ShapeRange, s As Shape
  Set box = ActiveSelectionRange
  
  ' ����������֣��Թ�CQL��ѯ
  For Each s In box
    s.name = "Container"
  Next s
  
  API.EndOpt
  MsgBox "�����������" & vbNewLine & "����: Container"
End Function

' ͼƬ������������
Public Sub Batch_ToPowerClip()
  API.BeginOpt "������������"
  Dim s As Shape, ssr As ShapeRange, box As ShapeRange
  Set ssr = Smart_Group(0.5) ' ����Ⱥ���ݲ� 0.5mm
  
  For Each s In ssr
    Image_ToPowerClip s
  Next s

  API.EndOpt
End Sub

' ͼƬ������������������
Public Function Image_ToPowerClip(arg As Shape)
  Dim box As ShapeRange
  Dim ssr As New ShapeRange, rmsr As New ShapeRange
  Set ssr = arg.UngroupEx
  ' CQL�������������
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  box.SetOutlineProperties Width:=0, Color:=Nothing
  ssr.AddToPowerClip box(1), 0
  box(1).name = "powerclip_ok"

End Function

' ͼƬOneKey��������
Public Sub OneKey_ToPowerClip()
  API.BeginOpt "ͼƬOneKey��������"
  Dim s As Shape, ssr As ShapeRange, box As ShapeRange
  
  ' �������������͸��
  Set box = ActiveSelectionRange
  For Each s In box
    If s.Type <> cdrBitmapShape Then s.name = "Container"
  Next s
  
  Set ssr = Smart_Group(0.5) ' ����Ⱥ���ݲ� 0.5mm
  
  Application.Optimization = True
  For Each s In ssr
    Image_ToPowerClip s
  Next s

  API.EndOpt
End Sub

' �� ɾ���������ӱ߽���������    �ۢ�
Public Function Remove_OutsideBox()
  Dim s As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim rmsr As New ShapeRange
  Dim x As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  ' CQL�������������
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    x = s.CenterX: Y = s.CenterY
    If box(1).IsOnShape(x, Y) = cdrOutsideShape Then rmsr.Add s
  Next s

  rmsr.Delete
End Function


Public Function Remove_OnMargin()
  Dim s As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim rmsr As New ShapeRange
  Dim x As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  ' CQL�������������
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    x = s.CenterX: Y = s.CenterY
    If box(1).IsOnShape(x, Y) = cdrOnMarginOfShape Then rmsr.Add s
  Next s

  rmsr.Delete
End Function


Public Function Select_OutsideBox()
  Dim s As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim SelSr As New ShapeRange
  Dim x As Double, Y As Double, radius
  
  Set ssr = ActiveSelectionRange
  ' CQL�������������
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    x = s.CenterX: Y = s.CenterY
    radius = s.SizeWidth / 2
    If box(1).IsOnShape(x, Y, radius) = cdrOutsideShape Then SelSr.Add s
  Next s
  
  ActiveDocument.ClearSelection
  SelSr.AddToSelection

End Function


Public Function Select_OnMargin()
  Dim s As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim SelSr As New ShapeRange
  Dim x As Double, Y As Double, radius
  
  Set ssr = ActiveSelectionRange
  ' CQL�������������
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.Count = 0 Then Exit Function
  
  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    x = s.CenterX: Y = s.CenterY
    radius = s.SizeWidth / 2
    If box(1).IsOnShape(x, Y, radius) = cdrOnMarginOfShape Then SelSr.Add s
  Next s
  
  ActiveDocument.ClearSelection
  SelSr.AddToSelection

End Function


' ����ӳ���������󣬵��ý�ɢ����;���
Public Sub Batch_Center()
    Dim s As Shape, ssr As ShapeRange
    Set ssr = Smart_Group
    For Each s In ssr
      Ungroup_Center s
    Next s
End Sub


' ���º�������ɢ�����������������
Private Function Ungroup_Center(os As Shape)
    Set grp = os.UngroupEx
    grp.Sort "@shape1.Width * @shape1.Height> @shape2.Width * @shape2.Height"
    cx = grp(1).CenterX
    cy = grp(1).CenterY
    For Each s In grp
      s.CenterX = cx
      s.CenterY = cy
    Next s
End Function

