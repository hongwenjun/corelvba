Attribute VB_Name = "ModulePlus"
'// 断开所有节点 分割线段
Public Function SplitSegment()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim ssr As ShapeRange, s As Shape
  Dim nr As NodeRange, nd As Node
  
  Set ssr = ActiveSelectionRange
  Set s = ssr.UngroupAllEx.Combine
  Set nr = s.Curve.Nodes.all
  
  nr.BreakApart
  s.BreakApartEx
  
ErrorHandler:
  API.EndOpt
End Function

'// 批量正方形 宽高统一
Public Function square_hw(Optional ByVal hw As String = "Height")
  API.BeginOpt
  Set os = ActiveSelectionRange
  Set ss = os.Shapes
  For Each s In ss
    If hw = "Height" Then s.SizeWidth = s.SizeHeight
    If hw = "Width" Then s.SizeHeight = s.SizeWidth
  Next s
  API.EndOpt
End Function

'// 节点优化减少
Public Function Nodes_Reduce()
  On Error GoTo ErrorHandler: API.BeginOpt
  Set doc = ActiveDocument
  Dim s As Shape
  ps = Array(1)
  doc.Unit = cdrTenthMicron
  Set os = ActivePage.Shapes
  If os.count > 0 Then
    For Each s In os
    s.ConvertToCurves
      If Not s.DisplayCurve Is Nothing Then
        s.Curve.AutoReduceNodes 50
      End If
    Next s
  End If
ErrorHandler:
  API.EndOpt
End Function

'// 标注线 选择文字 删除等
Public Function Dimension_Select_or_Delete(Shift As Long)
  On Error GoTo ErrorHandler: API.BeginOpt
  Dim os As ShapeRange, sr As ShapeRange, s As Shape
  Set doc = ActiveDocument
  Set sr = ActiveSelectionRange
  sr.RemoveAll

  If Shift = 4 Then
    Set os = ActiveSelectionRange
    For Each s In os.Shapes
      If s.Type = cdrTextShape Then sr.Add s
    Next s
  sr.CreateSelection
  
  ElseIf Shift = 1 Then
    Set os = ActiveSelectionRange
    For Each s In os.Shapes
      If s.Type = cdrLinearDimensionShape Then sr.Add s
    Next s
    sr.CreateSelection
    
  ElseIf Shift = 2 Then
    Set os = ActiveSelectionRange
    For Each s In os.Shapes
      If s.Type = cdrLinearDimensionShape Then sr.Add s
    Next s
    sr.Delete
    If os.count > 0 Then
      os.Shapes.FindShapes(Query:="@name ='DMKLine'").CreateSelection
      ActiveSelectionRange.Delete
    End If
  End If
  
ErrorHandler:
  API.EndOpt
End Function

'// 解绑尺寸，分离尺寸
Public Function Untie_MarkLines()
  On Error GoTo ErrorHandler: API.BeginOpt
  
  Dim os As ShapeRange, dss As New ShapeRange
  Set os = ActiveSelectionRange
  For Each s In os.Shapes
      If s.Type = cdrLinearDimensionShape Then
        dss.Add s
      End If
  Next s
  If dss.count > 0 Then
    dss.BreakApartEx
    os.Shapes.FindShapes(Query:="@name ='DMKLine'").CreateSelection
    ActiveSelectionRange.Delete
  End If
  
ErrorHandler:
  API.EndOpt
End Function

'// 函数：判断 ShapeRange 中的所有物件尺寸是否相同
Function IsAllSameSize(sr As ShapeRange) As Boolean
    Dim s As Shape
    Dim tol As Double
    tol = 0.01
 
    For Each s In sr
        If Abs(s.SizeWidth - sr.FirstShape.SizeWidth) > tol Or _
           Abs(s.SizeHeight - sr.FirstShape.SizeHeight) > tol Then
            IsAllSameSize = False
            Exit Function
        End If
    Next s

    IsAllSameSize = True
End Function

