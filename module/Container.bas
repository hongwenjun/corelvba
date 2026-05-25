Attribute VB_Name = "Container"
Public Function SetBoxName()
  API.BeginOpt "Undo SetBoxName"
  
  Dim box As ShapeRange, s As Shape
  Set box = ActiveSelectionRange
  
  For Each s In box
    s.name = "Container"
  Next s
  
  API.EndOpt
End Function


Public Function Batch_ToPowerClip()
  API.BeginOpt "Batch_ToPowerClip"
  Dim s As Shape, ssr As ShapeRange, box As ShapeRange
  Set ssr = API.Smart_Group(0.5)
  
  For Each s In ssr
    Image_ToPowerClip s
  Next s

  API.EndOpt
End Function

Public Function Image_ToPowerClip(arg As Shape)
  API.BeginOpt "ToPowerClip"
  Dim box As ShapeRange
  Dim ssr As New ShapeRange, rmsr As New ShapeRange
  Set ssr = arg.UngroupEx
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.count = 0 Then Exit Function
  
  box.SetOutlineProperties width:=0, Color:=Nothing
  ssr.AddToPowerClip box(1), 0
  box(1).name = "powerclip_ok"
  API.EndOpt
End Function

Public Function OneKey_ToPowerClip()
  API.BeginOpt "OneKey_ToPowerClip"
  Dim s As Shape, ssr As ShapeRange, box As ShapeRange
  
  Set box = ActiveSelectionRange
  For Each s In box
    If s.Type <> cdrBitmapShape Then s.name = "Container"
  Next s
  
  Set ssr = API.Smart_Group(0.5)
  
  Application.Optimization = True
  For Each s In ssr
    Image_ToPowerClip s
  Next s

  API.EndOpt
End Function

Public Function Remove_OutsideBox(radius As Double)
  API.BeginOpt "Undo Remove"
  On Error GoTo ErrorHandler
  Dim s As Shape, bc As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim rmsr As New ShapeRange
  Dim X As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.count = 0 Then GoTo ErrorHandler
  Set bc = box(1).Duplicate(0, 0)
  If bc.Type = cdrTextShape Then bc.ConvertToCurves
  
  For Each s In ssr
    X = s.CenterX: Y = s.CenterY
    If bc.IsOnShape(X, Y, radius) = cdrOutsideShape Then rmsr.Add s
  Next s
  
  rmsr.Add bc: rmsr.Delete: API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
  On Error Resume Next

End Function

Public Function Select_SideBox(side As cdrPositionOfPointOverShape)
  On Error GoTo ErrorHandler
  API.BeginOpt "Undo Select"
  Dim s As Shape, bc As Shape
  Dim ssr As ShapeRange, box As ShapeRange
  Dim SelSr As New ShapeRange
  Dim X As Double, Y As Double, radius As Double
  If GlobalUserData.Exists("Tolerance", 1) Then radius = Val(GlobalUserData("Tolerance", 1))
  
  Set ssr = ActiveSelectionRange
  If ssr.count = 1 Then ssr.AddRange ActivePage.Shapes.FindShapes(Query:="!@name ='Container'")
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.count = 0 Then GoTo ErrorHandler
  
  Set bc = box(1).Duplicate(0, 0)
  bc.Fill.ApplyUniformFill CreateCMYKColor(0, 100, 0, 0)
  If bc.Type = cdrTextShape Then bc.ConvertToCurves

  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    X = s.CenterX: Y = s.CenterY
    If side = (cdrInsideShape + cdrOnMarginOfShape) Then
      If bc.IsOnShape(X, Y, s.SizeWidth / 2 * radius) = cdrInsideShape Then SelSr.Add s
      If bc.IsOnShape(X, Y, s.SizeWidth / 2 * radius) = cdrOnMarginOfShape Then SelSr.Add s
    Else
      If bc.IsOnShape(X, Y, s.SizeWidth / 2 * radius) = side Then SelSr.Add s
    End If
  Next s
  
  ActiveDocument.ClearSelection
  bc.Delete: SelSr.AddToSelection: API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
End Function


Public Function Select_by_BlendGroup(radius As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt "Undo Select"
  Dim s As Shape, bc As Shape
  Dim ssr As ShapeRange, box As ShapeRange, gp As ShapeRange
  Dim SelSr As New ShapeRange
  Dim X As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.count = 0 Then GoTo ErrorHandler
  Set gp = box.Duplicate(0, 0).UngroupAllEx
  Set gp = gp.BreakApartEx.UngroupAllEx

  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    X = s.CenterX: Y = s.CenterY
    For Each bc In gp
      If bc.IsOnShape(X, Y, s.SizeWidth / 2 * radius) = cdrOnMarginOfShape Then SelSr.Add s
    Next bc
  Next s
  
  ActiveDocument.ClearSelection
  gp.Delete: SelSr.AddToSelection: API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function

Public Function Select_Quick_BlendGroup(radius As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt "Undo Select"
  Dim s As Shape, bc As Shape
  Dim ssr As ShapeRange, box As ShapeRange, gp As ShapeRange
  Dim SelSr As New ShapeRange
  Dim X As Double, Y As Double
  
  Set ssr = ActiveSelectionRange
  Set box = ssr.Shapes.FindShapes(Query:="@name ='Container'")
  ssr.RemoveRange box
  
  If box.count = 0 Then GoTo ErrorHandler
  Set gp = box.Duplicate(0, 0).UngroupAllEx
  Set bc = gp.BreakApartEx.UngroupAllEx.Combine

  ActiveDocument.Unit = cdrMillimeter
  For Each s In ssr
    X = s.CenterX: Y = s.CenterY
    If bc.IsOnShape(X, Y, s.SizeWidth / 2 * radius) = cdrOnMarginOfShape Then SelSr.Add s
  Next s
  
  ActiveDocument.ClearSelection
  bc.Delete: SelSr.AddToSelection: API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
  On Error Resume Next
End Function


