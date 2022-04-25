Sub TestMacro()
  ActiveDocument.Unit = cdrMillimeter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Set shs = ActiveSelection.Shapes
  For Each sh In shs
    Dim eff1 As Effect
    Set eff1 = sh.CreateContour(cdrContourOutside, 5, 1, cdrDirectFountainFillBlend, CreateRGBColor(26, 22, 35), CreateCMYKColor(0, 0, 0, 100), CreateCMYKColor(0, 0, 0, 100), 0, 0, cdrContourSquareCap, cdrContourCornerMiteredOffsetBevel, 15#)
    eff1.Contour.ContourGroup.Shapes(1).AddToSelection
    eff1.Separate
  Next sh

    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Set sh = OrigSelection.CustomCommand("Boundary", "CreateBoundary")

  ActiveSelection.Shapes.FindShapes(Query:="@Outline.Color=RGB(26, 22, 35)").CreateSelection
  For Each sh In ActiveSelection.Shapes
    sh.Delete
  Next sh

End Sub
