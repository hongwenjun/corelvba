Sub mm()
    Application.FrameWork.Automation.Invoke "8e843a39-b9a2-a7b3-4714-21523261745f"
End Sub


Sub make_PageMark()
  ActiveDocument.Unit = cdrMillimeter
  '// 获得页面中心点 x,y ; 页面大小
  px = ActivePage.CenterX
  py = ActivePage.CenterY
  Pw = ActivePage.SizeWidth
  Ph = ActivePage.SizeHeight

  '// 开始画圆
  Dim s As Shape
  Set s = ActiveLayer.CreateEllipse2(px, py, Pw / 2, Ph / 2)   '// 页面尺寸的圆
  
  r = 6# / 2    '// 圆直径6mm
  Set s1 = ActiveLayer.CreateEllipse2(8, 8, r, r)
  Set s2 = ActiveLayer.CreateEllipse2(Pw - 8, 8, r, r)
  Set s3 = ActiveLayer.CreateEllipse2(8, Ph - 8, r, r)
  Set s4 = ActiveLayer.CreateEllipse2(Pw - 8, Ph - 8, r, r)
  
  Set s3fz = ActiveLayer.CreateRectangle2(8 + r, Ph - 8 - 1 + r, 2, 1)
  
  '// 使用 ShapeRange 批量物件修改颜色和群组
  Dim sr As New ShapeRange
  sr.Add s1: sr.Add s2: sr.Add s3: sr.Add s4: sr.Add s3fz
  
  sr.ApplyUniformFill CreateCMYKColor(0, 0, 0, 100)
  
  For Each sh In sr
    sh.Outline.SetNoOutline
  Next sh
  
  '// 组合，建立名字
  Set s = sr.Combine
  s.Name = "RoundMark"
  s.AddToSelection
End Sub


Public Sub page_add_Rect()
  Dim sr As New ShapeRange
  W = 5: H = 5: x = 5
  x2 = ActivePage.SizeWidth - 10
  y = ActivePage.SizeHeight - 50
  
  For I = 1 To (ActivePage.SizeHeight + 140) / 160
    Set s1 = ActiveLayer.CreateRectangle2(x, y, W, H)
    Set s2 = ActiveLayer.CreateRectangle2(x2, y, W, H)
    y = y - 160
    sr.Add s1: sr.Add s2   '// 添加到sr 用以群组修改
  Next I
  
  '// 改颜色，群组选择
  sr.ApplyUniformFill CreateCMYKColor(0, 0, 0, 100)
  sr.Group: sr.CreateSelection
End Sub
