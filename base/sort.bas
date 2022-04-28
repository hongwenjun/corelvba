## CorelDRAW 好像没有多个物件的对准排列，工作中又经常用到，所以写了个简单代码
```
Sub 傻瓜火车排列()
  ActiveDocument.ReferencePoint = cdrBottomLeft  '// 设置对准基准 左下
  Dim ssr As ShapeRange, s As Shape     '// 定义选择物件数组 ssr， 和遍历物件 s
  Dim cnt As Integer                    '// 定义物件个数计数器
  Set ssr = ActiveSelectionRange
  cnt = 1
  
  For Each s In ssr
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).LeftX + ssr(cnt - 1).SizeWidth, ssr(cnt - 1).BottomY
    cnt = cnt + 1
  Next s
End Sub
```

## 修改优化
```
Sub 傻瓜火车排列()
  Dim ssr As ShapeRange, s As Shape
  Dim cnt As Integer
  Set ssr = ActiveSelectionRange
  cnt = 1
  
  ActiveDocument.ReferencePoint = cdrBottomLeft
  For Each s In ssr
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).RightX, ssr(cnt - 1).BottomY
    cnt = cnt + 1
  Next s

End Sub

Sub 傻瓜阶梯排列()
  Dim ssr As ShapeRange, s As Shape
  Dim cnt As Integer
  Set ssr = ActiveSelectionRange
  cnt = 1
  
  ActiveDocument.ReferencePoint = cdrTopLeft
  For Each s In ssr
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).LeftX, ssr(cnt - 1).BottomY
    cnt = cnt + 1
  Next s

End Sub
```

###  从左到右排序
```
Dim s As Shape
    Dim sr As ShapeRange
    ActiveDocument.Unit = cdrMillimeter
    Set sr = ActiveSelectionRange
    Dim i As Integer
    i = sr.count

    
    
 '   sr.sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
      sr.sort " @shape1.top>@shape2.top"
    sr.sort " @shape1.left<@shape2.left"
    Dim j As Integer
 
For j = 2 To i
   ' sr.Shapes.Item(j).TopY = sr.Shapes.Item(j - 1).TopY
    sr.Shapes.Item(j).LeftX = sr.Shapes.Item(j - 1).RightX + TextBox63
Next
```
