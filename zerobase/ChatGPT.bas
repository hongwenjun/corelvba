Attribute VB_Name = "ChatGPT"
Private Type Coordinate
    x As Double
    y As Double
End Type

Sub Z序排列()

  ActiveDocument.Unit = cdrMillimeter
  Dim dot As Coordinate
  Dim s As Shape, ssr As ShapeRange
  Dim cnt As Long: cnt = 1
  Set ssr = ActiveSelectionRange

  ssr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
  
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    s.OrderToFront
    puts dot.x, dot.y, cnt: cnt = cnt + 1
  Next s
End Sub

Sub U序排列()

  ActiveDocument.Unit = cdrMillimeter
  Set xdict = CreateObject("Scripting.dictionary")
  Set ydict = CreateObject("Scripting.dictionary")
  Dim dot As Coordinate
  Dim s As Shape, ssr As ShapeRange
  Dim cnt As Long: cnt = 1
  Set ssr = ActiveSelectionRange
  
  ssr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
  
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    If xdict.Exists(Int(dot.x)) = False Then xdict.Add Int(dot.x), dot.x
    If ydict.Exists(Int(dot.y)) = False Then ydict.Add Int(dot.y), dot.y
  Next s
  
  inverter = 1   ' 交流频率控制
  xc = xdict.Count: yc = ydict.Count

  For cnt = 0 To ydict.Count - 1
    If inverter Mod 2 = 0 Then
        ssr.Sort " @shape1.Left > @shape2.Left", cnt * xc + 1, cnt * xc + xc
    Else
        ssr.Sort " @shape1.Left < @shape2.Left", cnt * xc + 1, cnt * xc + xc
    End If
    inverter = inverter + 1
  Next cnt
  
  cnt = 1
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    s.OrderToFront
    puts dot.x, dot.y, cnt: cnt = cnt + 1
  Next s
  
End Sub


Sub 计算行列()   ' 字典使用计算行列

  ActiveDocument.Unit = cdrMillimeter
  Set xdict = CreateObject("Scripting.dictionary")
  Set ydict = CreateObject("Scripting.dictionary")
  Dim dot As Coordinate, Offset As Coordinate
  Dim s As Shape, ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  
  ' 当前选择物件的范围边界
  set_lx = ssr.LeftX: set_rx = ssr.RightX
  set_by = ssr.BottomY: set_ty = ssr.TopY
  ssr(1).GetSize Offset.x, Offset.y
  ' 当前选择物件 ShapeRange 初步排序
  ssr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
  
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    If xdict.Exists(Int(dot.x)) = False Then xdict.Add Int(dot.x), dot.x
    If ydict.Exists(Int(dot.y)) = False Then ydict.Add Int(dot.y), dot.y
  Next s
  
'  MsgBox "字典使用计算行列:" & xdict.Count & ydict.Count
  Dim cnt As Long: cnt = 1
  
  ' 遍历字典，输出
  Dim key As Variant
  For Each key In xdict.keys
      dot.x = xdict(key)
      puts dot.x, set_by - Offset.y / 2, cnt
      cnt = cnt + 1
  Next key
  
  cnt = 1
  For Each key In ydict.keys
      dot.y = ydict(key)
      puts set_lx - Offset.x / 2, dot.y, cnt
      cnt = cnt + 1
  Next key
  
End Sub

Private Sub puts(x, y, n)
  Dim st As String
  st = str(n)
  Set s = ActiveLayer.CreateArtisticText(0, 0, st)
  s.CenterX = x: s.CenterY = y
End Sub

'// 对数组进行排序[单维]
Public Function ArraySort(src As Variant) As Variant
  Dim out As Long, i As Long, tmp As Variant
  For out = LBound(src) To UBound(src) - 1
    For i = out + 1 To UBound(src)
      If src(out) > src(i) Then
        tmp = src(i): src(i) = src(out): src(out) = tmp
      End If
    Next i
  Next out
  
  ArraySort = src
End Function



Sub ShowMessage()
    MsgBox "Hello, World!"
End Sub


Sub DictionaryExample()
    ' 创建一个空的Dictionary
    Dim myDict As Object
    Set myDict = CreateObject("Scripting.Dictionary")
    
    ' 向Dictionary中添加键值对
    myDict.Add "orange", 4
    myDict.Add "banana", 2
    myDict.Add "apple", 3
    
    ' 访问键值对
    Debug.Print "The value of 'apple' is " & myDict("apple")
    
    ' 遍历Dictionary中的所有键值对
    Dim key As Variant
    For Each key In myDict.keys
        Debug.Print key & " : " & myDict(key)
    Next key
    
    ' 检查某个键是否存在
    If myDict.Exists("orange") Then
        Debug.Print "The key 'orange' exists"
    End If
    
    ' 删除某个键值对
    myDict.Remove "banana"
    
    ' 清空Dictionary
    myDict.RemoveAll
End Sub

Sub tongji使用字典统计()

  Dim s As Shape
  Dim sr As ShapeRange
  
  Set sr = ActiveSelection.Shapes.FindShapes(Query:="@name='wk-y标记'")
  
  Dim stn As String, str As String
  
  Set d = CreateObject("Scripting.dictionary")
  
  For Each s In sr
    If s.Type = cdrTextShape Then
      If s.text.Type = cdrArtistic Then
        stn = s.text.Story.text
        If d.Exists(stn) = True Then
          d.Item(stn) = d.Item(stn) + 1
        Else
          d.Add stn, 1
        End If: End If: End If
  Next s
  
  str = "   规   格" & vbTab & vbTab & vbTab & "数量" & vbNewLine

  a = d.keys: b = d.items
  For i = 0 To d.Count - 1
    str = str & a(i) & vbTab & vbTab & b(i) & "条" & vbNewLine
  Next

  ' 遍历Dictionary中的所有键值对
  Dim key As Variant
  For Each key In d.keys
      Debug.Print key & " : " & d(key)
  Next key

  Debug.Print str
End Sub




Sub 正式U序排列()
  Application.Optimization = True
  ActiveDocument.BeginCommandGroup  '一步撤消'

  ActiveDocument.Unit = cdrMillimeter
  Set xdict = CreateObject("Scripting.dictionary")
  Set ydict = CreateObject("Scripting.dictionary")
  Dim dot As Coordinate
  Dim s As Shape, ssr As ShapeRange
  Dim cnt As Long: cnt = 1
  Set ssr = ActiveSelectionRange
  
  ssr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
  
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    If xdict.Exists(Int(dot.x)) = False Then xdict.Add Int(dot.x), dot.x
    If ydict.Exists(Int(dot.y)) = False Then ydict.Add Int(dot.y), dot.y
  Next s
  
  inverter = 1   ' 交流频率控制
  xc = xdict.Count: yc = ydict.Count

  For cnt = 0 To ydict.Count - 1
    If inverter Mod 2 = 0 Then
        ssr.Sort " @shape1.Left > @shape2.Left", cnt * xc + 1, cnt * xc + xc
    Else
        ssr.Sort " @shape1.Left < @shape2.Left", cnt * xc + 1, cnt * xc + xc
    End If
    inverter = inverter + 1
  Next cnt
  
  cnt = 1
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    s.OrderToFront
    puts dot.x, dot.y, cnt: cnt = cnt + 1
  Next s
  
    ActiveDocument.EndCommandGroup
  '// 代码操作结束恢复窗口刷新
  Application.Optimization = False
  ActiveWindow.Refresh
  Application.Refresh
End Sub


























