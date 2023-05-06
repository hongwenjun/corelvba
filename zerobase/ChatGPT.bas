Attribute VB_Name = "ChatGPT"
Private Type Coordinate
    x As Double
    y As Double
End Type

Sub Z������()

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

Sub U������()

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
  
  inverter = 1   ' ����Ƶ�ʿ���
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


Sub ��������()   ' �ֵ�ʹ�ü�������

  ActiveDocument.Unit = cdrMillimeter
  Set xdict = CreateObject("Scripting.dictionary")
  Set ydict = CreateObject("Scripting.dictionary")
  Dim dot As Coordinate, Offset As Coordinate
  Dim s As Shape, ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  
  ' ��ǰѡ������ķ�Χ�߽�
  set_lx = ssr.LeftX: set_rx = ssr.RightX
  set_by = ssr.BottomY: set_ty = ssr.TopY
  ssr(1).GetSize Offset.x, Offset.y
  ' ��ǰѡ����� ShapeRange ��������
  ssr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
  
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    If xdict.Exists(Int(dot.x)) = False Then xdict.Add Int(dot.x), dot.x
    If ydict.Exists(Int(dot.y)) = False Then ydict.Add Int(dot.y), dot.y
  Next s
  
'  MsgBox "�ֵ�ʹ�ü�������:" & xdict.Count & ydict.Count
  Dim cnt As Long: cnt = 1
  
  ' �����ֵ䣬���
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

'// �������������[��ά]
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
    ' ����һ���յ�Dictionary
    Dim myDict As Object
    Set myDict = CreateObject("Scripting.Dictionary")
    
    ' ��Dictionary����Ӽ�ֵ��
    myDict.Add "orange", 4
    myDict.Add "banana", 2
    myDict.Add "apple", 3
    
    ' ���ʼ�ֵ��
    Debug.Print "The value of 'apple' is " & myDict("apple")
    
    ' ����Dictionary�е����м�ֵ��
    Dim key As Variant
    For Each key In myDict.keys
        Debug.Print key & " : " & myDict(key)
    Next key
    
    ' ���ĳ�����Ƿ����
    If myDict.Exists("orange") Then
        Debug.Print "The key 'orange' exists"
    End If
    
    ' ɾ��ĳ����ֵ��
    myDict.Remove "banana"
    
    ' ���Dictionary
    myDict.RemoveAll
End Sub

Sub tongjiʹ���ֵ�ͳ��()

  Dim s As Shape
  Dim sr As ShapeRange
  
  Set sr = ActiveSelection.Shapes.FindShapes(Query:="@name='wk-y���'")
  
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
  
  str = "   ��   ��" & vbTab & vbTab & vbTab & "����" & vbNewLine

  a = d.keys: b = d.items
  For i = 0 To d.Count - 1
    str = str & a(i) & vbTab & vbTab & b(i) & "��" & vbNewLine
  Next

  ' ����Dictionary�е����м�ֵ��
  Dim key As Variant
  For Each key In d.keys
      Debug.Print key & " : " & d(key)
  Next key

  Debug.Print str
End Sub




Sub ��ʽU������()
  Application.Optimization = True
  ActiveDocument.BeginCommandGroup  'һ������'

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
  
  inverter = 1   ' ����Ƶ�ʿ���
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
  '// ������������ָ�����ˢ��
  Application.Optimization = False
  ActiveWindow.Refresh
  Application.Refresh
End Sub


























