Attribute VB_Name = "API"
'// 获得剪贴板文本字符
Public Function GetClipBoardString() As String
    On Error Resume Next
    Dim MyData As New DataObject
    GetClipBoardString = ""
    MyData.GetFromClipboard
    GetClipBoardString = MyData.GetText
    Set MyData = Nothing
End Function

'// 文本字符复制到剪贴板
Public Function WriteClipBoard(s As String)
  On Error Resume Next
  Dim MyData As New DataObject
  MyData.SetText s
  MyData.PutInClipboard
End Function


'// 获得数组元素个数
Public Function arrlen(src As Variant) As Integer
  On Error Resume Next '空意味着 0 长度
  arrlen = (UBound(src) - LBound(src))
End Function

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

'// 测试数组排序
Private test_ArraySort()
  Dim arr As Variant, i As Integer
  arr = Array(5, 4, 3, 2, 1, 9, 999, 33)
  For i = 0 To arrlen(arr) - 1
    Debug.Print arr(i);
  Next i
  Debug.Print arrlen(arr)
  ArraySort arr
  For i = 0 To arrlen(arr) - 1
    Debug.Print arr(i);
  Next i
End Sub

Function FindAllShapes() As ShapeRange
    Dim s As Shape
    Dim srPowerClipped As New ShapeRange
    Dim sr As ShapeRange, srAll As New ShapeRange
    
    If ActiveSelection.Shapes.Count > 0 Then
        Set sr = ActiveSelection.Shapes.FindShapes()
    Else
        Set sr = ActivePage.Shapes.FindShapes()
    End If
    
    Do
        For Each s In sr.Shapes.FindShapes(Query:="!@com.powerclip.IsNull")
            srPowerClipped.AddRange s.PowerClip.Shapes.FindShapes()
        Next s
        srAll.AddRange sr
        sr.RemoveAll
        sr.AddRange srPowerClipped
        srPowerClipped.RemoveAll
    Loop Until sr.Count = 0
    
    Set FindAllShapes = srAll
End Function

