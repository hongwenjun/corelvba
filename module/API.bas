Attribute VB_Name = "API"
Public Function Speak_Msg(message As String)
  Speak_Help = Val(GetSetting("262235.xyz", "Settings", "SpeakHelp", "1"))
  
  If Val(Speak_Help) = 1 Then
    Dim sapi
    Set sapi = CreateObject("sapi.spvoice")
    sapi.Speak message
  Else
    ' 不说话
  End If

End Function

Public Function GetSet(s As String)
  Bleed = Val(GetSetting("262235.xyz", "Settings", "Bleed", "2.0"))
  Line_len = Val(GetSetting("262235.xyz", "Settings", "Line_len", "3.0"))
  Outline_Width = Val(GetSetting("262235.xyz", "Settings", "Outline_Width", "0.2"))
' Debug.Print Bleed, Line_len, Outline_Width

  If s = "Bleed" Then
    GetSet = Bleed
  ElseIf s = "Line_len" Then
    GetSet = Line_len
  ElseIf s = "Outline_Width" Then
    GetSet = Outline_Width
  End If
  
End Function

Public Function Create_Tolerance() As Double
  Dim text As String
  If GlobalUserData.Exists("Tolerance", 1) Then
    text = GlobalUserData("Tolerance", 1)
  End If
  text = InputBox("请输入容差值 0.1 --> 9.9", "容差值(mm)", text)
  If text = "" Then Exit Function
  GlobalUserData("Tolerance", 1) = text
  Create_Tolerance = Val(text)
End Function

Public Function Set_Space_Width() As Double
  Dim text As String
  If GlobalUserData.Exists("SpaceWidth", 1) Then
    text = GlobalUserData("SpaceWidth", 1)
  End If
  text = InputBox("请输入间隔宽度值 -99 --> 99", "设置间隔宽度(mm)", text)
  If text = "" Then Exit Function
  GlobalUserData("SpaceWidth", 1) = text
  Set_Space_Width = Val(text)
End Function

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
Public Function WriteClipBoard(ByVal s As String)
  On Error Resume Next

' VBA_WIN10(vba7) 使用PutInClipboard乱码解决办法
#If VBA7 Then
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .text = s
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
#Else
  Dim MyData As New DataObject
  MyData.SetText s
  MyData.PutInClipboard
#End If
End Function


'// 获得数组元素个数
Public Function arrlen(src As Variant) As Integer
  On Error Resume Next '空意味着 0 长度
  arrlen = (UBound(src) - LBound(src))
End Function

'// 对数组进行排序[单维]
Public Function ArraySort(src As Variant) As Variant
  Dim out As Long, I As Long, tmp As Variant
  For out = LBound(src) To UBound(src) - 1
    For I = out + 1 To UBound(src)
      If src(out) > src(I) Then
        tmp = src(I): src(I) = src(out): src(out) = tmp
      End If
    Next I
  Next out
  
  ArraySort = src
End Function

'//  把一个数组倒序
Public Function ArrayReverse(arr)
    Dim I As Integer, n As Integer
    n = UBound(arr)
    Dim p(): ReDim p(n)
    For I = 0 To n
        p(I) = arr(n - I)
    Next
    ArrayReverse = p
End Function

'// 测试数组排序
Private Function test_ArraySort()
  Dim arr As Variant, I As Integer
  arr = Array(5, 4, 3, 2, 1, 9, 999, 33)
  For I = 0 To arrlen(arr) - 1
    Debug.Print arr(I);
  Next I
  Debug.Print arrlen(arr)
  ArraySort arr
  For I = 0 To arrlen(arr) - 1
    Debug.Print arr(I);
  Next I
End Function

'// 两点连线的角度：返回角度(相对于X轴的角度)
'// p为末点，O为始点
Public Function alfaPP(p, o)
    Dim pi As Double: pi = 4 * Atn(1)
    Dim beta As Double
    If p(0) = o(0) And p(1) = o(1) Then '二点重合
        alfaPP = 0
        Exit Function
    ElseIf p(0) = o(0) And p(1) > o(1) Then
        beta = pi / 2
    ElseIf p(0) = o(0) And p(1) < o(1) Then
        beta = -pi / 2
    ElseIf p(1) = o(1) And p(0) < o(0) Then
        beta = pi
    ElseIf p(1) = o(1) And p(0) > o(0) Then
        beta = 0
    Else
        beta = Atn((p(1) - o(1)) / VBA.Abs(p(0) - o(0)))
        If p(1) > o(1) And p(0) < o(0) Then
            beta = pi - beta
        ElseIf p(1) < o(1) And p(0) < o(0) Then
            beta = -(pi + beta)
        End If
    End If
    alfaPP = beta * 180 / pi
End Function

'// 求过P点到线段AB上的垂足点(XY平面内的二维计算)
Public Function pFootInXY(p, a, B)
    If a(0) = B(0) Then
        pFootInXY = Array(a(0), p(1), 0#): Exit Function
    End If
    If a(1) = B(1) Then
        pFootInXY = Array(p(0), a(1), 0#): Exit Function
    End If
    Dim aa, bb, c, d, x, Y
    aa = (a(1) - B(1)) / (a(0) - B(0))
    bb = a(1) - aa * a(0)
    c = -(a(0) - B(0)) / (a(1) - B(1))
    d = p(1) - c * p(0)
    x = (d - bb) / (aa - c)
    Y = aa * x + bb
    pFootInXY = Array(x, Y, 0#)
End Function


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

' ************* 函数模块 ************* '
Function ExistsFile_UseFso(ByVal strPath As String) As Boolean

     Dim fso

     Set fso = CreateObject("Scripting.FileSystemObject")

     ExistsFile_UseFso = fso.FileExists(strPath)

     Set fso = Nothing

End Function

Function test()
  Dim message, sapi
  MsgBox ("Please use the headset and listen to what I have to say...")
  message = "This is a simple voice test on your Microsoft Windows."
  Set sapi = CreateObject("sapi.spvoice")
  sapi.Speak message
End Function
