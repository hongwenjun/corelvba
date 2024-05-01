Attribute VB_Name = "API"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "CorelVBA���ߴ�������"   CorelVBA Tool Window Launches  2023.6.11


'// CorelDRAW ����ˢ���Ż��͹ر�
Public Function BeginOpt(Optional ByVal name As String = "Undo")
  EventsEnabled = False
  ActiveDocument.BeginCommandGroup name
' ActiveDocument.SaveSettings
  ActiveDocument.Unit = cdrMillimeter
  Optimization = True
' ActiveDocument.PreserveSelection = False
End Function

Public Function EndOpt()
' ActiveDocument.PreserveSelection = True
' ActiveDocument.RestoreSettings
  EventsEnabled = True
  Optimization = False
  EventsEnabled = True
  ActiveDocument.ReferencePoint = cdrBottomLeft
  Application.Refresh
  ActiveDocument.EndCommandGroup
End Function

Public Function Speak_Msg(message As String)
  Speak_Help = Val(GetSetting("LYVBA", "Settings", "SpeakHelp", "0"))     '// ��ͣ��������
  
  If Val(Speak_Help) = 1 Then
    Dim sapi
    Set sapi = CreateObject("sapi.spvoice")
    sapi.Speak message
  Else
    ' ��˵��
  End If

End Function

Public Function GetSet(s As String)
  Bleed = Val(GetSetting("LYVBA", "Settings", "Bleed", "2.0"))
  Line_len = Val(GetSetting("LYVBA", "Settings", "Line_len", "3.0"))
  Outline_Width = Val(GetSetting("LYVBA", "Settings", "Outline_Width", "0.2"))
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
  text = InputBox("�������ݲ�ֵ 0.1 --> 9.9", "�ݲ�ֵ(mm)", text)
  If text = "" Then Exit Function
  GlobalUserData("Tolerance", 1) = text
  Create_Tolerance = Val(text)
End Function

Public Function Set_Space_Width(Optional ByVal OnlyRead As Boolean = False) As Double
  Dim text As String
  If GlobalUserData.Exists("SpaceWidth", 1) Then
    text = GlobalUserData("SpaceWidth", 1)
    If OnlyRead Then
      Set_Space_Width = Val(text)
      Exit Function
    End If
  End If
  text = InputBox("�����������ֵ -99 --> 99", "���ü�����(mm)", text)
  If text = "" Then Exit Function
  GlobalUserData("SpaceWidth", 1) = text
  Set_Space_Width = Val(text)
End Function

'// ��ü������ı��ַ�
Public Function GetClipBoardString() As String
  On Error Resume Next
  Dim MyData As New DataObject
  GetClipBoardString = ""
  MyData.GetFromClipboard
  GetClipBoardString = MyData.GetText
  Set MyData = Nothing
End Function

'// �ı��ַ����Ƶ�������
Public Function WriteClipBoard(ByVal s As String)
  On Error Resume Next

' VBA_WIN10(vba7) ʹ��PutInClipboard�������취
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

'// ����ת�ո� ����ո񻻳�һ���ո�
Public Function Newline_to_Space(ByVal str As String) As String
  str = VBA.Replace(str, Chr(13), " ")
  str = VBA.Replace(str, Chr(9), " ")
  Do While InStr(str, "  ")
      str = VBA.Replace(str, "  ", " ")
  Loop
  Newline_to_Space = str
End Function

'// �������Ԫ�ظ���
Public Function arrlen(src As Variant) As Integer
  On Error Resume Next '����ζ�� 0 ����
  arrlen = (UBound(src) - LBound(src))
End Function

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

'//  ��һ�����鵹��
Public Function ArrayReverse(arr)
    Dim i As Integer, n As Integer
    n = UBound(arr)
    Dim P(): ReDim P(n)
    For i = 0 To n
        P(i) = arr(n - i)
    Next
    ArrayReverse = P
End Function

'// ������������
Private Function test_ArraySort()
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
End Function

'// �������ߵĽǶȣ����ؽǶ�(�����X��ĽǶ�)
'// pΪĩ�㣬OΪʼ��
Public Function alfaPP(P, o)
    Dim pi As Double: pi = 4 * Atn(1)
    Dim beta As Double
    If P(0) = o(0) And P(1) = o(1) Then '�����غ�
        alfaPP = 0
        Exit Function
    ElseIf P(0) = o(0) And P(1) > o(1) Then
        beta = pi / 2
    ElseIf P(0) = o(0) And P(1) < o(1) Then
        beta = -pi / 2
    ElseIf P(1) = o(1) And P(0) < o(0) Then
        beta = pi
    ElseIf P(1) = o(1) And P(0) > o(0) Then
        beta = 0
    Else
        beta = Atn((P(1) - o(1)) / VBA.Abs(P(0) - o(0)))
        If P(1) > o(1) And P(0) < o(0) Then
            beta = pi - beta
        ElseIf P(1) < o(1) And P(0) < o(0) Then
            beta = -(pi + beta)
        End If
    End If
    alfaPP = beta * 180 / pi
End Function

'// ���P�㵽�߶�AB�ϵĴ����(XYƽ���ڵĶ�ά����)
Public Function pFootInXY(P, a, b)
    If a(0) = b(0) Then
        pFootInXY = Array(a(0), P(1), 0#): Exit Function
    End If
    If a(1) = b(1) Then
        pFootInXY = Array(P(0), a(1), 0#): Exit Function
    End If
    Dim aa, bb, c, d, x, Y
    aa = (a(1) - b(1)) / (a(0) - b(0))
    bb = a(1) - aa * a(0)
    c = -(a(0) - b(0)) / (a(1) - b(1))
    d = P(1) - c * P(0)
    x = (d - bb) / (aa - c)
    Y = aa * x + bb
    pFootInXY = Array(x, Y, 0#)
End Function


Public Function FindAllShapes() As ShapeRange
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

' ************* ����ģ�� ************* '
Public Function ExistsFile_UseFso(ByVal strPath As String) As Boolean
     Dim fso
     Set fso = CreateObject("Scripting.FileSystemObject")
     ExistsFile_UseFso = fso.FileExists(strPath)
     Set fso = Nothing
End Function

Public Function test_sapi()
  Dim message, sapi
  MsgBox ("Please use the headset and listen to what I have to say...")
  message = "This is a simple voice test on your Microsoft Windows."
  Set sapi = CreateObject("sapi.spvoice")
  sapi.Speak message
End Function


' Public Function WebHelp(url As String)
'  Dim h As Longer, r As Long
'  h = FindWindow(vbNullString, "Toolbar")
'  r = ShellExecute(h, "", url, "", "", 1)
' End Function


