Attribute VB_Name = "Module1"
Sub 统计文本()
  Dim s As Shape, sr As ShapeRange
  Set sr = ActiveSelectionRange
   
  Dim d As Variant, str As String
  Set d = CreateObject("Scripting.dictionary")
  
   For Each s In sr
    If s.Type = cdrTextShape Then
      str = s.text.Story.text
      If d.Exists(str) = True Then
        d.Item(str) = d.Item(str) + 1
      Else
        d.Add str, 1
      End If
    End If
  Next s
  

  str = "文  本" & vbTab & vbTab & "数量" & vbNewLine
  a = d.keys: b = d.items
  For i = 0 To d.Count - 1
    str = str & a(i) & vbTab & b(i) & "条" & vbNewLine
  Next
  str = str & "合计总量:" & vbTab & vbTab & d.Count & "条" & vbNewLine

  Debug.Print str
  
  Dim s1 As Shape
  x = sr.FirstShape.LeftX - 100
  y = sr.FirstShape.TopY
  Set s1 = ActiveLayer.CreateParagraphText(x, y, x + 90, y - 150, str, Font:="华文中宋")
End Sub

