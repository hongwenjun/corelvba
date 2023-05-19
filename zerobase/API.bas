Attribute VB_Name = "API"
Public Function BeginOpt(Name As String)
  EventsEnabled = False
  ActiveDocument.BeginCommandGroup Name
  ActiveDocument.SaveSettings
  ActiveDocument.unit = cdrMillimeter
  Optimization = True
' ActiveDocument.PreserveSelection = False
End Function

Public Function EndOpt()
' ActiveDocument.PreserveSelection = True
  ActiveDocument.RestoreSettings
  EventsEnabled = True
  Optimization = False
  EventsEnabled = True
  Application.Refresh
  ActiveDocument.EndCommandGroup
End Function

Public Function Create_Tolerance() As Double
  Dim text As String
  If GlobalUserData.Exists("Tolerance", 1) Then
    text = GlobalUserData("Tolerance", 1)
  End If
  text = InputBox("请输入容差值 0.1 --> 99.9", "容差值(mm)", text)
  If text = "" Then Exit Function
  GlobalUserData("Tolerance", 1) = text
  Create_Tolerance = Val(text)
End Function
