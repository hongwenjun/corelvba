VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_start()
  On Error GoTo ErrorHandler
  Dim creatTool As Boolean: creatTool = True
  StartButton = Val(GetSetting("LYVBA", "Settings", "StartButton", "0"))
  If StartButton = 1 Then creatTool = False
  
  If creatTool Then
    SaveSetting "LYVBA", "Settings", "StartButton", "1"
    AddPluginCommand "LYVBA.CorelVBA.Start", "CorelVBA.Start", "CorelVBA.Start"
    
    For Each Item In CommandBars
      If Item.Name = "CorelVBA" Then
          creatTool = False
      End If
    Next
    If creatTool Then CommandBars.Add "CorelVBA"
    
    With CommandBars.Item("CorelVBA")
      .Visible = True
      Set ctl = .Controls.AddCustomButton(cdrCmdCategoryMacros, "LYVBA.CorelVBA.Start")
      ' ctl.SetIcon2 ("guid://46327bd4-8bad-41c5-aba1-efa770b8e2c8")
      ctl.SetCustomIcon "C:\TSP\LYVBA.ico"
    End With
  End If
  
ErrorHandler:
End Sub
