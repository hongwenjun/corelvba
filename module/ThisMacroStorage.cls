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
    AddPluginCommand "LYVBA.CorelVBA.Start_Dimension", "CorelVBA.Start_Dimension", "CorelVBA.Start_Dimension"
    AddPluginCommand "LinesTool.lines.start", "lines.start", "lines.start"
    AddPluginCommand "ZeroBase.Hello_VBA.run", "Hello_VBA.run", "Hello_VBA.run"

    For Each Item In CommandBars
      If Item.name = "CorelVBA" Then
          creatTool = False
      End If
    Next
    If creatTool Then CommandBars.Add "CorelVBA"

    With CommandBars.Item("CorelVBA")
      .Visible = True
      Set ctl = .Controls.AddCustomButton(cdrCmdCategoryMacros, "LYVBA.CorelVBA.Start")
      ' ctl.SetIcon2 ("guid://a8e62a7a-d5d2-4a05-8d5d-e07d6bd21993")
      Set ctl2 = .Controls.AddCustomButton(cdrCmdCategoryMacros, "LYVBA.CorelVBA.Start_Dimension")
      Set ctl3 = .Controls.AddCustomButton(cdrCmdCategoryMacros, "LinesTool.lines.start")
      Set ctl4 = .Controls.AddCustomButton(cdrCmdCategoryMacros, "ZeroBase.Hello_VBA.run")
    End With
  End If
  
  refresh_Icon
ErrorHandler:
End Sub

Private Function refresh_Icon()
  With CommandBars.Item("CorelVBA")
    .Controls.Item(1).SetIcon2 ("guid://a8e62a7a-d5d2-4a05-8d5d-e07d6bd21993")
    .Controls.Item(2).SetIcon2 ("guid://b4b9632a-248b-4d80-a62d-88804e50a955")
    .Controls.Item(3).SetIcon2 ("guid://d2fdc0d9-09f8-4948-944c-4297395c05b7")
    .Controls.Item(4).SetIcon2 ("guid://1a0b1202-d0ef-4fe7-8a95-ac7617b30703")
  End With
End Function

