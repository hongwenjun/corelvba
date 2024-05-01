Attribute VB_Name = "快捷键"
Sub 木头人群组()
  autogroup("group", 1).CreateSelection
End Sub

Sub 角转平()
  Tools.角度转平
End Sub

Sub 对象交换()
  Tools.交换对象
End Sub

Sub 安全线()
    Tools.guideangle ActiveSelectionRange, 0#   ' 右键 0距离贴紧
End Sub


