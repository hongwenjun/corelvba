Attribute VB_Name = "HotKeys"
Sub Start_SelectSame()
  '// 选择相同工具增强版
  frmSelectSame.Show 0
End Sub

Sub Start_CQL_FIND()
  '// 简单查找
  CQL_FIND_UI.Show 0
End Sub

Sub Start_Batch_Replace()
  '// 批量替换
  Replace_UI.Show 0
End Sub

Sub Start_Arrange()
  '// 开始拼版
   ArrangeForm.Show 0
End Sub

Sub Start_CutLines()
  CutLines.Draw_Lines  '// 调用角线
End Sub

Sub Start_UniteOne()
  '// 开始拼版
   UniteOne.Show 0
End Sub


Sub Start_ContainerSelect()
  ContainerForm.Show 0
End Sub

Sub Start_CardsTools()
    CardsToolsForm.Show 0
End Sub
