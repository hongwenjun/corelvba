' CorelDRAW VBA自动化调用菜单命令的例子
'过程名称：打印全部菜单ID（用于自动化调用）

Public Sub listMenuItemIDs()
  On Error Resume Next
  Dim cmdbar, ctl
  For Each cmdbar In FrameWork.CommandBars
    Debug.Print cmdbar & "工具栏下面的菜单项："
    For Each ctl In cmdbar.Controls
     Debug.Print vbTab & ctl.id & " -> " & ctl.Caption
    Next
  Next
End Sub

' 拿到ID后，就可以通过自动化框架提供的方法来调用指定的菜单，参考代码如下：

Sub 自动化调用()
  Dim OrigSelection As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  OrigSelection.Application.FrameWork.Automation.InvokeItem "e6644135-9dab-4935-8ab9-fc85527810ca"
End Sub

Public Sub listMenuItemIDs_savefile()
  On Error Resume Next
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set F = fs.CreateTextFile("D:\MenuItemIDs.txt", True)

  Dim cmdbar, ctl
  For Each cmdbar In FrameWork.CommandBars
    F.WriteLine cmdbar & "工具栏下面的菜单项："
    For Each ctl In cmdbar.Controls
      F.WriteLine vbTab & ctl.id & " -> " & ctl.Caption
    Next
  Next

  F.Close
End Sub
