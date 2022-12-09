Attribute VB_Name = "Launcher"
'// 运行计算器
Public Function START_Calc()
    Shell "Calc"
End Function


'// 记事本打开备忘录
Public Function START_Notepad()
    cmd_line = "Notepad  C:\TSP\备忘录.txt"
    Shell cmd_line, vbNormalNoFocus
End Function


'// 记事本打开备忘录
Public Function START_GitBash()
    cmd_line = "cmd"
    Shell cmd_line, vbNormalNoFocus
End Function


'// 记事本打开备忘录
Public Function START_Bandicam()
    cmd_line = "C:\Program Files (x86)\Bandicam\BandicamPortable.exe"
    Shell cmd_line, vbNormalNoFocus
End Function

