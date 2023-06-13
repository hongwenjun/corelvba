Attribute VB_Name = "Launcher"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "其他工具启动"   Other Tools Start  2023.6.11


'// 运行计算器
Public Function START_Calc()
    Shell "Calc"
End Function


'// 记事本打开备忘录
Public Function START_Notepad()
    cmd_line = "Notepad  C:\TSP\备忘录.txt"
    Shell cmd_line, vbNormalNoFocus
End Function


'// 打开条码阅读器
Public Function START_Barcode_ImageReader()
    cmd_line = "C:\Program Files (x86)\Softek Software\Softek Barcode Toolkit 30 Day Evaluation\bin\ImageReader.exe"
    Shell cmd_line, vbNormalNoFocus
End Function


'// 矢量化工具 Vector Magic
Public Function START_Vector_Magic()
    cmd_line = "C:\Program Files (x86)\Vector Magic\vmde.exe"
    Shell cmd_line, vbNormalNoFocus
End Function

'// waifu2x 图片放大
Public Function START_waifu2x()
    cmd_line = "C:\soft\waifu2x-gui-1.2\waifu2x-gui.exe"
    Shell cmd_line, vbNormalNoFocus
End Function

'// 开始视频录制
Public Function START_Bandicam()
    cmd_line = "C:\Program Files (x86)\Bandicam\BandicamPortable.exe"
    Shell cmd_line, vbNormalNoFocus
End Function

'// 找字体 https://www.myfonts.com/pages/whatthefont
Public Function START_whatthefont()
    Weburl "https://www.myfonts.com/pages/whatthefont"
End Function


Function Weburl(url As String)
  CorelVBA.WebHelp url
End Function
