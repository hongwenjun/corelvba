' 实用小脚本:  选择遍历多个物件对象，按序号导出 JPEG
ActiveDocument.Unit = cdrCentimeter
Dim d As Document
Set d = ActiveDocument
cnt = 1
Dim sh As Shape, shs As Shapes
Set shs = ActiveSelection.Shapes

For Each sh In shs
    ActiveDocument.ClearSelection
    sh.CreateSelection
    MsgBox "选择物件尺寸: " & sh.SizeWidth & "x" & sh.SizeHeight
    Size = Str(Int(sh.SizeWidth + 0.5)) + "x" + Str(Int(sh.SizeHeight + 0.5))
    f = "R:\" + Str(cnt) + "_尺寸" + Size + ".jpg"
    ' 可以把获得的尺寸取整数，写到文件名中，或者把尺寸信息写到图片中
    d.Export f, cdrJPEG, cdrSelection
    cnt = cnt + 1
Next sh
