Private Sub Export_JPEG_Link_Click()
    ActiveDocument.Unit = cdrCentimeter
    Dim d As Document
    Set d = ActiveDocument
    cnt = 1
    Dim sh As Shape, shs As Shapes
    Set shs = ActiveSelection.Shapes

    ' 导出图片精度设置，还可以设置颜色模式
    Dim opt As New StructExportOptions
    opt.ResolutionX = 300
    opt.ResolutionY = 300

    ' 导入图片链接设置
    Dim impflt As ImportFilter
    Dim impopt As New StructImportOptions
    With impopt
     .Mode = cdrImportFull
     .LinkBitmapExternally = True
    End With

    ' 批处理图片
    For Each sh In shs
        ActiveDocument.ClearSelection
        sh.CreateSelection

        ' 导出图片
        f = d.FilePath & "Link_" & cnt & ".jpg"
        d.Export f, cdrJPEG, cdrSelection, opt

        ' 导入图片链接
        Set impflt = ActiveLayer.ImportEx(f, cdrTIFF, impopt)
            impflt.Finish
            
        cnt = cnt + 1
    Next sh
 
End Sub
