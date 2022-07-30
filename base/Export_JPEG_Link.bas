Private Sub cmd更新图片_Click()
    UpdateLink_Bitmap
End Sub

Private Sub Export_JPEG_Link_Click()
    Optimization = True

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
            
       ' 对齐原图，删除原图
        ActiveSelection.AlignToShape cdrAlignHCenter + cdrAlignVCenter, sh
        sh.Delete
        UpdateLink_Bitmap
        
        cnt = cnt + 1
    Next sh
    
    Optimization = False
    Application.Refresh
    ActiveWindow.Refresh
End Sub

''''''''''  显示精度优化 ''''''''''''
Private Function UpdateLink_Bitmap()

    Dim OrigSel As ShapeRange

    Set OrigSel = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrCenter
    ' 放大200%
    OrigSel.Stretch 2#, 2#

    ' 更新链接图片
    With ActiveShape.Bitmap
    If .ExternallyLinked = True Then
      .UpdateLink
    End If
    End With

    ' 缩回原大(50%)
    OrigSel.Stretch 0.5, 0.5
End Function

Private Sub FixOutdatedLinkedBitmaps_Click()
    Dim s As Shape
    Dim sr As ShapeRange
    Dim p As Page

    Optimization = True
    
    For Each p In ActiveDocument.Pages
        p.Activate
        Set sr = p.Shapes.FindShapes(, cdrBitmapShape, True)
        For Each s In sr
            If s.Bitmap.ExternallyLinked = True Then
                    s.Bitmap.UpdateLink
            End If
        Next s
    Next p
    
    Optimization = False
    Application.Refresh
    ActiveWindow.Refresh
End Sub
