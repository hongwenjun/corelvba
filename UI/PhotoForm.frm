VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhotoForm 
   Caption         =   "Batch Convert Img Or Export JPEG"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3855
   OleObjectBlob   =   "PhotoForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PhotoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Initialize()
    On Error Resume Next
    ComboBox1.AddItem "灰度"
    ComboBox1.AddItem "CMYK颜色"
    ComboBox1.AddItem "RGB颜色"
    ComboBox1.AddItem "黑白"
    
    ComboBox2.AddItem "300", 0
    ComboBox2.AddItem "450", 1
    ComboBox2.AddItem "600", 2
    ComboBox2.AddItem "150", 3
End Sub

Private Sub CovPhotos_Click()
    On Error Resume Next
    ActiveDocument.BeginCommandGroup
    Dim Color As String
    Dim a, b As Boolean
    Dim i, dpi As Integer
    
    a = True: b = True
    If ABox1.value = False Then a = False
    If BBox2.value = False Then b = False
    
    dpi = CInt(ComboBox2.text)
    
    Select Case ComboBox1.text
      Case Is = "灰度"
       Color = cdrGrayscaleImage
      Case Is = "CMYK颜色"
       Color = cdrCMYKColorImage
      Case Is = "RGB颜色"
       Color = cdrRGBColorImage
      Case Is = "黑白"
       Color = cdrBlackAndWhiteImage
    End Select
    
    Dim s As Shapes
    Set s = ActiveSelection.Shapes
    If s Is Nothing Then
        MsgBox "请先选中一个形状!"
        Exit Sub
    Else
        For i = 1 To s.Count
        Set s(i) = ActiveShape.ConvertToBitmapEx(Color, False, a, dpi, cdrNormalAntiAliasing, True, False, 95)
        Next i
    End If
    ActiveDocument.EndCommandGroup
End Sub

Private Sub Export_JPEG_Click()
    On Error Resume Next
    Dim d As Document
    Set d = ActiveDocument
    Dim sh As Shape, shs As Shapes
    Dim Color As String
    Set shs = ActiveSelection.Shapes
    
    dpi = CInt(ComboBox2.text)
    Select Case ComboBox1.text
    
    Case Is = "灰度"
      Color = cdrGrayscaleImage
    Case Is = "CMYK颜色"
      Color = cdrCMYKColorImage
    Case Is = "RGB颜色"
      Color = cdrRGBColorImage
    Case Is = "黑白"
      Color = cdrBlackAndWhiteImage
    End Select

    '// 导出图片精度设置，设置颜色模式
    Dim opt As New StructExportOptions
    opt.ResolutionX = dpi
    opt.ResolutionY = dpi
    opt.ImageType = Color
    
    Dim path$: path = CorelScriptTools.GetFolder
    '// 批处理导出图片
    For Each sh In shs
        ActiveDocument.ClearSelection
        sh.CreateSelection

        ' 导出图片 JPEG格式
        f = path & "\" & d.FileName & "_ID" & sh.StaticID & ".jpg"
        d.Export f, cdrJPEG, cdrSelection, opt
    Next sh
End Sub
