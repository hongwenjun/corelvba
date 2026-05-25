VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhotoForm 
   Caption         =   "Batch Convert Or Export JPEG PDF"
   ClientHeight    =   2265
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
    
    TextBox1.text = Left(ActiveDocument.fileName, InStrRev(ActiveDocument.fileName, ".") - 1)
    
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
        For i = 1 To s.count
        Set s(i) = ActiveShape.ConvertToBitmapEx(Color, False, a, dpi, cdrNormalAntiAliasing, True, False, 95)
        Next i
    End If
    ActiveDocument.EndCommandGroup
End Sub

'// 批量导出JPEG
Private Sub Export_JPEG_Click()
  On Error GoTo ErrorHandler
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
    
    Dim path$: path = CorelScriptTools.GetFolder(d.FilePath)
    '// 批处理导出图片
    For Each sh In shs
        ActiveDocument.ClearSelection
        sh.CreateSelection

        ' 导出图片 JPEG格式
        f = path & "\" & TextBox1.text & "_ID" & sh.StaticID & ".jpg"
        d.Export f, cdrJPEG, cdrSelection, opt
    Next sh
ErrorHandler:
End Sub

'// 批量导出 PDF
Private Sub Export_PDF_Click()
  On Error GoTo ErrorHandler
    Dim d As Document
    Set d = ActiveDocument
    With d.PDFSettings
        .PublishRange = 2 ' CdrPDFVBA.pdfSelection
        .BitmapCompression = 1 ' CdrPDFVBA.pdfLZW
        .JPEGQualityFactor = 2
        .SubsetPct = 80
        .Encoding = 1 ' CdrPDFVBA.pdfBinary
        .ColorResolution = 300
        .MonoResolution = 1200
        .GrayResolution = 300
        .Startup = 0 ' CdrPDFVBA.pdfPageOnly
        .Overprints = True
        .Halftones = True
        .FountainSteps = 256
        .pdfVersion = 6 ' CdrPDFVBA.pdfVersion15
        .ColorMode = 3 ' CdrPDFVBA.pdfNative
        .ColorProfile = 1 ' CdrPDFVBA.pdfSeparationProfile
        .JP2QualityFactor = 2
        .EncryptType = 1 ' CdrPDFVBA.pdfEncryptTypeStandard
        .TextAsCurves = True ' 文字转曲
    End With

    '// 选择物件，按群组批量导出PDF
    Dim path$: path = CorelScriptTools.GetFolder(d.FilePath)
    Dim sr As ShapeRange, sh As Shape
    Set sr = ActiveSelectionRange
    
    For Each sh In sr
        ActiveDocument.ClearSelection
        sh.CreateSelection
        f = path & "\" & TextBox1.text & "_ID" & sh.StaticID & ".pdf"
        d.PublishToPDF f
    Next sh
    
ErrorHandler:
End Sub

