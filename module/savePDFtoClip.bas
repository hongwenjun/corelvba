Attribute VB_Name = "savePDFtoClip"
#If VBA7 Then
  Private Declare PtrSafe Function vbadll Lib "lycpg64.cpg" (ByVal code As Long, ByVal x As Double) As Long
#Else
  Private Declare Function vbadll Lib "lycpg32.cpg" (ByVal code As Long, ByVal x As Double) As Long
#End If

Sub CorelDRAW_CopyPDF()
'//  savePDFtoClip.CdrCopyToAI
'// VBA调用CPG_CDR复制物件到AI()
 ret = vbadll(2, 0)
 
End Sub

Sub CorelDRAW_PastePDF()
'//  savePDFtoClip.AICopyToCdr
'// AI复制物件到CDR()
 ret = vbadll(1, 0)
End Sub

Private Function GetTempFile(ByVal sExtension As String) As String
    GetTempFile = CorelScriptTools.GetTempFolder() & "CDR2AI" & "." & sExtension
End Function

Public Function CdrCopyToAI()
  On Error GoTo ErrorHandler
  sTempFilePDF = GetTempFile("pdf")
  
    With ActiveDocument.PDFSettings
        .PublishRange = 2 ' CdrPDFVBA.pdfSelection
        .BitmapCompression = 1 ' CdrPDFVBA.pdfLZW
        .JPEGQualityFactor = 2
        .EmbedFonts = True
        .EmbedBaseFonts = True
        .TrueTypeToType1 = True
        .SubsetFonts = False
        .SubsetPct = 80
        .CompressText = True
        .Encoding = 1 ' CdrPDFVBA.pdfBinary
        .ColorResolution = 300
        .MonoResolution = 1200
        .GrayResolution = 300
        .Hyperlinks = True
        .Bookmarks = True
        .Thumbnails = True
        .Startup = 0 ' CdrPDFVBA.pdfPageOnly
        .Overprints = True
        .Halftones = True
        .FountainSteps = 256
        .EPSAs = 0 ' CdrPDFVBA.pdfPostscript
        .pdfVersion = 6 ' CdrPDFVBA.pdfVersion15
        .ColorMode = 3 ' CdrPDFVBA.pdfNative
        .ColorProfile = 1 ' CdrPDFVBA.pdfSeparationProfile
        .JP2QualityFactor = 2
        .TextExportMode = 0 ' CdrPDFVBA.pdfTextAsUnicode
        .PrintPermissions = 0 ' CdrPDFVBA.pdfPrintPermissionNone
        .EditPermissions = 0 ' CdrPDFVBA.pdfEditPermissionNone
        .EncryptType = 1 ' CdrPDFVBA.pdfEncryptTypeStandard
    End With
  
  ActiveDocument.PublishToPDF sTempFilePDF
  
  '// 调用 pdf2clip.exe 把PDF文件加载到剪贴板,  命令行按实际文件夹填写路径
  
  cmd_line = "C:\TSP\pdf2clip.exe  " & sTempFilePDF
  ret = Shell(cmd_line, vbHide)
  
ErrorHandler:
End Function

Public Function AICopyToCdr()
  On Error GoTo ErrorHandler
  sTempFilePDF = GetTempFile("pdf")
  '// 调用 clip2pdf.exe 把读取剪贴板保存成PDF
  cmd_line = "C:\TSP\clip2pdf.exe  " & sTempFilePDF

  Dim ret As Long
  ret = Shell(cmd_line, vbHide)
  
  '// 暂停 1 秒 让Shell 调用exe程序完成结果
  Dim startTime As Variant
  startTime = Now
  Do While (Now - startTime) < TimeSerial(0, 0, 1)
    DoEvents
  Loop

Dim impopt As StructImportOptions
Set impopt = CreateStructImportOptions
impopt.MaintainLayers = True

Dim impflt As ImportFilter
Set impflt = ActiveLayer.ImportEx(sTempFilePDF, cdrAI9, impopt)
impflt.Finish

ErrorHandler:
End Function

