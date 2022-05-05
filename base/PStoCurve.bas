Option Explicit

Sub PStoCurve()
If ActiveShape Is Nothing Then MsgBox "Nothing selected", vbExclamation, "PStoCurve": Exit Sub
    Dim OrigSelection As ShapeRange
     Dim impflt As ImportFilter
    Dim impopt As StructImportOptions
    Set OrigSelection = ActiveSelectionRange
    Dim expflt As ExportFilter
    Dim expopt As StructExportOptions
    Set expopt = New StructExportOptions
    Dim ptt As String
    expopt.UseColorProfile = False
    ptt = Environ$("TEMP") & "\PStoCurve.ai"
 '''''''''''''''''''''' Corel X4
'    If CorelDRAW.VersionMajor = 14 Then
       
       Set expflt = ActiveDocument.ExportEx(ptt, cdrAI, cdrSelection, expopt)
        With expflt
                .Version = 6 ' FilterAILib.aiVersion6
                .TextAsCurves = True
'                .Platform = 0 ' FilterAILib.aiPC
                .ConvertSpotColors = False
                .UseColorProfile = False
                .SimulateOutlines = False
                .SimulateFills = False
                .IncludePlacedImages = False
                .IncludePreview = False
                .Finish
         End With
   
    Set impopt = New StructImportOptions
    impopt.MaintainLayers = False
    Set impflt = ActiveLayer.ImportEx(ptt, cdrAI, impopt)
    impflt.Finish
         
'    End If
 '''''''''''''''''''''''''''''''''Corel X5
'    If CorelDRAW.VersionMajor = 15 Then
''       ptt = Environ$("appdata") & "\Corel\CorelDRAW Graphics Suite X5\Draw\GMS\PStoCurve.ai"
'       Set expflt = ActiveDocument.ExportEx(ptt, cdrAI, cdrSelection, expopt)
'    With expflt
'        .Version = 2 ' FilterAILib.aiVersionCS2
'        .TextAsCurves = True
'        .PreserveTransparency = True
'        .ConvertSpotColors = False
'        .SimulateOutlines = False
'        .SimulateFills = False
'        .IncludePlacedImages = False
'        .IncludePreview = False
'        .EmbedColorProfile = False
'        .Finish
'    End With
'
'    Set impopt = CreateStructImportOptions
'
'    With impopt
'        .MaintainLayers = False
'        With .ColorConversionOptions
'            .SourceColorProfileList = "sRGB IEC61966-2.1,ISO Coated v2 (ECI),Dot Gain 15%"
'            .TargetColorProfileList = "sRGB IEC61966-2.1,ISO Coated v2 (ECI),Dot Gain 15%"
'        End With
'    End With
'    Set impflt = ActiveLayer.ImportEx(ptt, 1283, impopt)
'    impflt.Finish
'
'    End If
  '''''''''''''''''''''''''''''''''''''''''''
    
    OrigSelection.Delete
    CorelScriptTools.Kill ptt
    
End Sub

