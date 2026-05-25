Attribute VB_Name = "CardsTools"
Public Function MakeRectangleToPowerClip(w As Double, h As Double)
    Dim ssr As ShapeRange, s As Shape
    Dim cnt As Integer
    Dim i As Integer

    Set ssr = ActiveSelectionRange
    cnt = ssr.count

    If cnt = 0 Then Exit Function

    Dim jxsr As New ShapeRange

    ' 为每个选择的对象创建一个矩形
    For i = 1 To cnt
        Set s = Rectangle(w, h)
        jxsr.Add s
    Next i

    sr_Arrangement jxsr, 30
    jxsr.SetOutlineProperties 0#   '// 没轮廓
    jxsr.Move 0, jxsr.SizeHeight + 30

    '// 批量调整尺寸和居中对齐
    For i = 1 To cnt
        SetShapeSize ssr(i), w, h
        ssr(i).CenterX = jxsr(i).CenterX
        ssr(i).CenterY = jxsr(i).CenterY
        jxsr(i).name = "powerclip_ok"
        ssr(i).AddToPowerClip jxsr(i)
    Next i

    jxsr.CreateSelection

End Function

'// 功能：解包当前选择的所有 PowerClip 对象
Public Function PowerClip_ExtractShapes()
    Dim s As Shape
    Dim pwc As PowerClip  ' 存储 PowerClip 对象

    For Each s In ActiveSelectionRange
        Set pwc = Nothing  ' 每次循环重置变量
        ' 错误处理：尝试获取形状的 PowerClip 属性
        On Error Resume Next
        Set pwc = s.PowerClip  ' 如果 s 不是 PowerClip，这里会出错
        On Error GoTo 0        ' 恢复正常错误处理
        ' 检查是否成功获取到 PowerClip 对象
        If Not pwc Is Nothing Then
            '//  s.CreateSelection     ' 选中当前 PowerClip 容器
            pwc.ExtractShapes    ' 解包：将内容从容器中取出
        End If
    Next s
End Function

'// 建立矩形 Width  x Height 单位 mm
Private Function Rectangle(width As Double, Height As Double) As Shape
    Dim s As Shape
    Set s = ActiveLayer.CreateRectangle(0, 0, 0 + width, 0 - Height)
    s.Fill.ApplyNoFill
    Set Rectangle = s
End Function

'// 简洁版本：确保一边正好等于目标尺寸，另一边不小于指定最小值
Private Function SetShapeSize(s As Shape, w As Double, h As Double)
    If s Is Nothing Then Exit Function

    Dim originalWidth As Double
    Dim originalHeight As Double
    Dim ratio As Double

    originalWidth = s.SizeWidth
    originalHeight = s.SizeHeight
    ratio = originalWidth / originalHeight

    Dim newWidth As Double
    Dim newHeight As Double

    '// 尝试方案1：宽固定为85，计算高
    newWidth = w
    newHeight = w / ratio

    '// 如果高太小（小于45），则改用方案2：高固定为54
    If newHeight < h Then
        newHeight = h
        newWidth = h * ratio

        '// 如果宽太小（小于85），则按比例放大直到宽等于85
        If newWidth < w Then
            newWidth = w
            newHeight = w / ratio
        End If
    End If

    '// 应用新尺寸
    s.SetSize newWidth, newHeight
End Function

Private Function sr_Arrangement(ssr As ShapeRange, Space_Width As Double)
    Dim s As Shape
    Dim cnt As Integer
    cnt = 1

    ActiveDocument.ReferencePoint = cdrTopLeft
    For Each s In ssr
        ActiveDocument.ReferencePoint = cdrTopLeft + cdrBottomTop
        If cnt > 1 Then s.SetPosition ssr(cnt - 1).RightX + Space_Width, ssr(cnt - 1).topY
        cnt = cnt + 1
    Next s

End Function

Public Function Save_CdrX4_File(CDRX4_FileName As String)
    Dim SaveOptions As StructSaveAsOptions
    Set SaveOptions = CreateStructSaveAsOptions
    With SaveOptions
        .EmbedVBAProject = True
        .Filter = cdrCDR
        .IncludeCMXData = False
        .Range = cdrAllPages
        .EmbedICCProfile = False
        .Version = cdrVersion14
    End With

    ActiveDocument.SaveAs CDRX4_FileName, SaveOptions
End Function

Private Function GetImageFiles(folderPath As String, fileList As Collection)
    Dim fileName As String
    Dim ext As String

    ' 确保路径以反斜杠结尾
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' 使用Dir函数获取第一个文件
    fileName = Dir(folderPath & "*.*")

    ' 遍历所有文件
    Do While fileName <> ""
        ' 获取文件扩展名
        ext = LCase(Right(fileName, Len(fileName) - InStrRev(fileName, ".")))

        ' 检查是否是图片文件
        Select Case ext
        Case "jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff"
            fileList.Add folderPath & fileName
        End Select

        ' 获取下一个文件
        fileName = Dir
    Loop

End Function

Private Function MoveImageFile_Name(Optional ByVal sourceFileName As String, Optional ByVal destFileName As String) As Boolean
    On Error Resume Next
    
    ' 如果目标文件存在，直接添加后缀
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(destFileName) Then
        Dim i As Long
        i = 1
        Do While fso.FileExists(destFileName)
            destFileName = Replace(destFileName, ".", "_" & i & ".")
            i = i + 1
        Loop
    End If
    
    ' 移动文件
    Name sourceFileName As destFileName
    
    MoveImageFile_Name = (err.Number = 0)
    On Error GoTo 0
End Function

Public Function Import_Images()
    Dim folderPath As String
    Dim backtupPath As String
    Dim fileList As New Collection
    Dim sr As New ShapeRange

    ' 设置文件夹路径
    folderPath = "D:\Cards"
    backtupPath = "D:\Cards\BACKUP"
    Call GetImageFiles(folderPath, fileList)

    ' 批量导入图片
    Dim f As Variant
    For Each f In fileList
        ActiveDocument.ClearSelection
        ActiveLayer.Import f
        sr.Add ActiveSelection
    Next f
    sr.CreateSelection

    ' 移动图片到备份文件夹
    Dim sourceFileName As String
    Dim dstFileName As String
    For Each f In fileList
      sourceFileName = f
      desFileName = Replace(sourceFileName, "D:\Cards", "D:\Cards\BACKUP")
      MoveImageFile_Name sourceFileName, desFileName
    Next f
    
End Function

Public Function Images2NewDoc()
    Dim doc As Document
    Set doc = CreateDocument()
    doc.Unit = cdrMillimeter

    Call Import_Images
End Function
