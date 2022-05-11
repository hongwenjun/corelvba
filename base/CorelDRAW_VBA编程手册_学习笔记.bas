' 在 CorelDRAW 中，Application 对象是所有其他对象的根对象。要从进程外控制器引用 CorelDRAW 对象模型，请使用其 Application 对象。
' 尽管您可以在 VBA 中使用上述代码，但 CorelDRAW 不需要，因为如果未指定其他根对象，则默认使用 Application 对象。

Dim cdr As CorelDRAW.Application
Set cdr = CreateObject("corelDRAW.Application.14")

' Document 对象。它还包含所有 Window 对象。有关详细信息，请参阅第 57 页上的“处理文档”。
' Document 对象包含其所有 Page 对象的 Pages 集合。有关详细信息，请参阅第 67 页的“使用页面”。
' 单个页面对象包含所有图层对象的图层集合。有关详细信息，请参阅第 71 页上的“使用图层”。
' 最后，Layer 对象包含其所有 Shape 对象的 Shapes 集合。有关详细信息，请参阅第 73 页的“使用形状”。
' 要查看 CorelDRAW 对象模型的图表  CorelDRAW VBA Object Model.pdf

MsgBox "文件名:" & cdr.ActiveDocument.FileName & "目录:" & cdr.ActiveDocument.FilePath
MsgBox "目录和文件名:" & cdr.ActiveDocument.FullFileName

' 每当打开 CorelDRAW 文件时，都会在该文档的 Application 对象中创建一个新的 Document 对象。 
Document Member  ' 描述
Activate		' 激活给定文档 - 将其带到 CorelDRAW 中的最前面。 ActiveDocument 设置为引用它。
ActiveLayer		' 表示在对象管理器中设置为活动的层
ActivePage		' 表示文档中的活动页面,即在 CorelDRAW 中编辑的当前页面
AddPages AddPagesEx		' 在文档末尾添加页面
BeginCommandGroup EndCommandGroup	' 创建一个“命令组”，即在撤消列表中显示为单个项目的一系列操作
ClearSelection		' 清除文档的选择以取消选择文档中的所有形状
Close		' 关闭文档
Export		' 从文档执行简单导出
ExportEx	' 从文档执行高度可配置的导出
ExportBitmap	' 导出到具有完全控制权的位图
FileName	' 获取文件名
FilePath	' 获取文件的路径
FullFileName	' 获取文档的完整路径和文件名
GetUserArea		' 允许您通过允许用户拖动区域来为宏添加交互性
GetUserClick	' 允许您通过允许用户单击来向宏添加交互性
InsertPages InsertPagesEx	' 将页面插入到文档中的指定位置
Pages		' 提供对 Pages 集合的访问
Printout PrintSettings		' 使用文档的打印设置打印文档
PublishToPDF PDFSettings	' 将文档发布为 Adobe Acrobat Reader (PDF) 格式
ReferencePoint		' 获取/设置许多 Shape 函数使用的参考点（例如用于转换形状或获取形状位置的函数）
Save		' 使用当前文件名保存文档
SaveAs		' 将文档保存为新文件名或使用新设置
Selection	' 将选择作为形状获取
SelectionRange	' 将选择作为 ShapeRange 获取
Unit		' 设置获取测量值的函数使用的文档单位，例如与大小和位置相关的函数。 这与标尺设置使用的单位无关。
Worldscale	' 还获取/设置绘图比例。 这会改变文档中的值； 但是，它必须明确地计算到采用测量值的函数中，默认情况下使用 1:1。

' 建立新文档，导入CDR文件，另存cdr，默认保存在软件目录
Dim d As Document
Set d = CreateDocument
d.ActiveLayer.Import "R:\CDX4JX\ColorMark.cdr"
d.SaveAs "学习VBA新文档.cdr"

' 获得文件名，关闭文件
MsgBox d.FileName & "  目录: " & d.FilePath
f = d.FullFileName
d.Close

' 打开文件CDR文件，导出图片和EPS
Dim doc As Document
Set doc = OpenDocument(f)
ActiveDocument.Export "R:\学习VBA新文档.jpg", cdrJPEG
ActiveDocument.Export "R:\学习VBA新文档.eps", cdrEPS

' 调整页面大小并设置其方向
ActiveDocument.Unit = cdrMillimeter
ActivePage.SetSize 210, 297
ActivePage.Orientation = cdrLandscape

' 要设置文档的默认页面大小，请设置文档的 Pages 集合中索引为 0 的项目的值：
Dim doc As Document
Set doc = ActiveDocument
doc.Unit = cdrMillimeter
doc.Pages(0).SetSize 297, 210
' doc.MasterPage.SetSize 297, 210   ' 或以使用 Document 对象的快捷方式属性 MasterPage

' 删除页面
ActivePage.Delete
If ActiveDocument.Pages.Count > 1 Then ActivePage.Delete

' 建立图层
ActivePage.CreateLayer "刀模线图层"

' 激活图层
ActivePage.Layers("刀模线图层").Activate

' 锁定隐藏图层
ActivePage.Layers("刀模线图层").Visible = True
ActivePage.Layers("刀模线图层").Editable = False

' 创建形状 Creating shapes
形状对象表示您使用绘图工具在 CorelDRAW 文档中创建的形状。您可以创建的形状包括矩形、椭圆、曲线和文本对象。
因为每个 Shape 对象都是 Shapes 集合的成员，它是 Page 上的其中一个 Layer 对象的成员，所以用于创建新形状的方法属于 Layer 类，它们都以单词 Create 开头。

' 创建矩形 Creating rectangles
' 有两个函数用于创建新的矩形形状 - CreateRectangle 和 CreateRectangle2 - 这两个函数都返回对新 Shape 对象的引用。这两个函数的不同之处仅在于它们采用的参数。
' 例如，以下代码使用 CreateRectangle 创建一个简单的 2×1 英寸矩形(文档有误实际3x2)，该矩形位于页面底部上方 6 英寸和页面左侧 3 英寸处：
Dim sh As Shape
ActiveDocument.Unit = cdrInch
Set sh = ActiveLayer.CreateRectangle(3, 7, 6, 5)
' 参数以左、上、右、下的形式给出，它们以文档的单位进行测量（可以在创建形状之前明确设置）。

' 另一种方法 CreateRectangle2 通过指定矩形左下角的坐标及其宽度和高度来创建矩形。以下代码创建与上面相同的矩形：
Dim sh As Shape
ActiveDocument.Unit = cdrInch
Set sh = ActiveLayer.CreateRectangle2(3, 6, 2, 1)
' 提供了替代方法以简化开发解决方案；他们提供相同的功能。

' 圆角矩形也可以使用 CreateRectangle 和 CreateRectangle2 方法创建。这两个函数都有四个可选参数，用于在创建矩形时设置角的圆度，但这些值对于两个函数的含义略有不同。
' CreateRectangle 方法的四个可选参数采用 C 到 100 范围内的整数值（默认值为 0）。
' 这些值将四个角的半径定义为最短边长一半的整数百分比。以下代码重新创建了之前的 2×1 英寸矩形(文档有误实际3x2)，
' 但四个角半径设置为最短边一半的 100%、75%、50% 和 0%；换句话说，半径将是 0.5 英寸、0.375 英寸、0.25 英寸和一个尖角：
Dim sh As Shape
ActiveDocument.Unit = cdrInch
Set sh = ActiveLayer.CreateRectangle(3, 7, 6, 5, 100, 75, 50 ,0)
Set sh = ActiveLayer.CreateRectangle2(3, 7, 6, 5, 1, 1.5, 2, 0)
' CreateRectangle2 方法以相同的顺序定义角半径，除了它采用双（浮点）值，即文档单位中的半径测量值。 


' 创建椭圆  Creating ellipses
' 创建椭圆有两种方法：CreateEllipse 和 CreateEllipse2。 它们的参数不同，因此您可以根据其边界框或根据其中心点和半径创建椭圆。 这两个函数还创建弧线或部分椭圆，或线段或饼图。
' CreateEllipse 方法采用与 CreateRectangle 相同的方式定义其边界框的四个参数——换句话说，以文档的单位为左、上、右、下。 以下代码创建一个 50 毫米的圆：
Dim sh As Shape
ActiveDocument.Unit = cdrMillimeter
Set sh = ActiveLayer.CreateEllipse(75, 150, 125, 100)

' CreateEllipse2 方法基于椭圆的中心点以及水平和垂直半径创建椭圆。 （如果只给定一个半径，则创建一个圆。）以下代码创建与上述代码相同的 50 毫米圆：
Dim sh As Shape
ActiveDocument.Unit = cdrMillimeter
Set sh = ActiveLayer.CreateEllipse2(100, 125, 25)

' 要创建椭圆，请提供第二个半径。 （第一个半径是水平半径，第二个是垂直半径。）
Set sh = ActiveLayer.CreateEllipse2(100, 125, 50, 25)


' 创建曲线  Creating curves
' 曲线由其他几个对象组成：每个 Curve 有一个或多个 SubPath 成员对象，每个 SubPath 有一个或多个 Segment 对象，
' 每个 Segment 有两个 Node 对象以及两个控制点位置/角度属性。
' 您可以使用 CreateCurve 方法或 Application 对象在 CorelDRAW 中创建曲线对象。
' 使用 CreateSubPath 成员函数在 Curve 对象内创建一个新的 SubPath。这将创建曲线的第一个节点。
' 接下来，使用 AppendLineSegment 和 AppendCurveSegment 成员函数将新的线型或曲线型段附加到子路径。
' 这将添加另一个节点并设置该段的控制手柄的位置。根据需要重复此操作以构建 Lip the Curve。
' 使用 CreateCurve 成员函数在 Layer 上创建曲线形状。
' 您可以向曲线添加额外的子路径，并使用它们自己的节点构建这些嘴唇。您还可以通过将曲线的 Closed 属性设置为 True 来关闭曲线。
' 以下代码创建了一条 D 形闭合曲线：

Dim sh As Shape, spath As SubPath, crv As Curve
ActiveDocument.Unit = cdrCentimeter
Set crv = Application.CreateCurve(ActiveDocument) ' 创建曲线对象
Set spath = crv.CreateSubPath(6, 6) ' 创建一个子路径
spath.AppendLineSegment 6, 3 ' 添加短垂直线段
spath.AppendCurveSegment 3, 0, 2, 270, 2, 0 ' 下曲线
spath.AppendLineSegment 0, 0 ' 底部直边
spath.AppendLineSegment 0, 9 ' 左直边
spath.AppendLineSegment 3, 9 ' 顶部直边
spath.AppendCurveSegment 6, 6, 2, 0, 2, 90 ' 上曲线
spath.Closed = True ' 关闭曲线
Set sh = ActiveLayer.CreateCurve(crv) ' 创建曲线形状

' 创建文本对象 Creating text objects
' 文本对象是另一种类型的 Shape 对象。但是，处理文本比处理其他形状更复杂。
Dim sh As Shape
Set sh = ActiveLayer.CreateArtisticText(0, 0, "Hello World")

' 选择形状 Selecting shapes
' 要确定一个 Shape 是否被选中，你可以测试它的 Selected Boolean 属性：
Dim sh As Shape
Set sh = ActivePage.Shapes(1)
If sh.Selected = False Then sh.CreateSelection

' 您只需将其 Selected 属性设置为 True 即可将 Shape 添加到选择中；'这将选择形状而不取消选择所有其他形状。
ActivePage.Shapes(3).Selected = True

' 若要仅选择一个形状而不选择任何其他形状，请使用 CreateSelection 方法，如前面的代码中所示。
' 要取消选择所有形状，请调用文档的 ClearSelection 方法：
ActiveDocument.ClearSelection

' 要选择页面或图层上的所有形状，请使用以下代码：
ActivePage.Shapes.All.CreateSelection

' 引用 ActiveSelection 对象
Dim sel As Shape
Set sel = ActiveDocument.Selection
MsgBox "选择物件尺寸: " & sel.SizeWidth & "x" & sel.SizeHeight

' 选择遍历多个物件对象
Dim sh As Shape, shs As Shapes
Set shs = ActiveSelection.Shapes
For Each sh In shs
    MsgBox "选择物件尺寸: " & sh.SizeWidth & "x" & sh.SizeHeight
Next sh

' 实践应用: 选择物件群组,页面设置物件大小,物件页面居中
ActiveDocument.Unit = cdrMillimeter
Dim OrigSelection As ShapeRange, sh As Shape
Set OrigSelection = ActiveSelectionRange
Set sh = OrigSelection.Group

MsgBox "选择物件尺寸: " & sh.SizeWidth & "x" & sh.SizeHeight
ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter

' cdrAlignType 枚举具有以下常量:
cdrAlignLeft 1 指定左对齐
cdrAlignRight 2 指定右对齐
cdrAlignHCenter 3 指定水平居中对齐
cdrAlignTop 4 指定顶部对齐
cdrAlignBottom 8 指定底部对齐
cdrAlignVCenter 12 指定垂直居中对齐

' 添加页面框线
Dim s1 As Shape
Set s1 = ActiveLayer.CreateRectangle2(0, 0, 210, 297)
s1.Fill.ApplyNoFill
s1.OrderToFront
s1.OrderToBack
s1.Outline.SetProperties 0.04, Color:=CreateCMYKColor(0, 100, 0, 0)
s1.Move 100, 0#
s1.Move 0#, -61.8

Dim posX As Double, posY As Double
ActiveDocument.ReferencePoint = cdrBottomLeft
s1.GetPosition posX, posY
MsgBox "左下坐标: " & posX & ", " & posY

' 以下代码将活动文档中每个选定形状的右下角位置设置为 (3, 2)，以英寸为单位：
Dim sh As Shape
ActiveDocument.Unit = cdrInch
ActiveDocument.ReferencePoint = cdrBottomRight
For Each sh In ActiveSelection.Shapes
    sh.SetPosition 3, 2
Next sh

' 在页面四边放置中线
Dim sh As Shape
Set sh = ActiveDocument.Selection
sh.AlignToPage cdrAlignHCenter + cdrAlignTop
sh.Duplicate 0, 0
sh.Rotate 180
sh.AlignToPage cdrAlignHCenter + cdrAlignBottom
sh.Duplicate 0, 0
sh.Rotate 90
sh.AlignToPage cdrAlignVCenter + cdrAlignRight
sh.Duplicate 0, 0
sh.Rotate 180
sh.AlignToPage cdrAlignVCenter + cdrAlignLeft

' 在页面四角放置套准标记线
Dim sh As Shape
Set sh = ActiveDocument.Selection
sh.AlignToPage cdrAlignLeft + cdrAlignTop
sh.Duplicate 0, 0
sh.Rotate 180
sh.AlignToPage cdrAlignRight + cdrAlignBottom
sh.Duplicate 0, 0
sh.Flip cdrFlipHorizontal   ' 物件镜像
sh.AlignToPage cdrAlignLeft + cdrAlignBottom
sh.Duplicate 0, 0
sh.Rotate 180
sh.AlignToPage cdrAlignRight + cdrAlignTop


'// 获得选择物件大小信息
Sub get_all_size()
  ActiveDocument.Unit = cdrMillimeter
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.CreateTextFile("R:\size.txt", True)
  Dim sh As Shape, shs As Shapes
  Set shs = ActiveSelection.Shapes
  Dim s As String
  For Each sh In shs
    size = Trim(Str(Int(sh.SizeWidth + 0.5))) + "x" + Trim(Str(Int(sh.SizeHeight + 0.5))) + "mm"
    f.WriteLine (size)
    s = s + size + vbNewLine
  Next sh
  f.Close
  MsgBox "输出物件尺寸信息到文件" & "R:\size.txt" & vbNewLine & s
  WriteClipBoard s
End Sub

Private Function WriteClipBoard(s As String)
  On Error Resume Next
  Dim MyData As New DataObject
  MyData.SetText s
  MyData.PutInClipboard
End Function

' GetSetting 函数
' 从 Windows 注册表中 或 (Macintosh中)应用程序初始化文件中的信息的应用程序项目返回注册表项设置值。                                                                                    
Sub 加ID()
ActiveDocument.Unit = cdrMillimeter
Dim n As String
Dim s1 As Shape
Dim s As Shape
Set s = ActiveShape
If s Is Nothing Then
    MsgBox "请选择一个图形"
    Exit Sub
End If
n = vba.GetSetting("addID", "nm", "id")
If n = "" Then
    n = "1"
    vba.SaveSetting "addID", "nm", "id", "1"
Else
    n = CStr(Val(vba.GetSetting("addID", "nm", "id")) + 1)
    vba.SaveSetting "addid", "nm", "id", n
End If
Set s1 = ActiveLayer.CreateArtisticText(0, 0, "ID " & n, , , , 30)
s1.CenterX = s.CenterX
s1.CenterY = s.CenterY
End Sub

'// 查找文本选择
Sub find_id()
    Find_Text "ID"
End Sub
                                                                                                            
Public Function Find_Text(s_s As String)
  Dim s As Shape
  For Each s In ActivePage.FindShapes(, cdrTextShape)
  '  这里添加 文字判断
    If s.Text.Type = cdrArtisticText And InStr(s.Text.Story, s_s) <> 0 Then
     ' s.Text.Story = "找到 ID"
      s.AddToSelection
    End If
  Next s
End Function

'// 屏幕分辨率                                                                                                   
Public SystemX As Long
Public SystemY As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Function filePath()
  filePath = Application.Path & "GMS"
End Function

Function GetSysM(SystemX As Long, SystemY As Long)
    Dim XVal As Long, YVal As Long
    SystemX = GetSystemMetrics(0)
    SystemY = GetSystemMetrics(1)
    GetSysM = SystemX & "#" & SystemY
End Function
