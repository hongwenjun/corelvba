VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectSame 
   Caption         =   "Similar Selection Plus"
   ClientHeight    =   5745
   ClientLeft      =   495
   ClientTop       =   5895
   ClientWidth     =   3255
   OleObjectBlob   =   "frmSelectSame.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmSelectSame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "相似选择-魔改版 蘭雅"   frmSelectSame   2023.6.12

Option Explicit
'需要显式声明所有变量。 这可以防止无意中使用缓慢的“Variant”类型变量，这些变量在特定类型未知时使用。
'Requires explicit declaration of all variables. This protects against inadvertent use of the slow 'Variant' type variables which are used when the specific type is unknown.

Public ssreg As ShapeRange

Private Const TOOLNAME As String = "VBA_SelectSame"
Private Const SECTION As String = "Options"

Private Sub UserForm_Initialize()
  LNG_CODE = Val(GetSetting("LYVBA", "Settings", "I18N_LNG", "1033"))
  Init_Translations Me, LNG_CODE
  Me.Caption = i18n("Similar Selection Plus", LNG_CODE)
End Sub

Private Sub btnSelect_Click()
    If 0 = ActiveSelectionRange.Count Then Exit Sub
    On Error GoTo ErrorHandler
    
    Dim fLeft As Double, fTop As Double
    fLeft = frmSelectSame.Left
    fTop = frmSelectSame.Top
    SaveSetting "SelectSame", "Preferences", "form_left", fLeft
    SaveSetting "SelectSame", "Preferences", "form_top", fTop
    
    '// 区域范围选择，需要关闭刷新优化
    If OptBt.value = False Then
      API.BeginOpt
    Else
      add_ssreg
    End If
    
    If (chkFill = False And chkOutline = False And chkOutlineColor = False And _
      chkOutlineLength = False And chkSize = False And chkWHratio = False And _
      chkType = False And chkNodes = False And chkSegments = False And _
      chkPaths = False And chkFontName = False And chkFontSize = False And chkShapeName = False) Then
        MsgBox "请至少选择一个选项", vbCritical, "Select Same"
        GoTo ErrorHandler
    End If


'// "ME"是一个VBA保留字，返回对当前代码所在窗体（或类模块）的引用。 chk... 函数返回同名复选按钮的当前值。
'// "ME" is a VBA reserved word, returning a reference to the form (or class module) in which the current code is located.
'//  The chk... functions return the current Value of the check buttons of the same name.
    With Me
      .SelectAllSimilar .chkFill, .chkOutline, .chkOutlineColor, .chkOutlineLength, _
      .chkSize, .chkWHratio, .chkType, .chkNodes, .chkSegments, .chkPaths, _
      .OptDoc, .Optpage, .Optlayer, .chkInGroups, .chkColorMark, .chkIndiv, _
      .chkFontName, .chkFontSize, .chkShapeName
    End With
    
    API.EndOpt
    
Exit Sub
ErrorHandler:
  Application.Optimization = False
End Sub

Sub SelectAllSimilar(Optional CheckFill As Boolean = True, _
                    Optional CheckOutline As Boolean = True, _
                    Optional CheckOutlineColor As Boolean = True, _
                    Optional CheckOutlineLength As Boolean = True, _
                    Optional CheckSize As Boolean = False, _
                    Optional CheckWHratio As Boolean = False, _
                    Optional CheckType As Boolean = True, _
                    Optional CountNodes As Boolean = False, _
                    Optional CountSegments As Boolean = False, _
                    Optional CountPaths As Boolean = False, _
                    Optional WithinDoc As Boolean = False, _
                    Optional WithinPage As Boolean = True, _
                    Optional WithinLayer As Boolean = False, _
                    Optional WithinGroups As Boolean = True, _
                    Optional CheckColorMark As Boolean = False, _
                    Optional CheckIndiv As Boolean = True, _
                    Optional CheckFontName As Boolean = False, _
                    Optional CheckFontSize As Boolean = False, _
                    Optional CheckShapeName As Boolean = False)
                    
    'Object variables.              Reference to:
    Dim shpsSelected As Shapes          'selected shapes,
    Dim shpsToTest As Shapes            'full set of shapes to be tested,  ' 待测形状全部集合
    Dim pagesr As ShapeRange           'pages shapes collection,
    Dim docsr As New ShapeRange
    Dim shpModel As Shape               'a pre-selected shape,
    Dim shpToMatch As Shape             'a shape to be matched,
    'Dim oScript As Object               'CorelScript object,
    Dim clnModelShapes As Collection    'our list of pre-selected shapes,  '定义源对象集合
    Dim clnSubShapes As Collection      'our list of shapes inside a group. '定义群组内的目标对象
    Dim P As Page, p1 As Page           '文档中查找使用
    Dim shr As ShapeRange, sr As New ShapeRange
    Dim i As Integer  ' '文档中循环查找计数使用
    Dim fsn As Shape  '// 扩展功能: 字体字号标记名检测源对象

    On Error GoTo NothingSelected       'Get a reference to any
    Set shr = ActiveSelectionRange
    Set shpsSelected = ActiveDocument.Selection.Shapes
'    On Error GoTo 0                     'pre-selected shapes. 将文档中当前选中的范围作为源对象
    
    If shpsSelected.Count > 0 Then          'Gather the pre-selected shapes
        Set clnModelShapes = New Collection 'into a new collection for
        For Each shpModel In shpsSelected   'simple processing. 建立源对象集合
           clnModelShapes.Add shpModel
        Next
        

        '// 魔改分支 字体-字号-标记名
        If CheckFontName Or CheckFontSize Or CheckShapeName Then
          Set fsn = shr(1)
        End If

        '===================================
        ' TurnOptimizations cdrOptimizationOn
        '===================================
       
        If WithinPage Then

          If OptBt.value = True Then
            Set shpsToTest = ssreg.Shapes
            OptBt.value = 0
            API.BeginOpt
          Else
            Set shpsToTest = ActivePage.Shapes
          End If
                                            'Ensure that "Edit across layers"
                                            'is ON. Otherwise, selecting
'            Set oScript = CorelScript       'across layers, followed by
'            oScript.SetMultiLayer True      'grouping, can flatten all
'            Set oScript = Nothing           'layers into one. 选中表示将对当前页面的所有对象与源对象进行匹配，否则只匹配当前图层的对象
 
            'Replace the above with this line, CoreScript is not longer support X7+
            ActiveDocument.EditAcrossLayers = True
        End If
        If WithinLayer Then
            Set shpsToTest = ActivePage.ActiveLayer.Shapes
        End If
        
        If WithinDoc Then '在当前文档查找，将当前页面相应的对象加入到待比较范围
            For i = 1 To ActiveDocument.Pages.Count
                ActiveDocument.Pages(i).Activate
                Set p1 = ActiveDocument.Pages(i)
                Set pagesr = ActivePage.SelectShapesFromRectangle(0, p1.CenterY * 2, p1.CenterX * 2, 0, False).Shapes.all
                Debug.Print p1.CenterY * 2 & p1.CenterX * 2
                docsr.AddRange pagesr '各页面依次查找，相应的对象加入到待比较范围
                
            Next i
            Set shpsToTest = docsr.Shapes
'            MsgBox "共有待比较对象 " & shpsToTest.Count & " 个"
            Label13.Caption = "共有待比较对象 " & shpsToTest.Count & " 个"
            'p1.Activate
        End If
        
        If WithinGroups Then                'Check through flattened list.
            Set clnSubShapes = FlatShapeList(shpsToTest)
            '=======
            For Each shpToMatch In clnSubShapes
                If Not shpToMatch.Selected Then 'If the shape is not yet selected,
                
                   '====================     'check the models for a match.
                    For Each shpModel In clnModelShapes
                        If ShapesMatch(shpToMatch, shpModel, CheckFill, _
                                CheckOutline, CheckOutlineColor, CheckOutlineLength, CheckSize, CheckWHratio, _
                                CheckType, CountNodes, CountSegments, CountPaths, CheckIndiv) Then
                            'shpToMatch.AddToSelection
                            sr.Add shpToMatch
                            Exit For        'If a match has now been found,
                        End If              'we can skip any remaining models.
                    Next
                   '=====================
                   
                End If
            Next
            '=======
        Else                                'Check through top-level list.
            For Each shpToMatch In shpsToTest
                If Not shpToMatch.Selected Then 'If the shape is not yet selected,
                                            'check the models for a match.
                    For Each shpModel In clnModelShapes
                        If ShapesMatch(shpToMatch, shpModel, CheckFill, _
                                CheckOutline, CheckOutlineColor, CheckOutlineLength, CheckSize, CheckWHratio, _
                                CheckType, CountNodes, CountSegments, CountPaths, CheckIndiv) Then
                               'shpToMatch.AddToSelection
                            sr.Add shpToMatch
                            Exit For        'If a match has now been found,
                        End If              'we can skip any remaining models.
                    Next
                    
                End If
            Next
        End If
            
        '===================================
       ' TurnOptimizations cdrOptimizationOff
        'CorelScript.RedrawScreen
        '===================================
        'sr.Add ActiveDocument.Selection
        If CheckColorMark And sr.Count > 0 Then sr.SetOutlineProperties , , CreateCMYKColor(0, 100, 0, 0) '轮廓线上色
        sr.AddRange shr
    
        '// 魔改分支 字体-字号-标记名
        If CheckFontName Or CheckFontSize Or CheckShapeName Then
          If CheckFontName Then ShapesMatch_Font_Name fsn, sr, "FontName"
          If CheckFontSize Then ShapesMatch_Font_Name fsn, sr, "FontSize"
          If CheckShapeName Then ShapesMatch_Font_Name fsn, sr, "ShapeName"
        End If
        
       sr.CreateSelection
        '// 显示找到对象
        Label13.Caption = "共找到 " & sr.Count & " 个对象"
    End If
    
    Set clnModelShapes = Nothing               'Release the memory allocated
    Set shpsToTest = Nothing
    Exit Sub
NothingSelected:
End Sub

'// 添加区域选择分支
Private Function add_ssreg()
    Dim ssr As ShapeRange, shr As ShapeRange
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim Shift As Long
    Dim b As Boolean
    Set shr = ActiveSelectionRange
    b = ActiveDocument.GetUserArea(x1, y1, x2, y2, Shift, 10, False, cdrCursorWeldSingle)
    If Not b Then
      Set ssreg = ActivePage.SelectShapesFromRectangle(x1, y1, x2, y2, True).Shapes.all
    End If
    ActiveDocument.ClearSelection
    shr.CreateSelection
End Function

'// 魔改分支 字体-字号-标记名  检查匹配
Private Function ShapesMatch_Font_Name(ByVal fsn As Shape, sr As ShapeRange, Check_Case As String)
  Dim xz As String, sh_name As String, strFontName As String
  Dim FontSize As Double
  Dim srText As ShapeRange
  Set srText = sr.Shapes.FindShapes(Type:=cdrTextShape)
      
  Select Case Check_Case
  Case "FontName"
    If fsn.Type = cdrTextShape Then
      strFontName = fsn.text.Story.Font
      Set sr = srText.Shapes.FindShapes(Query:="@type ='text:artistic' or @type ='text:paragraph' and @com.text.story.font = '" & strFontName & "'")
    End If
    
  Case "FontSize"
    If fsn.Type = cdrTextShape Then
      FontSize = fsn.text.Story.size
      Set sr = srText.Shapes.FindShapes(Query:="@type ='text:artistic' or @type ='text:paragraph' and (@com.text.story.size - " & FontSize & ").abs() < 0.1 ")
    End If
    
  Case "ShapeName"
    sh_name = fsn.name
      Set sr = sr.Shapes.FindShapes(Query:="@name ='" & sh_name & "'")
  End Select
End Function


Private Function ShapesMatch(shpShape As Shape, shpModel As Shape, _
                    Optional CheckFill As Boolean = True, _
                    Optional CheckOutline As Boolean = True, _
                    Optional CheckOutlineColor As Boolean = True, _
                    Optional CheckOutlineLength As Boolean = True, _
                    Optional CheckSize As Boolean = False, _
                    Optional CheckWHratio As Boolean = False, _
                    Optional CheckType As Boolean = True, _
                    Optional CountNodes As Boolean = False, _
                    Optional CountSegments As Boolean = False, _
                    Optional CountPaths As Boolean = False, _
                    Optional CheckIndiv As Boolean = False) As Boolean
    
    'Sizes "match" if they differ by less than one per cent
    Dim ToleranceSize As Double     '面积大小允许波动
    ToleranceSize = Me.TextBox1 / 100  '面积大小允许波动,以百分比为单位
    
    Dim ToleranceLength As Double   '线长允许波动
    ToleranceLength = Me.TextBox2 / 100 '长度允许波动,以百分比为单位
    
    Dim ToleranceNodesCount As Long  '节点数量允许波动,以 点 单位
    ToleranceNodesCount = Me.TextBox3 '节点数量允许波动,以 点 单位
    
    Dim ToleranceSubPathsCount As Long  '子路径 子线段 允许波动,以 条 为单位
    ToleranceSubPathsCount = Me.TextBox4 '子路径 子线段 允许波动,以 条 为单位
    
    Dim ToleranceWHratio As Double  '长宽比 允许波动,以 百分比 为单位
    ToleranceWHratio = Me.TextBox5  '长宽比 允许波动,以 百分比 为单位
    
    Dim ToleranceSegmentsCount As Long  '线段数 允许波动,以 个 为单位
    ToleranceSegmentsCount = Me.TextBox6 '线段数 允许波动,以 个 为单位
        
    'Object Variables.        'Reference to:
    Dim clrModel As Color           'color features of model shape,
    Dim clrShape As Color           'color features of shape to be tested
    Dim fillModel As Fill           'fill style of model shape,
    Dim outlnModel As Outline       'outline style of model shape,
    Dim crvModel As Curve           'Bezier curve of model shape,
    Dim crvShape As Curve           'Bezier curve of shape to be tested,
    Dim fntModel As StructFontProperties  'font properties of model text shape,
    Dim trgModel As text            'general text properties of model shape.
    Dim spath As SubPath, opath As SubPath
    Dim j As Integer
    
    'Simple Variables.              Storage of:
    Dim dblWidth As Double              'width of a shape,
    Dim dblHeight As Double             'height of a shape,
    Dim lngShapeType As cdrShapeType    'code for type of shape to be tested,
    Dim lngModelType As cdrShapeType    'code for the type of a model shape,
    Dim lngType As Long                 'code for the type of a fill, color,
                                        'or outline.
                                        
    
                                        'Does the SHAPE match the MODEL ?
                                        'Exit immediately on any mismatch.
    With shpShape
        lngShapeType = .Type            'Same basic TYPE of shape ?
        lngModelType = shpModel.Type
        
        If CheckType Then If lngShapeType <> lngModelType Then GoTo NoMatch
                                        'A GROUP ? delegate to GroupsMatch()
'        If lngShapeType = cdrGroupShape Then
'            ShapesMatch = GroupsMatch(shpShape, shpModel, CheckSize, _
'                                CountNodes, CountPaths)
'            Exit Function
'        End If

                                        'Does SIZE count ? Is so, are the
        If CheckSize Then               'size differences significant ?
            dblWidth = shpModel.SizeWidth
            If Abs(.SizeWidth - dblWidth) > (dblWidth * _
                 ToleranceSize) Then GoTo NoMatch
            dblHeight = shpModel.SizeHeight
            If Abs(.SizeHeight - dblHeight) > (dblHeight * _
                ToleranceSize) Then GoTo NoMatch
        End If
        
        If CheckWHratio Then               'size width and height ratio differences significant ?
            dblWidth = shpModel.SizeWidth
            dblHeight = shpModel.SizeHeight
            If Abs(.SizeHeight / .SizeWidth - dblHeight / dblWidth) > (dblHeight / dblWidth * ToleranceWHratio) Then GoTo NoMatch
        End If
        

            If CountNodes Or CountPaths Or CheckOutlineLength Or CountSegments Then
                                        'Only Curves can match ...
                If lngShapeType <> cdrCurveShape Then GoTo NoMatch
                
                Set crvShape = .Curve
                Set crvModel = shpModel.Curve
                
                'If CheckIndiv Then '逐条子路径比较
                    'If Abs(crvShape.SubPaths.Count - crvModel.SubPaths.Count) <> 0 Then GoTo NoMatch
                    'For j = 1 To crvShape.SubPaths.Count
                            'If Abs(crvShape.SubPath(j).Nodes.Count - crvModel.SubPath(j).Nodes.Count) > ToleranceNodesCount Then GoTo NoMatch
                     
                     'Next j
                
                If CountPaths Then      'Do the PATH counts match ?
                    
                    If VersionMajor > 12 Then 'GDG ##########################################
                        If Abs(crvShape.SubPaths.Count - crvModel.SubPaths.Count) > ToleranceSubPathsCount Then GoTo NoMatch
                        'MsgBox "subpaths1: " & crvShape.SubPaths.Count & "subpaths2: " & crvModel.SubPaths.Count
                    Else
                        If Abs(crvShape.SubPathCount - crvModel.SubPathCount) > ToleranceSubPathsCount Then GoTo NoMatch
                    End If 'GDG #############################################################
                    
                End If
                
                
                 
                 
                If CountNodes Then      'Do the NODE counts match ?
                
                    If VersionMajor > 12 Then 'GDG ##########################################
                        If Abs(crvShape.Nodes.Count - crvModel.Nodes.Count) > ToleranceNodesCount Then GoTo NoMatch
                    Else
                        If Abs(crvShape.NodeCount - crvModel.NodeCount) > ToleranceNodesCount Then GoTo NoMatch
                    End If 'GDG #############################################################
                    
                End If
                
                If CountSegments Then      'Do the Segments counts match ?
                
                    If VersionMajor > 12 Then 'GDG ##########################################
                        If Abs(crvShape.Segments.Count - crvModel.Segments.Count) > ToleranceSegmentsCount Then GoTo NoMatch
                    Else
                        If Abs(crvShape.SegmentCount - crvModel.SegmentCount) > ToleranceSegmentsCount Then GoTo NoMatch
                    End If 'GDG #############################################################
                    
                End If
        
                
                
                If CheckOutlineLength Then      'Do the curve length match ?
                
                    If VersionMajor > 12 Then 'GDG ##########################################
                        If Abs(crvShape.Length - crvModel.Length) > crvModel.Length * ToleranceLength Then GoTo NoMatch
                        'MsgBox "subpaths1: " & crvShape.SubPaths.Count & "subpaths2: " & crvModel.SubPaths.Count
                    Else
                        If Abs(crvShape.Length - crvModel.Length) > crvModel.Length * ToleranceLength Then GoTo NoMatch
                    End If 'GDG #############################################################
                    
                End If
            End If
        If CheckFill Then
            Set fillModel = shpModel.Fill
            With .Fill                  'Is the FILL type the same ?
                lngType = .Type
                If lngType <> shpModel.Fill.Type Then GoTo NoMatch
                If lngType = cdrUniformFill Then
'Does the uniform fill match ?
                    If VersionMajor > 12 Then 'GDG ##########################################
                        'GDG ##########################################
                        Dim col1 As New Color
                        col1.CopyAssign .UniformColor
                        Dim col2 As New Color
                        col2.CopyAssign shpModel.Fill.UniformColor
                        'GDG ##########################################
                        If col1.IsSame(col2) = False Then GoTo NoMatch
                    Else
                        Set clrModel = fillModel.UniformColor
                        lngType = .UniformColor.Type
                        If lngType <> clrModel.Type Then GoTo NoMatch
                        If .UniformColor.name(True) <> clrModel.name(True) Then GoTo NoMatch
                    End If  'GDG #############################################################
                End If
            End With
        End If
        
        
        
        If CheckOutline Then            '(Groups have no outline)
            If lngShapeType <> cdrGroupShape Then
                Set outlnModel = shpModel.Outline
                If Not outlnModel Is Nothing Then
                    With .Outline
                        lngType = .Type
                        If lngType <> outlnModel.Type Then GoTo NoMatch
                                                
                        If lngType > 0 Then     'Does the shape have an OUTLINE ?
                                                'Same LINE WIDTH ?
                            If .width <> outlnModel.width Then GoTo NoMatch
                                                'Matching LINE COLOR ?
'                            Set clrShape = .Color
'                            lngType = clrShape.Type
'                            Set clrModel = outlnModel.Color
'                            If lngType <> clrModel.Type Then GoTo NoMatch
'                            If clrShape.Name(True) <> clrModel.Name(True) Then GoTo NoMatch
                        End If
                    End With
                End If
            End If
        End If
        
        
        If CheckOutlineColor Then            '(Groups have no outline)
            If lngShapeType <> cdrGroupShape Then
 
               
                Set outlnModel = shpModel.Outline
                If Not outlnModel Is Nothing Then
                    
                    With .Outline
                        lngType = .Type
                        If lngType <> outlnModel.Type Then GoTo NoMatch
                                                
                        If lngType > 0 Then     'Does the shape have an OUTLINE ?
                                                'Matching LINE COLOR ?
                            
                            If VersionMajor > 12 Then 'GDG ##########################################
                                'GDG ##########################################
                                Dim col3 As New Color
                                col3.CopyAssign .Color
                                Dim col4 As New Color
                                col4.CopyAssign shpModel.Outline.Color
                                'GDG ##########################################
                                If col3.IsSame(col4) = False Then GoTo NoMatch
                            Else
                                Set clrShape = .Color
                                lngType = clrShape.Type
                                Set clrModel = outlnModel.Color
                                If lngType <> clrModel.Type Then GoTo NoMatch
                                If clrShape.name(True) <> clrModel.name(True) _
                                    Then GoTo NoMatch
                            End If
                        End If
                    End With
                End If
            End If
        End If
        
    End With
    
    ShapesMatch = True
    Exit Function
    
NoMatch:
    ShapesMatch = False
    
NoMatchExit:
    ShapesMatch = False
    Exit Function
End Function

Private Function GroupsMatch(Group As Shape, GroupModel As Shape, _
                    Optional CheckFill As Boolean = True, _
                    Optional CheckOutline As Boolean = True, _
                    Optional CheckOutlineColor As Boolean = True, _
                    Optional CheckOutlineLength As Boolean = True, _
                    Optional CheckSize As Boolean = False, _
                    Optional CheckType As Boolean = True, _
                    Optional CountNodes As Boolean = False, _
                    Optional CountPaths As Boolean = False) As Boolean
    
    'Object Variables.              Reference to:
    Dim shpsModels As Shapes            'shapes in the pre-selected group,
    Dim shpsInGroup As Shapes           'shapes in the group to be tested,
    Dim shpModel As Shape               'a shape in the pre-selected group,
    Dim shpInGroup As Shape             'a shape in the group to be tested.
    
    'Simple Variables               Storage of:
    Dim lngInGroup As Long              'number of shapes in a group,
    Dim i As Long                       'a numeric index to a
                                        'particular sub-group.
                                        
    'On Error GoTo NoMatch              'Shape & model must be groups.
    Set shpsModels = GroupModel.Shapes
    Set shpsInGroup = Group.Shapes
    'On Error GoTo 0
                                        'Same number of shapes
    lngInGroup = shpsModels.Count       'in each group ?
    If shpsInGroup.Count <> lngInGroup Then GoTo NoMatch
        
    For i = 1 To lngInGroup             'Try to Match all sub-shapes,
        Set shpInGroup = shpsInGroup(i) 'and GroupsMatch all sub-groups.
        Set shpModel = shpsModels(i)
        
        If shpModel.Type <> cdrGroupShape Then
            If Not ShapesMatch(shpInGroup, shpModel, _
                            CheckSize, CountNodes) Then GoTo NoMatch
        Else
            If Not GroupsMatch(shpInGroup, shpModel, _
                            CheckSize, CountNodes) Then GoTo NoMatch
        End If
    Next i
    
    GroupsMatch = True
    Exit Function
NoMatch:
    GroupsMatch = False
End Function


Private Function FlatShapeList(TopLevelShapes As Shapes) As Collection
    
    'Object Variables.          Reference to:
    Dim shpTopLevel As Shape        'a top-level shape,
    Dim shpInGroup As Shape         'a shape inside a group,
    Dim clnAllShapes As Collection  'our list of all members and
                                    'descendants of TopLevelShapes.
                                       
    If TopLevelShapes.Count Then
        Set clnAllShapes = New Collection
        For Each shpTopLevel In TopLevelShapes
                                    'Add shape to list, keyed under
                                    'a string version of its unique ID
             clnAllShapes.Add shpTopLevel
                                    'If the shape is a group, then
                                    'also gather all its descendants
                                    'and add them to the list.
            If shpTopLevel.Type = cdrGroupShape Then
                For Each shpInGroup In ShapesInGroup(shpTopLevel)
               clnAllShapes.Add shpInGroup
                Next
            End If
        Next
        Set FlatShapeList = clnAllShapes  'Return the assembled collection.
    Else
        Set FlatShapeList = Nothing
    End If
End Function

Private Function ShapesInGroup(GroupShape As Shape) As Collection

    'Object Variables.              Reference to:
    Dim shpsInGroup As Shapes           'the set of shapes inside a group,
    Dim shpInGroup As Shape             'a particular shape in a group,
    Dim shpNested As Shape              'a shape inside a sub-group,
    Dim clnShapeList As Collection      'our list of all nested shapes.
    
    If GroupShape.Type = cdrGroupShape Then
        Set shpsInGroup = GroupShape.Shapes 'Get a reference to the
                                            'shapes in this group.
        Set clnShapeList = New Collection
        For Each shpInGroup In shpsInGroup
            clnShapeList.Add shpInGroup     'Add all shapes in the group to
                                            'our main collection.
            If shpInGroup.Type = cdrGroupShape Then
                                            'Recurse to get nested shapes.
                For Each shpNested In ShapesInGroup(shpInGroup)
                    clnShapeList.Add shpNested
                Next
            End If
        Next
        Set ShapesInGroup = clnShapeList    'Return the assembled collection.
    Else
        Set ShapesInGroup = Nothing         'Release the memory if the
    End If                                  'collection is not needed
End Function

Private Sub UserForm_Activate()
    Const YES As String = "True"
    Const NO As String = "False"

    OptDoc = GetSetting(TOOLNAME, SECTION, "InDoc", NO)
    Optlayer = GetSetting(TOOLNAME, SECTION, "InLayer", NO)
    Optpage = GetSetting(TOOLNAME, SECTION, "InPage", YES)
    
    chkColorMark = GetSetting(TOOLNAME, SECTION, "ColorMark", YES)
    chkFill = GetSetting(TOOLNAME, SECTION, "Fill", YES)
    chkInGroups = GetSetting(TOOLNAME, SECTION, "InGroups", YES)
    chkNodes = GetSetting(TOOLNAME, SECTION, "Nodes", NO)
    chkSegments = GetSetting(TOOLNAME, SECTION, "Segments", NO)
    chkOutline = GetSetting(TOOLNAME, SECTION, "Outline", YES)
    chkOutlineColor = GetSetting(TOOLNAME, SECTION, "OutlineColor", NO)
    chkOutlineLength = GetSetting(TOOLNAME, SECTION, "OutlineLength", YES)
    chkPaths = GetSetting(TOOLNAME, SECTION, "Paths", NO)
    chkSize = GetSetting(TOOLNAME, SECTION, "Size", NO)
    chkWHratio = GetSetting(TOOLNAME, SECTION, "WHratio", NO)
    chkType = GetSetting(TOOLNAME, SECTION, "Type", YES)
    chkIndiv = GetSetting(TOOLNAME, SECTION, "Indiv", NO)
    chkColorMark = GetSetting(TOOLNAME, SECTION, "ColorMark", NO)
    saveFormPos False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    saveFormPos True
End Sub

Sub saveFormPos(bDoSave As Boolean)
    Dim dL, dT
    If bDoSave Then 'save position
         SaveSetting TOOLNAME, SECTION, "form_left", Me.Left
         SaveSetting TOOLNAME, SECTION, "form_top", Me.Top
    Else 'place instead.
        dL = GetSetting(TOOLNAME, SECTION, "form_left", 900)
        dT = GetSetting(TOOLNAME, SECTION, "form_top", 200)
        Me.Left = dL: Me.Top = dT
    End If
End Sub

Private Sub OptDoc_Click()
    SaveSetting TOOLNAME, SECTION, "InDoc", CStr(OptDoc)
End Sub
Private Sub Optlayer_Click()
    SaveSetting TOOLNAME, SECTION, "InLayer", CStr(Optlayer)
End Sub
Private Sub Optpage_Click()
    SaveSetting TOOLNAME, SECTION, "InPage", CStr(Optpage)
End Sub
Private Sub chkColorMark_Click()
    SaveSetting TOOLNAME, SECTION, "ColorMark", CStr(chkColorMark)
End Sub
Private Sub chkIndiv_Click()
    SaveSetting TOOLNAME, SECTION, "Indiv", CStr(chkIndiv)
End Sub
Private Sub chkFill_Click()
    SaveSetting TOOLNAME, SECTION, "Fill", CStr(chkFill)
End Sub
Private Sub chkInGroups_Click()
    SaveSetting TOOLNAME, SECTION, "InGroups", CStr(chkInGroups)
End Sub
Private Sub chkNodes_Click()
    SaveSetting TOOLNAME, SECTION, "Nodes", CStr(chkNodes)
End Sub
Private Sub chkSegments_Click()
    SaveSetting TOOLNAME, SECTION, "Segments", CStr(chkSegments)
End Sub
Private Sub chkOutline_Click()
    SaveSetting TOOLNAME, SECTION, "Outline", CStr(chkOutline)
End Sub
Private Sub chkOutlineColor_Click()
    SaveSetting TOOLNAME, SECTION, "OutlineColor", CStr(chkOutlineColor)
End Sub
Private Sub chkPaths_Click()
    SaveSetting TOOLNAME, SECTION, "Paths", CStr(chkPaths)
End Sub
Private Sub chkSize_Click()
    SaveSetting TOOLNAME, SECTION, "Size", CStr(chkSize)
End Sub
Private Sub chkWHratio_Click()
    SaveSetting TOOLNAME, SECTION, "WHratio", CStr(chkWHratio)
End Sub
Private Sub chkType_Click()
    SaveSetting TOOLNAME, SECTION, "Type", CStr(chkType)
End Sub
Private Sub chkOutLineLength_Click()
    SaveSetting TOOLNAME, SECTION, "OutlineLength", CStr(chkOutlineLength)
End Sub
