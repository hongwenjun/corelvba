Attribute VB_Name = "SmartGroup"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// Attribute VB_Name = "智能群组"   SmartGroup  2026.05.23 更换AI转的VBA 智能群群租

' 定义边界框结构
Private Type BoundingBox
    X As Double
    Y As Double
    w As Double
    h As Double
End Type

Public Function Smart_Group(Optional ByVal tr As Double = 0) As ShapeRange
  On Error GoTo ErrorHandler
  API.BeginOpt

  Box_AutoGroup_VBA tr   '// 2026.05.23 更换AI转的VBA 智能群群租

ErrorHandler:
  API.EndOpt
End Function

'// 旧智能群组 原理版
Private Function Smart_Group_ABC()
  ActiveDocument.Unit = cdrMillimeter
  
  Dim OrigSelection As ShapeRange, brk1 As ShapeRange
  Set OrigSelection = ActiveSelectionRange
  Dim s1 As Shape, sh As Shape, s As Shape
  
  Set s1 = OrigSelection.CustomCommand("Boundary", "CreateBoundary")
  Set brk1 = s1.BreakApartEx

  For Each s In brk1
    If s.SizeHeight > 10 Then
      Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.topY, s.RightX, s.BottomY, False)
      sh.Shapes.all.Group
    End If
    s.Delete
  Next
End Function

' 1. 检查两个矩形是否重叠 (AABB 碰撞检测)
Private Function IsOverlapped(a As BoundingBox, b As BoundingBox) As Boolean
    IsOverlapped = (a.X < b.X + b.w) And (a.X + a.w > b.X) And _
                   (a.Y < b.Y + b.h) And (a.Y + a.h > b.Y)
End Function

' 2. 并查集：查找根节点（含路径压缩）
Private Function FindParent(ByRef Parent() As Long, ByVal i As Long) As Long
    If Parent(i) <> i Then
        Parent(i) = FindParent(Parent, Parent(i))
    End If
    FindParent = Parent(i)
End Function

' 3. 并查集：合并集合
Private Sub UnionSet(ByRef Parent() As Long, ByVal X As Long, ByVal Y As Long)
    Dim rootX As Long, rootY As Long
    rootX = FindParent(Parent, X)
    rootY = FindParent(Parent, Y)
    If rootX <> rootY Then Parent(rootX) = rootY
End Sub

' 核心功能：自动分组
Public Function Box_AutoGroup_VBA(Optional ByVal exp As Double = 0)
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    ' 如果没选，尝试全选
    If sr.count = 0 Then
        ActivePage.Shapes.all.CreateSelection
        Set sr = ActiveSelectionRange
    End If
    
    If sr.count = 0 Then Exit Function

    Dim i As Long, j As Long
    Dim count As Long: count = sr.count
    Dim boxes() As BoundingBox
    Dim parentArr() As Long
    
    ReDim boxes(1 To count)
    ReDim parentArr(1 To count)

    ' --- 第一步：获取所有形状的边界框并初始化并查集 ---
    Dim s As Shape
    For i = 1 To count
        Set s = sr.Shapes(i)
        ' 获取边界框 (VBA 中获取左、下、宽、高)
        s.GetBoundingBox boxes(i).X, boxes(i).Y, boxes(i).w, boxes(i).h
        
        ' 扩展边界框 (逻辑同 C++ expand_bounding_boxes)
        If Abs(exp) > 0.02 Then
            boxes(i).X = boxes(i).X - exp
            boxes(i).Y = boxes(i).Y - exp
            boxes(i).w = boxes(i).w + 2 * exp
            boxes(i).h = boxes(i).h + 2 * exp
        End If
        
        parentArr(i) = i ' 初始化父节点为自己
    Next i

    ' --- 第二步：运行 Union-Find 算法检测重叠 ---
    For i = 1 To count
        For j = i + 1 To count
            If IsOverlapped(boxes(i), boxes(j)) Then
                UnionSet parentArr, i, j
            End If
        Next j
    Next i

    ' --- 第三步：根据根节点进行物理分组 ---
    ' 使用 Collection 模拟 C++ 的 std::map<int, std::vector<int>>
    Dim Groups As New Collection
    Dim rootID As Long
    Dim groupMemberSR As ShapeRange
    
    ' 预处理：将同一组的形状放到一起
    ' 我们用数组记录每个根节点对应的 ShapeRange
    Dim GroupSRs() As ShapeRange
    ReDim GroupSRs(1 To count)
    
    For i = 1 To count
        rootID = FindParent(parentArr, i)
        If GroupSRs(rootID) Is Nothing Then
            Set GroupSRs(rootID) = CreateShapeRange
        End If
        GroupSRs(rootID).Add sr.Shapes(i)
    Next i

    
    ActiveDocument.ClearSelection

    ' 遍历并执行 Group 操作
    Dim finalSR As New ShapeRange
    Dim totalGroups As Long: totalGroups = 0
    
    For i = 1 To count
        If Not GroupSRs(i) Is Nothing Then
            If GroupSRs(i).count > 1 Then
                finalSR.Add GroupSRs(i).Group
                totalGroups = totalGroups + 1
            Else
                finalSR.Add GroupSRs(i)(1)
                totalGroups = totalGroups + 1
            End If
        End If
    Next i

    finalSR.CreateSelection
    
End Function

