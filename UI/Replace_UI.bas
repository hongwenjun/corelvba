
Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Dim pos_x As Variant
  Dim pos_Y As Variant
  pos_x = Array(307, 27)
  pos_Y = Array(64, 126, 188, 200)

  If Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(0)) < 30 Then
    Call copy_shape_replace
  ElseIf Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(1)) < 30 Then
    Call copy_shape_replace_resize
  ElseIf Abs(X - pos_x(0)) < 30 And Abs(Y - pos_Y(2)) < 30 Then
    Call image_replace
  ElseIf Abs(X - pos_x(1)) < 30 And Abs(Y - pos_Y(3)) < 30 Then
    CorelVBA.WebHelp "https://262235.xyz/index.php/tag/vba/"
  End If
  
  Replace_UI.Hide
End Sub


Private Sub image_replace()
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True
  Dim image_path As String
  image_path = API.GetClipBoardString
  ActiveDocument.ReferencePoint = cdrCenter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim X As Double, Y As Double
  Set shs = ActiveSelection.Shapes
  cnt = 0
  For Each sh In shs
    If cnt = 0 Then
      ActiveDocument.ClearSelection
      ActiveLayer.Import image_path
      Set sc = ActiveSelection
      cnt = 1
    Else
      sc.Duplicate 0, 0
    End If
    sh.GetPosition X, Y
    sc.SetPosition X, Y
    
    sh.GetSize X, Y
    sc.SetSize X, Y
    sh.Delete
    
  Next sh

    '// 代码操作结束恢复窗口刷新
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh:    Application.Refresh
Exit Sub
ErrorHandler:
    MsgBox "请先复制图片的完整路径，本工具能自动替换图片!"
    Application.Optimization = False
    On Error Resume Next
End Sub

Private Sub copy_shape_replace_resize()
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True

  ActiveDocument.ReferencePoint = cdrCenter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim X As Double, Y As Double
  Set shs = ActiveSelection.Shapes
  cnt = 0
  For Each sh In shs
    If cnt = 0 Then
      Set sc = ActiveDocument.ActiveLayer.Paste
      cnt = 1
    Else
      sc.Duplicate 0, 0
    End If
    sh.GetPosition X, Y
    sc.SetPosition X, Y
    
    sh.GetSize X, Y
    sc.SetSize X, Y
    sh.Delete
    
  Next sh

    '// 代码操作结束恢复窗口刷新
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh:    Application.Refresh
Exit Sub
ErrorHandler:
    MsgBox "请先复制Ctrl+C，然后选择要替换的物件运行本工具!"
    Application.Optimization = False
    On Error Resume Next
End Sub


Private Sub copy_shape_replace()
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True

  ActiveDocument.ReferencePoint = cdrCenter
  Dim sh As Shape, shs As Shapes, cs As Shape
  Dim X As Double, Y As Double
  Set shs = ActiveSelection.Shapes
  cnt = 0
  For Each sh In shs
    If cnt = 0 Then
      Set sc = ActiveDocument.ActiveLayer.Paste
      cnt = 1
    Else
      sc.Duplicate 0, 0
    End If
    sh.GetPosition X, Y
    sc.SetPosition X, Y
    sh.Delete
  Next sh

    '// 代码操作结束恢复窗口刷新
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveWindow.Refresh:    Application.Refresh
Exit Sub
ErrorHandler:
    MsgBox "请先复制Ctrl+C，然后选择要替换的物件运行本工具!"
    Application.Optimization = False
    On Error Resume Next
End Sub

