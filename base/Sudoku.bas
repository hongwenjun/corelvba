Private Sub Sudoku_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  Me.Sudoku.Visible = False
  ' 定义九宫格方向上下左右等
  Dim pos_x As Variant, pos_y As Variant
  pos_x = Array(16, 16, 27, 27, 27, 16, 4, 4, 4, 16)
  pos_y = Array(16, 4, 4, 16, 27, 27, 27, 16, 4, 16)
  
  If Abs(X - pos_x(0)) < 4 And Abs(Y - pos_y(0)) < 4 Then
    Me.Sudoku.Picture = bmp0.Picture
    
  ElseIf Abs(X - pos_x(1)) < 4 And Abs(Y - pos_y(1)) < 4 Then
    Me.Sudoku.Picture = bmp1.Picture

  ElseIf Abs(X - pos_x(2)) < 4 And Abs(Y - pos_y(2)) < 4 Then
    Me.Sudoku.Picture = bmp2.Picture

  ElseIf Abs(X - pos_x(3)) < 4 And Abs(Y - pos_y(3)) < 4 Then
    Me.Sudoku.Picture = bmp3.Picture

  ElseIf Abs(X - pos_x(4)) < 4 And Abs(Y - pos_y(4)) < 4 Then
    Me.Sudoku.Picture = bmp4.Picture
    
  ElseIf Abs(X - pos_x(5)) < 4 And Abs(Y - pos_y(5)) < 4 Then
    Me.Sudoku.Picture = bmp5.Picture
  
  ElseIf Abs(X - pos_x(6)) < 4 And Abs(Y - pos_y(6)) < 4 Then
    Me.Sudoku.Picture = bmp6.Picture
    
  ElseIf Abs(X - pos_x(7)) < 4 And Abs(Y - pos_y(7)) < 4 Then
    Me.Sudoku.Picture = bmp7.Picture
    
  ElseIf Abs(X - pos_x(8)) < 4 And Abs(Y - pos_y(8)) < 4 Then
    Me.Sudoku.Picture = bmp8.Picture
    
  End If
  
  Me.Sudoku.Visible = True
End Sub

