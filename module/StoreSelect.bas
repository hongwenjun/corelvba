Attribute VB_Name = "StoreSelect"
Private sr_mem(3) As New ShapeRange
Public StoreCount As String

Public Function Store_Instruction(id As Integer, INST As String) As String
  On Error GoTo ErrorHandler
  API.BeginOpt "Undo MRC"
  '// 选择指令执行
  Case_Select_Range id, INST
  
  StoreCount = "Store Count: A->" & sr_mem(1).Count & "  B->" & sr_mem(2).Count & "  C->" & sr_mem(3).Count
  API.EndOpt
  
Exit Function

ErrorHandler:
  Application.Optimization = False
End Function

Private Function Case_Select_Range(id As Integer, INST As String)
  On Error GoTo ErrorHandler
  Select Case INST
    Case "add"
      sr_mem(id).AddRange ActiveSelectionRange
    Case "sub"
      sr_mem(id).RemoveRange ActiveSelectionRange
    Case "lw"
     '// ActiveDocument.ClearSelection
      sr_mem(id).AddToSelection
    Case "zero"
      If id = 3 Then
        sr_mem(3).RemoveAll: sr_mem(1).RemoveAll: sr_mem(2).RemoveAll
      Else
        sr_mem(id).RemoveAll
    End If

  End Select
  
Exit Function

ErrorHandler:
  Application.Optimization = False
End Function
