Attribute VB_Name = "StoreSelect"
Public sr_mem(5) As New ShapeRange
Public StoreCount As String

Public Function Store_Instruction(id As Integer, INST As String) As String
  On Error GoTo ErrorHandler
  API.BeginOpt "Undo MRC"
  '// Ñ¡ÔñÖ¸ÁîÖ´ÐÐ
  Case_Select_Range id, INST
  
  StoreCount = "Store Count: A->" & sr_mem(1).count & "  B->" & sr_mem(2).count & "  C->" & sr_mem(3).count

ErrorHandler:
  API.EndOpt
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
      '// sr_mem(id).AddToSelection
      sr_mem(id).CreateSelection
    Case "zero"
      If id = 3 Then
        sr_mem(3).RemoveAll: sr_mem(1).RemoveAll: sr_mem(2).RemoveAll
      Else
        sr_mem(id).RemoveAll
      End If
    Case "sw"
      sr_mem(id).RemoveAll
      sr_mem(id).AddRange ActiveSelectionRange
  End Select

ErrorHandler:
End Function

Public Function SRMInst(id As Integer, INST As String)
  Case_Select_Range id, INST
End Function
