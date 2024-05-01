VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} arrow 
   Caption         =   "¼ýÍ·Ìæ»»¹¤¾ß    github.com/hongwenjun"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   OleObjectBlob   =   "arrow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "arrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
  arrowtool.SetArrow
End Sub


Private Sub CommandButton2_Click()
  arrowtool.arrow_manual_tool
End Sub

Private Sub CommandButton3_Click()
  arrowtool.arrow_Batch_repalce
End Sub

Private Sub CommandButton4_Click()
  arrowtool.turn_over
End Sub
