Attribute VB_Name = "HDPI"
#If VBA7 Then
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
  Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#Else
  Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
  Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

Public Function GetHDPI() As Double
    Const LOGPIXELSX = 88
    Const HORZRES = 8
    Dim hdc As Long, dpi As Long, width As Long
    
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    width = GetDeviceCaps(hdc, HORZRES)
    ReleaseDC 0, hdc
    
    GetHDPI = dpi / width * 25.4
End Function

Public Function GetHDPIPercentage() As Integer
    Const LOGPIXELSX = 88
    Const HORZRES = 8
    Dim hdc As Long, dpi As Long, width As Long
    
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    width = GetDeviceCaps(hdc, HORZRES)
    ReleaseDC 0, hdc
    
    GetHDPIPercentage = Round(dpi / 96 * 100, 0)
End Function

Sub MSG_HDPIPercentage()
  MsgBox "HDPI Percentage:" & GetHDPIPercentage
End Sub
