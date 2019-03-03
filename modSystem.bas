Attribute VB_Name = "modSystem"
' -------------------------------------------------------------------
'  system functions
'  - copyScreen(lngDestDC): copy the screen to the form
' -------------------------------------------------------------------
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

' -------------------------------------------------------------------
'  copies screen to a form
' -------------------------------------------------------------------
Public Sub copyScreen(ByVal lngDestDC As Long)
    Dim lngDeskWnd As Long, lngDeskDC As Long
    Dim lngW As Long, lngH As Long
    
    lngW = Screen.Width / Screen.TwipsPerPixelX
    lngH = Screen.Height / Screen.TwipsPerPixelY
    
    lngDeskWnd = GetDesktopWindow
    lngDeskDC = GetWindowDC(lngDeskWnd)
    
    BitBlt lngDestDC, 0, 0, lngW, lngH, lngDeskDC, 0, 0, &HCC0020
    
    ReleaseDC lngDeskWnd, lngDeskDC
End Sub
