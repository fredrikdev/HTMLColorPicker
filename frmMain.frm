VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   MouseIcon       =   "frmMain.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   30
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1815
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Bytes As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private m_blnPicked As Boolean
Private m_lngPickCount As Long

Private Sub Form_Load()
    Const HWND_TOPMOST = -1
    Const SWP_NOACTIVATE = &H10, SWP_NOMOVE = &H2, SWP_NOSIZE = &H1, SWP_SHOWWINDOW = &H40
    Dim lngHRgn1 As Long, lngHRgn2 As Long, lngHRgn As Long, strTemp As String

    ' size the window to fit the entire screen, then copy the screen
    Me.Move 0, 0, Screen.Width, Screen.Height
    modSystem.copyScreen Me.hdc
    
    ' load zoom glas & create clipping
    lngHRgn1 = CreateEllipticRgn(1, 1, 77, 77)
    lngHRgn2 = CreateEllipticRgn(78, 2, 122, 46)
    lngHRgn = CreateEllipticRgn(0, 0, 0, 0)
    CombineRgn lngHRgn, lngHRgn1, lngHRgn2, 2
    
    SetWindowRgn picInfo.hwnd, lngHRgn, False
    
    DeleteObject lngHRgn1
    DeleteObject lngHRgn2
    DeleteObject lngHRgn
    
    strTemp = GetSetting(App.Title, "Statistics", "PickCount", "0")
    If Not IsNumeric(strTemp) Then m_lngPickCount = 0 Else m_lngPickCount = CLng(strTemp)
    m_blnPicked = False
    
    If m_lngPickCount = 0 Then
        MsgBox "The zoom & preview panel will now open. Left-click a color to copy it (right-click to abort)" & vbCrLf & vbCrLf & "This message will not be displayed again.", vbInformation Or vbOKOnly
    End If
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub

Private Function getWebColor(lngColor As Long) As String
    Dim b(1 To 4) As Byte, strReturn As String
    
    MemCopy b(1), lngColor, 4
    
    ' r
    strReturn = Hex$(b(3))
    If Len(strReturn) = 1 Then strReturn = "0" & strReturn
    
    ' g
    strReturn = Hex$(b(2)) & strReturn
    If Len(strReturn) = 3 Then strReturn = "0" & strReturn
    
    ' b
    strReturn = Hex$(b(1)) & strReturn
    If Len(strReturn) = 5 Then strReturn = "0" & strReturn
    
    getWebColor = "#" & strReturn
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_blnPicked = True Then Exit Sub
    m_blnPicked = True
        
    If Button = vbLeftButton Then
        Clipboard.Clear
        Clipboard.SetText getWebColor(picInfo.FillColor), ClipBoardConstants.vbCFText
        
        If m_lngPickCount = 0 Then
            MsgBox "The selected color has now been copied to your clipboard as #RRGGBB." & vbCrLf & vbCrLf & "This message will not be displayed again.", vbInformation Or vbOKOnly
        End If
        m_lngPickCount = m_lngPickCount + 1
        SaveSetting App.Title, "Statistics", "PickCount", m_lngPickCount
    
        frmTray.lblTitle.Caption = App.Title & " - " & m_lngPickCount & " Colors Picked"
    End If
    
    Me.Visible = False
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const conZoomWidth = 11
    Const conZoomFactor = 6 * conZoomWidth
    Const conCrosshairPos = 40
    Dim lngX As Long, lngY As Long
    
    ' determine position
    lngX = X + 2
    lngY = Y + 2
    If lngX + picInfo.ScaleWidth > Me.ScaleWidth Then lngX = X - 8 - picInfo.ScaleWidth
    If lngY + picInfo.ScaleHeight > Me.ScaleHeight Then lngY = Y - 18 - picInfo.ScaleHeight
    
    ' generate zoomed image
    StretchBlt picInfo.hdc, 6, 6, conZoomFactor, conZoomFactor, Me.hdc, X - 5, Y - 5, 11, 11, &HCC0020
    
    ' draw zoom circle
    picInfo.DrawWidth = "5"
    picInfo.FillStyle = FillStyleConstants.vbFSTransparent
    picInfo.Circle (38, 38), 35
        
    ' draw color circle
    picInfo.DrawWidth = "3"
    picInfo.FillStyle = FillStyleConstants.vbFSSolid
    picInfo.FillColor = Me.Point(X, Y)
    picInfo.Circle (99, 23), 20
    
    ' draw crosshair
    picInfo.Line (conCrosshairPos - 3, conCrosshairPos - 3)-(conCrosshairPos, conCrosshairPos), vbBlack, BF
    
    picInfo.Move lngX, lngY
    picInfo.Visible = True
End Sub

