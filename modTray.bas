Attribute VB_Name = "modTray"
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private m_udtTrayNID As NOTIFYICONDATA

Public Sub TrayAddModify(strToolTip As String)
    Const NIF_ICON = &H2
    Const NIF_TIP = &H4
    Const NIF_MESSAGE = &H1
    Const NIM_ADD = &H0
    Const NIM_MODIFY = &H1
    Const WM_MOUSEMOVE = &H200
    
    With m_udtTrayNID
        .cbSize = Len(m_udtTrayNID)
        .hwnd = frmTray.hwnd
        .uID = 1
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .szTip = strToolTip & vbNullChar
        If .hIcon = 0 Then
            .hIcon = frmTray.Icon.Handle
        End If
    End With
    
    If Not Shell_NotifyIcon(NIM_MODIFY, m_udtTrayNID) Then
        Shell_NotifyIcon NIM_ADD, m_udtTrayNID
    End If
End Sub

Public Sub TrayDelete()
    Const NIM_DELETE = &H2
    
    Shell_NotifyIcon NIM_DELETE, m_udtTrayNID
    If m_udtTrayNID.hIcon <> 0 Then DestroyIcon m_udtTrayNID.hIcon
End Sub



