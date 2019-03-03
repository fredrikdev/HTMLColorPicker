VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Color Picker"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
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
   Icon            =   "frmTray.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton btnOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3945
      TabIndex        =   5
      Top             =   4170
      Width           =   1125
   End
   Begin VB.Timer ctlTimer 
      Interval        =   2500
      Left            =   4785
      Top             =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Tip: To easily find a certain color, first open this dialog then pick a color from the spectrum above!"
      Height          =   600
      Left            =   135
      TabIndex        =   6
      Top             =   2355
      Width           =   4950
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "frmTray.frx":058A
      Top             =   0
      Width           =   5220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.hoursandminutes.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2820
      MouseIcon       =   "frmTray.frx":470C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3615
      Width           =   1995
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) by Port Jackson Computing. For upgrades, questions and other fine software navigate to:"
      Height          =   465
      Left            =   135
      TabIndex        =   3
      Top             =   3420
      UseMnemonic     =   0   'False
      Width           =   4845
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTray.frx":4A16
      Height          =   885
      Left            =   135
      TabIndex        =   2
      Top             =   1380
      UseMnemonic     =   0   'False
      Width           =   4950
   End
   Begin VB.Label lblInstructionsTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions:"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   1050
      Width           =   915
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   390
      Width           =   45
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   367
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Line ctlLine 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   369
      Y1              =   59
      Y2              =   59
   End
   Begin VB.Shape ctlShape 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   915
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   5520
   End
   Begin VB.Shape ctlShape 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   1
      Left            =   0
      Top             =   3975
      Width           =   5520
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuHelp 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long

Private Sub btnOk_Click()
    Me.Visible = False
End Sub

Private Sub ctlTimer_Timer()
    modTray.TrayAddModify App.Title
End Sub

Private Sub Form_Load()
    Dim lpMutexAttributes As SECURITY_ATTRIBUTES

    lpMutexAttributes.nLength = Len(lpMutexAttributes)
    CreateMutex lpMutexAttributes, False, "PJHTMLColorPicker"
    
    App.TaskVisible = False
    ctlTimer_Timer
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const WM_LBUTTONDBLCLK = &H203
    Const WM_RBUTTONUP = &H205
    
    Select Case X
        Case WM_LBUTTONDBLCLK
            If frmTray.Visible Then modTray.SetForegroundWindow frmTray.hwnd
            frmMain.Show
        Case WM_RBUTTONUP
            modTray.SetForegroundWindow frmTray.hwnd
            PopupMenu frmTray.mnuTray, , , , mnuHelp
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modTray.TrayDelete
End Sub

Private Sub Label1_Click()
    ShellExecute hwnd, "open", "http://www.hoursandminutes.com", vbNullString, vbNullString, 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHelp_Click()
    Dim lngPickCount As Long
    lngPickCount = GetSetting(App.Title, "Statistics", "PickCount", "0")
    If Not IsNumeric(lngPickCount) Then lngPickCount = "0"
    
    lblTitle.Caption = App.Title & " - " & lngPickCount & " Colors Picked"
    frmTray.Visible = True
End Sub
