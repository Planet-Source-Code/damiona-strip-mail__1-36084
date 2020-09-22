VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmControls 
   Caption         =   "Form1"
   ClientHeight    =   570
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   2025
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   990
      Picture         =   "frmControls.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   60
      Width           =   285
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   540
      Top             =   30
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   30
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu munstart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnustop 
         Caption         =   "S&top"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconfigure 
         Caption         =   "&Configure"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents tray As CTray
Attribute tray.VB_VarHelpID = -1

Private Sub Form_Load()
    'set up and display the systray icon
    Set tray = New CTray
    tray.PicBox = Picture1
    tray.TipText = "Strip Mail"
    tray.ShowIcon
End Sub

Private Sub Form_Terminate()
    'delete the icon
    tray.DeleteIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'delete the icon
    tray.DeleteIcon
End Sub

Private Sub mnuconfigure_Click()
    Load frmConfig
    frmConfig.Show 1
End Sub

Private Sub mnuexit_Click()
    'delete the icon
    tray.DeleteIcon
    End
End Sub

Private Sub mnustop_Click()
    'disable the timer so that the mail checking wil stop
    Timer1.Enabled = False
End Sub

Private Sub munstart_Click()
Dim Username As String
Dim Password As String
Dim Domain As String

    'start checking the mail
    If connected = True Then
        Timer1.Enabled = True
    ElseIf connected = False Then
        'if is not connected then do this
        Username = readINI("Mail", "Username")
        Password = Encrypt.DecryptString(readINI("Mail", "Password"), encKey)
        Domain = readINI("Mail", "Domain")
        Logon Username, Password, Domain
        If connected = True Then
            'if we are now connected then start the timer.
            Timer1.Enabled = True
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Check_Mail
End Sub

Private Sub tray_LButtonDblClick()
On Error Resume Next
        Load frmConfig
        frmConfig.Show 1
End Sub

