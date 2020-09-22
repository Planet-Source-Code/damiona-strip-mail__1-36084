VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Strip Mail - (Config Window)"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   3630
      TabIndex        =   13
      Top             =   570
      Width           =   950
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   255
      Left            =   3630
      TabIndex        =   12
      Top             =   270
      Width           =   950
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   2160
      Width           =   405
   End
   Begin VB.TextBox txtSave 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   930
      TabIndex        =   9
      Top             =   1770
      Width           =   2595
   End
   Begin VB.TextBox txtShell 
      Height          =   285
      Left            =   930
      TabIndex        =   7
      Top             =   2145
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mail Options"
      Height          =   1515
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   3405
      Begin VB.TextBox txtDomain 
         Height          =   285
         Left            =   1020
         TabIndex        =   3
         Top             =   1095
         Width           =   2265
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1020
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   690
         Width           =   2265
      End
      Begin VB.TextBox txtUsername 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1020
         TabIndex        =   1
         Top             =   270
         Width           =   2265
      End
      Begin VB.Label Label3 
         Caption         =   "Domain:"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   1110
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   705
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Save Path:"
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   1785
      Width           =   825
   End
   Begin VB.Label Label4 
      Caption         =   "Shell:"
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   2160
      Width           =   465
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
On Error GoTo err
    'common dialog set to find exe files
    frmControls.cd1.Filter = "Executable Files(*.exe)|*.exe"
    frmControls.cd1.InitDir = App.Path
    frmControls.cd1.Filename = "*.exe"
    frmControls.cd1.ShowOpen
    txtShell.Text = frmControls.cd1.Filename
    Exit Sub
err:
    If err.Number <> 32755 Then
        writeError (CStr(err.Number) + " " + CStr(err.Description) + " " + CStr(Me.Caption) + " Ok Click " + CStr(Date) + " " + CStr(Time))
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
'write the info to the ini file
    If txtUsername.Text = "" Or txtPassword.Text = "" Or txtSave.Text = "" Then
        'error if not all required fields are filled out.
        MsgBox "All the yellow fields are required", vbOKOnly, "Strip Mail - Missing Information"
        Exit Sub
    End If
    'write the information to the ini file.
    writeINI txtUsername, "Mail", "Username"
    writeINI Encrypt.EncryptString(txtPassword.Text, encKey), "Mail", "Password"
    writeINI txtSave.Text, "Settings", "SavePath"
    If txtDomain.Text <> "" Then
        writeINI txtDomain.Text, "Mail", "Domain"
    End If
    If txtShell.Text <> "" Then
        writeINI txtShell, "Settings", "Shell"
    End If
    'done writing to ini now unload me
    Unload Me
End Sub

Private Sub Form_Load()
    If Exist(App.Path + "\mail.ini") = True Then
        'load all the form values
        txtUsername.Text = readINI("Mail", "Username")
        txtPassword.Text = Encrypt.DecryptString(readINI("Mail", "Password"), encKey)
        txtDomain.Text = readINI("Mail", "Domain")
        txtSave.Text = readINI("Settings", "SavePath")
        txtShell.Text = readINI("Settings", "Shell")
    End If
End Sub
