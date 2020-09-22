VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pNid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Type NOTIFYICONDATA
    lSize As Long
    hWnd As Long
    lId As Long
    lFlags As Long
    lCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private mNID As NOTIFYICONDATA

Private WithEvents mPic As PictureBox
Attribute mPic.VB_VarHelpID = -1

Public Event RButtonDown()
Public Event RButtonUp()
Public Event RButtonDblClick()
Public Event LButtonDown()
Public Event LButtonUp()
Public Event LButtonDblClick()

Private Sub Class_Initialize()

    With mNID
        .lSize = Len(mNID)
        .lCallBackMessage = WM_MOUSEMOVE
        .lFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .lId = 1&
    End With
    
End Sub

Private Sub Class_Terminate()

    DeleteIcon
    
    Set mPic = Nothing
    
End Sub

Public Property Let PicBox(ByVal PicBox As PictureBox)

    Set mPic = PicBox
    
    With mNID
        .hWnd = mPic.hWnd
        .hIcon = mPic
    End With
    
End Property

Public Property Get TipText() As String

    TipText = mNID.szTip
    
End Property

Public Property Let TipText(ByVal TipText As String)

    mNID.szTip = TipText & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, mNID

End Property

Public Function ShowIcon() As Boolean

    If mPic Is Nothing Then

        ShowIcon = False
        
    Else

        Shell_NotifyIcon NIM_ADD, mNID
        ShowIcon = True

    End If

End Function

Public Sub DeleteIcon()

    Shell_NotifyIcon NIM_DELETE, mNID

End Sub

Private Sub mPic_Change()

    mNID.hIcon = mPic
    Shell_NotifyIcon NIM_MODIFY, mNID

End Sub

Private Sub mPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static bRec As Boolean
    Dim lMsg As Long

    lMsg = X / Screen.TwipsPerPixelX

    If bRec = False Then

        bRec = True

        Select Case lMsg

            Case WM_LBUTTONDBLCLK:
                RaiseEvent LButtonDblClick

            Case WM_LBUTTONDOWN:
                RaiseEvent LButtonDown

            Case WM_LBUTTONUP:
                RaiseEvent LButtonUp

            Case WM_RBUTTONDBLCLK:
                RaiseEvent RButtonDblClick

            Case WM_RBUTTONDOWN:
                RaiseEvent RButtonDown
            
            Case WM_RBUTTONUP:
                RaiseEvent RButtonUp

        End Select

        bRec = False

    End If

End Sub
