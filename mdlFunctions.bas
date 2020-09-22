Attribute VB_Name = "mdlFunctions"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
Global entry$   'Passed to WritePrivateProfileString
Global iniPath$ 'Path to .ini file

Global Encrypt As New clsRC4
Public Const encKey = "stripmail"
Global objOutlook As Outlook.Application
Global objMapiName As Outlook.NameSpace
Global objMail As Outlook.MailItem
Global connected As Boolean

Function readINI(AppName$, KeyName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   readINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), iniPath$))
End Function
Function writeINI(Text As String, headr As String, KeyName As String)
Dim v%
    entry$ = Text
    v% = WritePrivateProfileString(headr, KeyName, entry$, iniPath$)
    If v% <> 1 Then MsgBox "An error occurred."
End Function
Public Function Exist(file As String, Optional fileDir As String) As Boolean
    'Checks for the existence of a file or folder.
    'Returns a Boolean value (T or F)
    On Error Resume Next
    If UCase(fileDir) = UCase("dir") Then
        If Dir(file, vbDirectory) = "" Then
            Exist = False
        Else
            Exist = True
        End If
    Else
        If Dir(file) = "" Then
            Exist = False
        Else
            Exist = True
        End If
    End If
End Function
Sub Logon(Username As String, Password As String, Optional Domain As String)
'if command line set to debug then do not resume on error
If Command <> "-debug" Then
    On Error GoTo err
End If
    'begin logon
    Set objOutlook = New Outlook.Application
    Set objMapiName = objOutlook.GetNamespace("MAPI")
    'not used fully yet, need code for adding domain... May work if not using exchange.
    objMapiName.Logon Username, Password, True, True
    connected = True
    Exit Sub
err:
    connected = False
    objMapiName.Logoff
    MsgBox "An error is preventing the application from loging on to the mail system. Please check the error log, and try again."
    writeError (CStr(err.Number) + " " + CStr(err.Description) + " " + "(Sub Logon)" + " " + CStr(Date) + " " + CStr(Time))
End Sub
Sub writeError(Text As String)
Open App.Path & "\error.log" For Append As #1
    Print #1, Text
Close #1
End Sub
