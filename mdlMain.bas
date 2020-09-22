Attribute VB_Name = "mdlMain"
'**********************************************************
'                    Strip Mail v1.0
'
'   Developed By:   Damion Allen
'   Purpose:        Strip attachments off of email, and
'   shell out to another proggie to do processing of that
'   file.
'
'**********************************************************
Option Explicit
Dim SavePath As String
Dim ShellProg As String

Sub main()
'Begin declaring
If Command <> "-debug" Then
    On Error GoTo err
End If
Dim Username As String
Dim Password As String
Dim Domain As String

    iniPath$ = App.Path + "\mail.ini"
    If Exist(App.Path + "\mail.ini") = False Then
        'if the ini file not there then setup config info.
        MsgBox "You must set up the Strip Mail configuration", vbOKOnly, "Strip Mail - Setup"
        Load frmConfig
        frmConfig.Show 1
    End If
        Username = readINI("Mail", "Username")
        Password = Encrypt.DecryptString(readINI("Mail", "Password"), encKey)
        Domain = readINI("Mail", "Domain")
        SavePath = readINI("Settings", "SavePath")
        ShellProg = readINI("Settings", "Shell")
        Logon Username, Password, Domain
        Load frmControls
        If connected = True Then
            Check_Mail
            frmControls.Timer1.Enabled = True
        Else
            MsgBox "There was an error preventing the program to connect, please check the error log file, and try again."
        End If
        Exit Sub
err:
    Load frmControls
    writeError (CStr(err.Number) + " " + CStr(err.Description) + " " + "(Sub Main)" + " " + CStr(Date) + " " + CStr(Time))
End Sub
Sub Check_Mail()
Dim intCount As Integer
Dim tmp As String
Dim i As Integer
Dim l As Integer
Dim oFolder As Outlook.MAPIFolder

'debug let errors show
If Command <> "-debug" Then
    On Error Resume Next
End If
    'if connected then begin looking for new mail
    If connected = True Then
        frmControls.Timer1.Enabled = False
        intCount = objMapiName.GetDefaultFolder(olFolderInbox).UnReadItemCount
        Set oFolder = objMapiName.GetDefaultFolder(olFolderInbox)
        'if new mail, then look for attachments
        If intCount > 0 Then
            For Each objMail In oFolder.Items
                With objMail
                'Check for new mail with attachments
                    If .UnRead = True And .Attachments.Count > 0 Then
                        For i = 1 To .Attachments.Count
                            'if attachemnt then save it to specified path
                            .Attachments.Item(i).SaveAsFile SavePath + .Attachments.Item(i).DisplayName
                            If ShellProg <> "" Then
                                'after attachment is saved then shell it to a program
                                If InStr(UCase(ShellProg), "<SAVEPATH>") Then
                                    'allows the use of the save path string in the config setup
                                    For l = 1 To Len(ShellProg)
                                        If Mid(ShellProg, l, 1) <> "<" Then
                                            tmp = tmp + Mid(ShellProg, l, 1)
                                        Else
                                            If Mid(SavePath, Len(SavePath) - 1, 1) <> "\" Then
                                                SavePath = SavePath + "\"
                                            End If
                                                tmp = tmp + SavePath + .Attachments.Item(i).DisplayName
                                            l = l + 9
                                        End If
                                    Next l
                                Else
                                    tmp = .Attachments.Item(i).DisplayName
                                End If
                                'shell now
                                Shell (tmp)
                            End If
                        Next i
                        'change the status of the mail to unread
                        .UnRead = False
                    End If
                End With
            Next
        End If
        frmControls.Timer1.Enabled = True
    End If
End Sub
