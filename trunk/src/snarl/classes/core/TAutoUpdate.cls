VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents myUpdateCheck As BWebRequest
Attribute myUpdateCheck.VB_VarHelpID = -1
Dim mManual As Boolean

Private Sub Class_Initialize()
'//
End Sub

Private Sub Class_Terminate()

    If Not (myUpdateCheck Is Nothing) Then _
        myUpdateCheck.Abort

End Sub

Public Function Check(ByVal OverrideLastCheck As Boolean, ByVal PathToUpdateInfo As String) As Boolean

    If Not (myUpdateCheck Is Nothing) Then
        g_Debug "TAutoUpdate.Check(): already in progress..."
        Exit Function

    End If

Dim n As Long

    mManual = OverrideLastCheck

    If Not OverrideLastCheck Then
        n = Abs(DateDiff("d", gPrefs.last_update_check, Now))
        If n < 1 Then
            g_Debug "TAutoUpdate.Check(): last check was less than a day ago"
            Exit Function

        End If

    Else
        g_Debug "TAutoUpdate.Check(): forcing a check..."

    End If

    ' /* initiate the async operation */

    Set myUpdateCheck = New BWebRequest
    Check = myUpdateCheck.GetContent(PathToUpdateInfo)

End Function

Private Sub uDecodeUpdateFile(ByVal Content As String)
Dim szURL As String
Dim sz As String
Dim i As Long

    Content = Replace$(Content, vbCrLf, Chr$(10))       '// avoid unix/win encoding clashes

    g_Debug "TAutoUpdate.uDecodeUpdateFile(): #" & Content & "#"

    With New CConfFile
        If .SetFromText(Content, Chr$(10)) Then
            ' /* check it's the right file */
            If .FindEntry("app", sz) Then
                If sz = "Snarl" Then
                    ' /* get the version number */
                    If .FindEntry("release", sz) Then
                        ' /* is it newer than this one? */
                        If uIsNewer(sz) Then
                            sz = ""

                            ' /* get the update file name */
                            .FindEntry "setup", sz
                            gUpdateFilename = sz

                            ' /* get the update friendly name */
                            sz = ""
                            .FindEntry "friendly", sz
                            .FindEntry "url", szURL

                            If (gUpdateFilename = "") Or (sz = "") Or (szURL = "") Then
                                g_Debug "TAutoUpdate.uDecodeUpdateFile(): missing filename, friendly name or URL", LEMON_LEVEL_CRITICAL
                                Exit Sub

                            End If

                            ' /* get the change log */

'                            Set mGetFileRequest = New CHTTPRequest
'                            mGetFileRequest.GetFile "http://www.fullphat.net/updateinfo/snarl.changelog", g_MakePath(App.Path) & ".changelog"
'
'                            Do While Not mGetFileRequest.Completed
'                                Sleep 10
'                                DoEvents
'
'                            Loop

'                            Set mPanel = New TBetaPanel
'                            mPanel.Go True, mPage.Panel.hWnd, sz
'                            uDecodeUpdateFile = True

                            g_Debug "TAutoUpdate.uDecodeUpdateFile(): update found - version = " & sz, LEMON_LEVEL_INFO

                            i = g_PrivateNotify(SNARL_CLASS_GENERAL, _
                                                "Update Available", _
                                                sz & " is available for download (you have " & App.Comments & ")." & vbCrLf & vbCrLf & "Click this notification to go to the download website.", _
                                                40, _
                                                g_MakePath(App.Path) & "etc\icons\snarl-update.png", , szURL)

                            Exit Sub

                        Else
                            g_Debug "TAutoUpdate.uDecodeUpdateFile(): this version is the most current"

                            If mManual Then _
                                i = g_PrivateNotify(SNARL_CLASS_GENERAL, "Good news!", "You are running the latest version of Snarl.")

                            Exit Sub

                        End If
                    Else
                        g_Debug "TAutoUpdate.uDecodeUpdateFile(): 'release' tag not found", LEMON_LEVEL_CRITICAL

                    End If
                Else
                    g_Debug "TAutoUpdate.uDecodeUpdateFile(): mismatched app name '" & sz & "'", LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "TAutoUpdate.uDecodeUpdateFile(): 'app' tag not found", LEMON_LEVEL_CRITICAL

            End If
        Else
            g_Debug "TAutoUpdate.uDecodeUpdateFile(): bad file format", LEMON_LEVEL_CRITICAL

        End If

    End With

'    MsgBox "No update is available", vbInformation Or vbOKOnly, App.Title

End Sub

Private Function uIsNewer(ByVal Version As String) As Boolean
Dim i As Long

    i = InStr(Version, ".")
    If i = 0 Then _
        Exit Function

Dim iv As Long
Dim ir As Long

    iv = Val(g_SafeLeftStr(Version, i - 1))
    ir = Val(g_SafeRightStr(Version, Len(Version) - i))

    Debug.Print "Current: " & App.Major & "." & App.Revision & " New: " & iv & "." & ir
    uIsNewer = (iv > App.Major) Or ((iv = App.Major) And (ir > App.Revision))

End Function

Private Sub myUpdateCheck_Completed(ByVal WasSuccessful As Boolean)

    If WasSuccessful Then
        ' /* async transfer of "snarl.updateinfo" succeeded */
        g_Debug "TAutoUpdate.myUpdateCheck_Completed(): updateinfo download completed", LEMON_LEVEL_INFO
        ' /* it may have completed, but is there an update available? */
        uDecodeUpdateFile myUpdateCheck.Content
        Set myUpdateCheck = Nothing

    Else
        g_Debug "TAutoUpdate.myUpdateCheck_TimedOut()", LEMON_LEVEL_PROC
        ' /* for now we do nothing else except reset the http request */
        Set myUpdateCheck = Nothing

    '    MsgBox "Couldn't connect to the update server.  This may be due to temporary network conditions or your computer's network configuration." & vbCrLf & vbCrLf & _
    '           "If you're connecting via a VPN or similar secure link you may need to disconnect first.", _
    '           vbCritical Or vbOKOnly, App.Title
    '
    '    uResetUpdateCheck

    End If

End Sub
