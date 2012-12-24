VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAboutPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TAboutPage.cls
'/
'/  Description:    The about prefs panel page
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Dim mPage As BPrefsPage
'Dim mGetFileRequest As CHTTPRequest
'Dim mPanel As TBetaPanel

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

    With mPage
        Set pm = New CTempMsg
        pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\snarl.png"
        pm.Add "image-height", 32
        pm.Add "valign", "centre"
        .Add new_BPrefsControl("labelex", "", "Snarl " & App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ")", , , pm)

        Set pm = New CTempMsg
        pm.Add "file", g_MakePath(App.Path) & "read-me.rtf"
        Set pc = new_BPrefsControl("rtf", "rtf")
        pc.DoExCmd "load", pm
        .Add pc

        .Add new_BPrefsControl("fancytoolbar", "ftb>web_stuff", "Snarl Website|Discussion Group|Blog")

'        .Add new_BPrefsControl("fancybutton2", "go_web_site", "Snarl Website")
'        .Add new_BPrefsControl("fancybutton2", "go_forum", "Snarl Forums")
'        .Add new_BPrefsControl("fancybutton2", "go_cvs", "Snarl CVS Repository")

        Set pm = New CTempMsg
        pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\open_source.jpg"
        pm.Add "image-height", 48
        pm.Add "valign", "centre"
        .Add new_BPrefsControl("labelex", "", " Released under the Simplified BSD Licence.", , , pm)

'        .Add new_BPrefsSeparator
'        .Add new_BPrefsControl("banner", "", "Snarl Update")
'        .Add new_BPrefsControl("label", "lbl_updates", Format$(gPrefs.last_update_check, "long date"), "Last Checked:")
'        .Add new_BPrefsControl("fancybutton2", "check_update", "Check now", , , , uHasXML())

        


'            If frmAbout.bIsDownloadingUpdate() Then
'                ' /* currently downloading */
'                .Add new_BPrefsControl("fancybutton2", "go_updates", "Downloading Update", , , , False)
'                .Add new_BPrefsControl("label", "lbl_updates", "Snarl is currently downloading an update.  You will be notified when it is ready to be installed.")
'
'            Else
'                .Add new_BPrefsControl("fancybutton2", "go_updates", "Check for Update")
'                .Add new_BPrefsControl("label", "lbl_updates", " ")
'
'            End If
'        End If

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    If Control.GetName() = "ftb>web_stuff" Then
        Select Case Val(Value)
        Case 1
            ' /* site */
            ShellExecute 0, "open", "http://www.fullphat.net/", vbNullString, vbNullString, SW_SHOW

        Case 2
            ' /* forum */
            'ShellExecute 0, "open", "http://sourceforge.net/forum/?group_id=191100", vbNullString, vbNullString, SW_SHOW
            ShellExecute 0, "open", "http://groups.google.co.uk/group/snarl-discuss?hl=en", vbNullString, vbNullString, SW_SHOW

        Case 3
            ' /* blog */
            ShellExecute 0, "open", "http://www.snarl-development.blogspot.com/", vbNullString, vbNullString, SW_SHOW

        End Select

    End If

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()
    Case "go_web_site"
        ShellExecute 0, "open", "http://www.fullphat.net/", vbNullString, vbNullString, SW_SHOW

    Case "go_forum"
'        ShellExecute 0, "open", "http://sourceforge.net/forum/?group_id=191100", vbNullString, vbNullString, SW_SHOW
        ShellExecute 0, "open", "http://groups.google.co.uk/group/snarl-discuss?hl=en", vbNullString, vbNullString, SW_SHOW

    Case "go_cvs"
        ShellExecute 0, "open", "http://snarlwin.cvs.sourceforge.net/snarlwin/", vbNullString, vbNullString, SW_SHOW

'    Case "go_updates"
'        If uCheckForUpdate() Then
'            Control.SetText "Please wait..."
'            Control.SetEnabled False
'
'        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 0

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()

'    If Not (theGetUpdateInfoReq Is Nothing) Then
'        theGetUpdateInfoReq.Abort
'        Set theGetUpdateInfoReq = Nothing
'
'    End If

End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub


'Private Sub uResetUpdateCheck()
'Dim pc As BControl
'
'    If mPage.Find("go_updates", pc) Then
'        pc.SetText "Check for Updates"
'        pc.SetEnabled True
'
'    End If
'
'    Set theGetUpdateInfoReq = Nothing
'
'End Sub






Private Function uHasXML() As Boolean
Dim pServer As ServerXMLHTTP30

    On Error Resume Next
    Err.Clear
    Set pServer = New ServerXMLHTTP30
    uHasXML = (Err.Number = 0)

End Function

'Private Sub uWriteToFile(ByVal Content As String)
'Dim i As Integer
'
'    On Error Resume Next
'
'    i = FreeFile()
'    Err.Clear
'    Open g_MakePath(App.Path) & ".changelog" For Output As #i
'    If Err.Number <> 0 Then _
'        Exit Sub
'
'    Print #i, Content
'    Close #i
'
'End Sub


