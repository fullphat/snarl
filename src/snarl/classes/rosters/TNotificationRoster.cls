VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNotificationRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TNotificationRoster.cls
    '/
    '/  Description:    Manages on-screen, missed and delayed notifications
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Const OBJ_DATE = "7-Apr-2011"

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Dim mToken As Long
Dim mRefs As Long
Dim mMissedCount2 As Long

    ' /* lists - a notification should only ever be in one of them */
Dim mActive As BTagList
Dim mWait As BTagList
Dim mMissedX As BTagList
Dim mIconCache As BTagList          ' // R2.4.2 DR3
Dim mHistory As BTagList            ' // R2.4.2 DR3

'Enum SN_STYLE_FORWARD_FLAGS
'    SN_SFF_WHEN_BUSY = 1
'    SN_SFF_WHEN_AWAY = 2
'    SN_SFF_WHEN_NOT_AWAY_OR_BUSY = 4
'
'End Enum

Dim mNotificationCentre As TNotificationCentre

Implements MLibrary
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT

    g_Debug "TNotificationRoster.Close()", LEMON_LEVEL_PROC

    mRefs = mRefs - 1
    If mRefs > 0 Then _
        Exit Function

    ' /* run tidyup code */

'    mNotificationCentre.Hide
    Set mNotificationCentre = Nothing

Dim i As Long

    With mActive
        If .CountItems Then
            For i = .CountItems To 1 Step -1
                uAsNotification(.TagAt(i)).Quit

            Next i

            .MakeEmpty

        End If

    End With

    ' /* remove temporary icons */

    EmptyIconCache

    g_Debug "TNotificationRoster.Close(): done"

End Function

Private Sub MLibrary_Initialize()
    ' /* not currently used */
End Sub

Private Function MLibrary_Magic() As Long
    MLibrary_Magic = &H23232323
End Function

Private Function MLibrary_Open() As melon.M_RESULT

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */
        mToken = (&HC0 + (Rnd * 80)) And &HFFFFFFFC
        Set mActive = new_BTagList()
        Set mWait = new_BTagList()
        Set mMissedX = new_BTagList()
        Set mIconCache = new_BTagList()
        Set mHistory = new_BTagList()

'        mMonitorCount = g_CountMonitors()

        Set mNotificationCentre = New TNotificationCentre

    End If

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()
    ' /* not currently used */
End Sub

Private Property Get MVersionInfo_Date() As String
    MVersionInfo_Date = OBJ_DATE
End Property

Private Property Get MVersionInfo_Name() As String
    MVersionInfo_Name = "notification.roster"
End Property

Private Property Get MVersionInfo_Revision() As Long
    MVersionInfo_Revision = App.Revision
End Property

Private Property Get MVersionInfo_Version() As Long
    MVersionInfo_Version = App.Major
End Property

Friend Function Add(ByRef Info As T_NOTIFICATION_INFO, ByRef Args As BPackedData, Optional ByVal AddToHistory As Boolean = True, Optional ByVal IgnorePresence As Boolean = False) As Long

    ' /* adds a notification to the roster: should only be called by:
    '       TAlert.ShowNotification()
    '       g_DoSchemePreview()
    '       TNotification (when download complete)
    '       Missed Notifications panel (when "Display" button invoked by user)
    '
    '    Note: old-style function: returns zero on error and sets lasterror
    '
    ' */

    g_SetLastError SNARL_ERROR_FAILED
    If g_IsQuitting Then _
        Exit Function               ' // shutting down - can't do anything

    ' /* log only? */

    If g_ConfigGet("log_only") = "1" Then
        g_WriteToLog Info.Title, Info.Text
        Add = -1                    ' // V40.7 - changed from M_OK to -1 (indicates success but no token available)
        g_SetLastError SNARL_SUCCESS
        Exit Function               ' // TO-DO: determine if we should return 0 and set last_error instead

    End If

    ' /* for anything else, the class must be valid */

    If (Info.ClassObj Is Nothing) Then
        g_Debug "TNotificationRoster.Add(): missing class object", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* R2.6 - minor kludge, but saves us doing constant error checking */

    If ISNULL(Args) Then _
        Set Args = New BPackedData

    ' /* R2.5.1: can direct notifications to the current user */

    If Args.Exists("recipient") Then
        If LCase$(Args.ValueOf("recipient")) <> LCase$(g_GetUserName()) Then
            g_Debug "TNotificationRoster.Add(): notification not intended for this user", LEMON_LEVEL_INFO
            g_SetLastError SNARL_ERROR_DISCARDED
            Exit Function

        End If
    End If


'    ' /* R2.31: 'duplicates' is now a simple 'stop spam' toggle - what we do with the potential
'    '    spam is for future releases; for now we just discard it */
'
'Dim lQuantum As Long
'
'    lQuantum = g_SafeLong(g_ConfigGet("flood_quantum"))
'    If lQuantum < 100 Then _
'        lQuantum = 750
'
'    If (Info.ClassObj.GetValue("block_spam") = "1") And ((GetTickCount() - Info.ClassObj.LastNotificationGeneratedTick) < lQuantum) Then
'        ' /* R2.31 - for now, block and return a failure code.  For the future, it would be lovely to
'        '    allow more functionality (e.g. update notification, append to notification, send to missed log, etc. */
'        Debug.Print "TNotificationRoster.Add(): flooding detected (quantum=" & CStr(lQuantum) & ")"
'        g_SetLastError SNARL_ERROR_FLOODING
'        Exit Function
'
'    End If

    ' /* ===== end of error checking ===== */

    ' /* R2.31: log last time a notification was generated */

    gLastNotification = Now()
    g_SetLastError SNARL_SUCCESS

    If (Info.Actions Is Nothing) Then _
        Set Info.Actions = new_BTagList()                   ' // create list even if there are no actions defined

    ' /* R2.4.2: erase callback if url detection enabled */
    If (g_ConfigGet("no_callback_urls") = "1") And (g_IsURL(Info.DefaultAck)) Then _
        Info.DefaultAck = ""

    ' /* only record the tick of the last _generated_ notification */
    Info.ClassObj.LastNotificationGeneratedTick = GetTickCount()

    ' /* R2.3: set timestamp */
    Info.DateStamp = Now()

    ' /* R2.6: redaction */

    If (uShouldRedact(Args, Info.ClassObj)) And (Info.ClassObj.GetValue("redact-always") = "1") Then
        g_Debug "TNotificationRoster.Add(): redacting on-screen content...", LEMON_LEVEL_INFO
        Info.Text = Info.ClassObj.GetValue("redact-with")

    End If

    ' /* R2.6: callback button label */
    Info.AckButtonLabel = Args.ValueOf("callback-label")

    ' /* R2.4: overflow trimming */
    Info.Title = g_TrimLines(Info.Title)
    Info.Text = g_TrimLines(Info.Text)

'    ' /* R2.4 DR7 - can include actions in request */
'
'    If (Info.Actions Is Nothing) Then _
'        Set Info.Actions = new_BTagList()                   ' // create list even if there are no actions defined
'
'Dim sn As String
'Dim sv As String
'
'    If Not (Args Is Nothing) Then
'        With Args
'            .Rewind
'            Do While .GetNextItem(sn, sv)
'                If sn = "action" Then _
'                    uGetAction Info.Actions, sv
'
'            Loop
'
'        End With
'    End If

    ' /* 40.7 - icon can be an http:// or https:// url - if it is we initiate an
    '    async download of the icon and put the notification onto a stack while
    '    we wait for the download to complete */

Dim pz As TNotification
Dim pt As BTagItem

    If g_IsURL(Info.IconPath) Then
        g_Debug "TNotificationRoster.Add(): remote icon '" & Info.IconPath & "'"

        ' /* to-do: allow user to suppress this by specifying a maximum download wait
        '    of zero seconds - in which case use the class default icon */

        If mIconCache.Find(Info.IconPath, pt) Then
            ' /* icon is already cached */
            g_Debug "TNotificationRoster.Add(): icon found in cache (" & pt.Value & ")", LEMON_LEVEL_INFO
            Info.IconPath = pt.Value

        Else
            ' /* icon isn't already cached */
            Set pt = new_BTagItem(Info.IconPath, g_GetTempPath(True) & LCase$(g_CreateGUID(True)) & ".png")
            g_Debug "TNotificationRoster.Add(): starting async download to '" & pt.Value & "'..."
            mIconCache.Add pt

            Set pz = New TNotification
            If pz.StartAsyncIconDownload(pt, Info, Args) Then
                Info.Token = Me.GetNextToken()
                pz.SetTo SN_NS_WAITING, Info
                ' /* add to the queue */
                mWait.Add pz
                Add = pz.Info.Token
                Exit Function

            Else
                g_Debug "TNotificationRoster.Add(): remote icon specified but couldn't start icon download", LEMON_LEVEL_WARNING

            End If
        End If

    ElseIf g_IsFileURI(Info.IconPath) Then
        ' /* R2.4.2 DR3: support for "file://" protocol */
        Info.IconPath = g_PathFromURI(Info.IconPath)

    End If

    ' /* R2.4: generate the "OriginalContent" content */
    g_KludgeNotificationInfo Info

Dim pInfo As T_NOTIFICATION_INFO

    ' /* ----------------------------------------------------------------- */
    ' /* send to external destinations: redirectors, forwards, subscribers */

    LSet pInfo = Info
    If uShouldRedact(Args, pInfo.ClassObj) Then _
        pInfo.Text = pInfo.ClassObj.GetValue("redact-with")

    ' /* R2.6: redirect content */
'    If Info.ClassObj.GetValue("send_to_redirects") = "1" Then _
        uDoGlobalRedirect pInfo

    ' /* R3.0 B3: redirect */
    Info.ClassObj.DoRedirection pInfo

    ' /* R2.6: forward it */
    Info.ClassObj.DoForwarding pInfo

    ' /* R2.6: send it to subscribed computers */
    If Info.ClassObj.GetValue("send_to_subscribers") = "1" Then _
        g_SubsRoster.BroadcastToSubscribers pInfo

    ' /* ----------------------------------------------------------------- */


    ' /* ----------------------------------------------------------------- */
    '   R2.4 DR8: presence management, as follows (in order of priority):
    '       1. Make Sticky
    '       2. DND
    '       3. Away
    '       4. Active
    '
    '   End result is thus:
    '
    '       State \ Priority |  Low         Normal      High
    '       -----------------+------------------------------------
    '       Make Sticky      |  Sticky      Sticky      Sticky
    '       Busy (DND=true)  |  Discarded   <user>      Displayed
    '       Away             |  Discarded   <user>      Missed
    '
    '   Items marked <user> are user configurable.
    '
    ' */

'    Debug.Print "++ current presence: " & g_HexStr(g_GetPresence())

    If (g_IsIDE()) And (gDebugMode) Then _
        uAddToMissedList Info, Args

    ' /* sticky snarls take precendence, then DND mode setting, then Away mode setting */

Dim lWhatToDo As SN_PRESENCE_ACTIONS

    If Not IgnorePresence Then
        lWhatToDo = uApplyPresence(Info)
        g_Debug "TNotificationRoster.Add(): applied presence: " & CStr(lWhatToDo)

        If lWhatToDo <> SN_PA_DO_DEFAULT Then
            ' /* we have something specific to do with the notification */
            Select Case lWhatToDo
            Case SN_PA_LOG_AS_MISSED
                ' /* put in missed list (unless class setting says otherwise) and return the token */
    '            If Info.ClassObj.GetValue("log_during_dnd") = "1" Then
                    uAddToMissedList Info, Args
                    Add = Info.Token
                    Exit Function
    
    '            End If
    
            Case SN_PA_DO_NOTHING
                ' /* bin it */
                g_SetLastError SNARL_ERROR_CLASS_BLOCKED
                Exit Function
    
            Case SN_PA_MAKE_STICKY
                Info.Timeout = 0
    
            Case SN_PA_DISPLAY_URGENT
                Info.Priority = 1
    
            Case SN_PA_DISPLAY_NORMAL
                ' /* allow it through */
    
'            Case SN_PA_FORWARD
'                Info.StyleName = style_GetStyleName(g_ConfigGet(IIf(g_IsDND, "busy_style", "away_style")))
'                Info.SchemeName = style_GetSchemeName(g_ConfigGet(IIf(g_IsDND, "busy_style", "away_style")))
    
            End Select
    
        End If
    End If


    ' /* =============================================================================
    '
    '      if we're here we need to create an on-screen instance of the notification
    '
    '    ============================================================================= */


    ' /* R2.4: merging */

    If (Info.IntFlags And SN_NF_MERGE) Then
        ' /* is there an existing notification we can merge with? */
        If uFindMergeCandidate(Info.ClassObj, Info.Title, pz) Then
            ' /* update content */
            pz.UpdateOrMerge Args, True
            g_SetLastError SNARL_WAS_MERGED                      ' // set LastError to indicate the merge
            Add = pz.Info.Token                                 ' // return existing token
            Exit Function

        End If
    End If

    ' /* bounds/logic checking */

    With Info
        If (.Priority < 0) And (.Timeout = 0) Then
            ' /* R2.4: low-priority notifications cannot be sticky */
            .Timeout = Val(g_ConfigGet("default_duration"))

        ElseIf .Timeout = -1 Then
            ' /* V41: if timeout is -1 then use global default timeout */
            .Timeout = Val(g_ConfigGet("default_duration"))

        End If

        If .StyleName = "" Then
            .StyleName = style_GetStyleName(g_ConfigGet("default_style"))
            .SchemeName = style_GetSchemeName(g_ConfigGet("default_style"))

        End If

    End With



Dim hr As M_RESULT
Dim ar As BRect
Dim p As Long

    ' /* get the appropriate instance for the notification before we do anything else
    '    we must pass a valid style instance to CSnarlWindow, even if the style
    '    roster is invalid */

Dim psi As IStyleInstance
Dim pStyle As TStyle

    ' /* get the style from the roster here - this is guaranteed to return a usable TStyle */

    g_StyleRoster.SafeGetStyle Info.StyleName & "/" & Info.SchemeName, pStyle, psi

    ' /* R2.4: generate the "OriginalContent" content */

    g_KludgeNotificationInfo Info

    ' /* get the path to the style - we need this for style-provided icons and sounds */

'    Info.IconPath = uTranslateIconPath(Info.IconPath, pStyle.Path)

    ' /* if S_STYLE_IS_WINDOWLESS is set, just pass it to the handler */
    ' /* TO-DO: do we still need to record it? what about our return value? */

    If (pStyle.Flags And S_STYLE_IS_WINDOWLESS) Then
        g_Debug "TNotificationRoster.Add(): style is handling notification display"
'        psi.UpdateContent pStyle.CreateNotificationInfo(Info.Title, Info.Text, Info.IconPath, Info.SchemeName, IIf(Info.Priority > 0, S_NOTIFICATION_IS_PRIORITY, 0))

        ' /* uCopyTo() will return TRUE if the style asked for the notification
        '    to be passed back for on-screen display */

        If g_StyleRoster.RedirectTo("", Info) Then
            ' /* style is asking for the notification to be displayed on-screen */
            Info.StyleName = style_GetStyleName(g_ConfigGet("default_style"))
            Info.SchemeName = style_GetSchemeName(g_ConfigGet("default_style"))
            Add = Me.Add(Info, Args)

        Else
            ' /* KLUDGE: add the notification to our history */
            Set pz = New TNotification
            pz.SetTo SN_NS_ACTIVE, Info
            uAddToHistory pz
            Add = -1                ' // TEMPORARY while we figure out what best to return here

        End If

        Exit Function

    End If

    ' /* is it a single instance style?  if so, we need to check to see if we already have a
    '    a notification of this style on-screen */

    If (pStyle.Flags And S_STYLE_SINGLE_INSTANCE) Then
        If uFindbyStyleName(Info.StyleName, pz) Then
            ' /* notification found, so decide what to do next... */
            If (pStyle.Flags And S_STYLE_DONT_QUEUE) Then
                ' /* don't queue notifications - so remove the existing one and create new */
                pz.Remove

            Else
                ' /* TO DO: add to queue */
                Exit Function

            End If
        End If
    End If

    ' /* if the token hasn't been pre-allocated set it here */

    If Info.Token = 0 Then
        Info.Token = mToken
        mToken = mToken + 1

    End If

'Dim dwStyleFlags As S_STYLE_FLAGS
'
'    dwStyleFlags = S_STYLE_MULTI_INSTANCE
'
'    If Not (pStyle Is Nothing) Then _
'        dwStyleFlags = pStyle.Flags

'    Info.StyleToUse = style_GetSchemeName(Info.StyleToUse)

    ' /* create a new instance of the requested style - if this fails for any reason,
    '    we simply create an instance of the internal style */




    ' /* R2.4 DR8: can use a style which allows blocking the notification - if the style
    '    allows the notification, the default style will be used to display it */

'Dim pn As notification_info
'
'    If (pStyle.Flags And S_STYLE_WANTS_VETO) Then
'
'        psi.UpdateContent pn
'        If pn.Flags = -1 Then
'            Add = 0
'            Exit Function
'
'        Else
'            ' /* re-create the style
'            uSafeGetStyle g_ConfigGet("default_style"), pStyle, psi
'
'        End If
'
'    End If

    ' /* if S_STYLE_WANT_APP_NAME is set, prefix title with the sending app name */

    If (pStyle.Flags And S_STYLE_WANT_APP_NAME) Then
        g_Debug "TNotificationRoster.Add(): style wants application name"
        Info.Title = Info.ClassObj.App.Name & "#?" & Info.Title

    End If


    ' /* R2.4.2 DR3: if this was a low priority notification, remember it */

    If Info.Priority < 0 Then _
        Info.ClassObj.App.SetCurrentLowPriority Info.Token




    ' /* create the on-screen instance */


    ' /* =======================================================
    '
    '   For future development: it would be better if the
    '   TNotification object actually created the
    '   CSnarlWindow itself, or at least did the V42
    '   kludging for us
    '
    '    ======================================================= */

Dim psw As CSnarlWindow

'    MsgBox Info.OriginalContent

    Set psw = New CSnarlWindow
    hr = psw.Create(Info, psi, pStyle.Flags, pStyle.Path)

    If hr <> M_OK Then
        ' /* TO-DO: Create() should probably return a more useful set of errors rather than just M_OK */
        g_Debug "TNotificationRoster.Add(): couldn't create window", LEMON_LEVEL_CRITICAL
        Add = 0
        Exit Function

    End If

Dim i As Long

    ' /* figure out its location on screen */

    If Not uGetNotificationPosition(psw, Info.Position, ar, i) Then
        g_Debug "TNotificationRoster.Add(): no free space on-screen", LEMON_LEVEL_CRITICAL

        If Info.Priority < 0 Then
            ' /* low priority, so just add as missed and we're done */
            g_SetLastError SNARL_ERROR_COULD_NOT_DISPLAY
            uAddToMissedList Info, Args
            Add = Info.Token
            ' /* zap the notification window */
            psw.Remove
            Set psw = Nothing
            Exit Function
        
        Else
            Do
                ' /* remove oldest (i.e. notification at head of list) */
                If Not uHide(1) Then
                    g_SetLastError SNARL_ERROR_COULD_NOT_DISPLAY
                    uAddToMissedList Info, Args
                    Add = Info.Token
                    ' /* zap the notification window */
                    psw.Remove
                    Set psw = Nothing
                    Exit Function

                End If
            Loop While Not uGetNotificationPosition(psw, Info.Position, ar, i)
        End If
    End If


'    If Not uGetNotificationPosition(psw, Info.Position, ar, i) Then
'        g_Debug "TNotificationRoster.Add(): no free space on-screen", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_COULD_NOT_DISPLAY
'
'        ' /* V42: couldn't find suitable position - mostly likely no room left on-screen */
'
'        If Info.Priority >= 0 Then
'            ' /* normal and high priority notifications are added to missed list */
'            uAddToMissedList Info
'            Add = Info.Token
'
'        Else
'            Add = 0
'
'        End If
'
'        ' /* zap the notification window */
'
'        psw.Remove
'        Set psw = Nothing
'        Exit Function
'
'    End If


    g_Debug "TNotificationRoster.Add(): placing at " & ar.Left & "," & ar.Top & "-" & ar.Right & "," & ar.Bottom

    ' /* add it to our list */

    Set pz = New TNotification
    pz.SetTo SN_NS_ACTIVE, Info, psw, i, ar
    mActive.Add pz

    ' /* return value */

    Add = Info.Token

    ' /* add to history */

    If AddToHistory Then _
        uAddToHistory pz

    ' /* R2.4 DR8: last thing to do - play the sound */

Dim szSound As String

    If g_ConfigGet("enable_sounds") = "1" Then
        szSound = trim(Info.SndFile)
        If g_ConfigGet("use_style_sounds") = "1" Then
            If (szSound = "") Or ((szSound <> "") And (g_ConfigGet("prefer_style_sounds") = "1")) Then _
                szSound = ""        ' // get_style_sound()

        End If

        ' /* no specific or style-provided sound? */

        If szSound = "" Then _
            szSound = IIf(Info.Priority > 0, g_ConfigGet("default_priority_sound"), g_ConfigGet("default_notification_sound"))

        ' /* play what we have - if we have something */

        g_QuickPlaySound szSound

    End If

End Function

Public Function Remove(ByVal Token As Long) As Boolean

    Remove = uRemove(mActive.IndexOf(CStr(Token)))

End Function

Private Function uHide(ByVal Index As Long) As Boolean

    If (Index < 1) Or (Index > mActive.CountItems) Then
        g_Debug "TNotificationRoster.uHide(): bad index #" & CStr(Index), LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    uAsNotification(mActive.TagAt(Index)).Quit
    uAsNotification(mActive.TagAt(Index)).Unset
    mActive.Remove Index
    uHide = True

End Function

Private Function uRemove(ByVal Index As Long) As Boolean

    If (Index < 1) Or (Index > mActive.CountItems) Then
        g_Debug "TNotificationRoster.uRemove(): bad index #" & CStr(Index), LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    uAsNotification(mActive.TagAt(Index)).Unset
    mActive.Remove Index
    uRemove = True

End Function

Public Function PrivateHide(ByVal Token As Long) As Boolean
Dim pn As TNotification

    PrivateHide = mActive.Find(CStr(Token), pn)
    If PrivateHide Then _
        pn.Remove

End Function

Public Function Find(ByVal Token As Long, ByVal uID As String, ByVal AppSig As String, ByVal Password As String, ByRef Notification As TNotification) As Boolean

    ' /* searches all notification lists for the specified notification (token takes priority)
    '    and returns True if it's found, and sets Notification to it */

Dim pn As TNotification
Static i As Long

    If Token <> 0 Then
        ' /* search on token */
        i = mActive.IndexOf(CStr(Token))
        If i Then
            Set pn = mActive.TagAt(i)
            If pn.Info.ClassObj.App.Password = Password Then
                ' /* password ok */
                Set Notification = pn
                g_Debug "TNotificationRoster.Find(): token found"
                g_SetLastError SNARL_SUCCESS
                Find = True

            Else
                ' /* password failure */
                g_Debug "TNotificationRoster.Find(): token found but auth failure", LEMON_LEVEL_CRITICAL
                g_SetLastError SNARL_ERROR_AUTH_FAILURE

            End If

        Else
            ' /* search wait list */
            i = mWait.IndexOf(CStr(Token))
            If i Then
                Set pn = mWait.TagAt(i)
                If pn.Info.ClassObj.App.Password = Password Then
                    ' /* password ok */
                    Set Notification = pn
                    g_SetLastError SNARL_SUCCESS
                    Find = True

                Else
                    ' /* password failure */
                    g_SetLastError SNARL_ERROR_AUTH_FAILURE

                End If

            Else
                ' /* search missed list */
                i = mMissedX.IndexOf(CStr(Token))
                If i Then
                    Set pn = mMissedX.TagAt(i)
                    If pn.Info.ClassObj.App.Password = Password Then
                        ' /* password ok */
                        Set Notification = pn
                        g_SetLastError SNARL_SUCCESS
                        Find = True

                    Else
                        ' /* password failure */
                        g_SetLastError SNARL_ERROR_AUTH_FAILURE

                    End If

                Else
                    g_SetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND

                End If
            End If
        End If

    ElseIf uID <> "" Then
        ' /* search on UID/App-Sig pair */
        uID = AppSig & uID

        With mActive
            .Rewind
            Do While .GetNextTag(pn) = B_OK
                Debug.Print "looking for '" & uID & "' found '" & pn.AppSigAndUID & "'"
                If pn.AppSigAndUID = uID Then
                    If pn.Info.ClassObj.App.Password = Password Then
                        ' /* password ok */
                        g_Debug "TNotificationRoster.Find(): uid/app found"
                        Set Notification = pn
                        g_SetLastError SNARL_SUCCESS
                        Find = True

                    Else
                        ' /* password failure */
                        g_Debug "TNotificationRoster.Find(): uid/app found but auth failure", LEMON_LEVEL_CRITICAL
                        g_SetLastError SNARL_ERROR_AUTH_FAILURE

                    End If

                    Exit Function

                End If
            Loop

        End With

        With mWait
            .Rewind
            Do While .GetNextTag(pn) = B_OK
                If pn.AppSigAndUID = uID Then
                    If pn.Info.ClassObj.App.Password = Password Then
                        ' /* password ok */
                        Set Notification = pn
                        g_SetLastError SNARL_SUCCESS
                        Find = True

                    Else
                        ' /* password failure */
                        g_SetLastError SNARL_ERROR_AUTH_FAILURE

                    End If

                    Exit Function

                End If
            Loop

        End With

        ' /* search missed list */

        With mMissedX
            .Rewind
            Do While .GetNextTag(pn) = B_OK
                If pn.AppSigAndUID = uID Then
                    If pn.Info.ClassObj.App.Password = Password Then
                        ' /* password ok */
                        Set Notification = pn
                        g_SetLastError SNARL_SUCCESS
                        Find = True

                    Else
                        ' /* password failure */
                        g_SetLastError SNARL_ERROR_AUTH_FAILURE

                    End If

                    Exit Function

                End If
            Loop

        End With

        g_SetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND

    Else
        ' /* no token or UID */
        g_SetLastError SNARL_ERROR_ARG_MISSING

    End If

End Function

Public Function Hide(ByVal Token As Long, ByVal uID As String, ByVal AppSig As String, ByVal Password As String) As Boolean
Dim pn As TNotification

    ' /* uFind() will set LastError for us */

    If Find(Token, uID, AppSig, Password, pn) Then
        Select Case pn.State
        Case SN_NS_ACTIVE
            If pn.Info.ClassObj.GetValue("prevent_close") = "1" Then
                g_Debug "TNotificationRoster.Hide(): user has blocked notification closure", LEMON_LEVEL_WARNING
                g_SetLastError SNARL_ERROR_ACCESS_DENIED
                Exit Function

            Else
                pn.Remove
                Hide = True

            End If

        Case SN_NS_MISSED
            ' /* remove from missed list */
            mMissedX.Remove mMissedX.IndexOf(CStr(pn.Info.Token))
            frmAbout.bUpdateMissedList

        Case SN_NS_WAITING
            ' /* remove from wait list */

        End Select

    End If

End Function

Public Function IsVisible(ByVal Token As Long, ByVal uID As String, ByVal AppSig As String, ByVal Password As String) As Boolean

    ' /* uFind sets lasterror for us */
    IsVisible = Find(Token, uID, AppSig, Password, Nothing)

End Function

Public Function WasMissed(ByVal Token As Long, ByVal uID As String, ByVal AppSig As String, ByVal Password As String) As Long
Dim pn As TNotification

    If Find(Token, uID, AppSig, Password, pn) Then _
        WasMissed = CLng(pn.State = SN_NS_MISSED)

End Function

Private Function uIsFreeSpace(ByRef aRect As BRect) As Boolean
Dim pn As TNotification

    ' /* return True if passed rect doesn't overlap an existing notification */

    uIsFreeSpace = True

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            If pn.Window.Frame.Intersects(aRect) Then
                uIsFreeSpace = False
                Exit Function

            End If

        Loop
    End With

''    g_Debug "uIsFreeSpace(): testing " & aRect.Left & "," & aRect.Top & "-" & aRect.Right & "," & aRect.Bottom
''                g_Debug "blocked by item " & CStr(i) & " at " & .Left & "," & .Top & ";" & .Right & "," & .Bottom

End Function

Private Function uFindExtent(ByVal Column As Integer, ByVal RightToLeft As Boolean) As Long
Dim pn As TNotification

    uFindExtent = IIf(RightToLeft, 8192, 0)

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            uFindExtent = pn.GetExtent(Column, RightToLeft, uFindExtent)

        Loop

    End With

    g_Debug "TNotificationRoster.uFindExtent(): returning " & uFindExtent

End Function

Public Sub CloseMultiple(ByVal AppToken As Long)
Dim n As Integer
Dim i As Long

    If AppToken = 0 Then
        ' /* close them all */
        g_Debug "TNotificationRoster.CloseMultiple(): closing all..."

        If mActive.CountItems Then
            
            ' /* tell them they've been closed */
            For i = 1 To mActive.CountItems
                uAsNotification(mActive.TagAt(i)).Window.bNotifyWasClosed

            Next i

            ' /* fade all notifications out together */
            For n = 255 To 0 Step -(Val(g_ConfigGet("step_size")) * 8)
                For i = 1 To mActive.CountItems
                    uAsNotification(mActive.TagAt(i)).Window.bSetAlpha n

                Next i

            Next n

            ' /* make sure they're at zero alpha */

            For i = 1 To mActive.CountItems
                uAsNotification(mActive.TagAt(i)).Window.bSetAlpha 0

            Next i
        
            ' /* now zap them */

            For i = mActive.CountItems To 1 Step -1
                uAsNotification(mActive.TagAt(i)).Quit

            Next i

            mActive.MakeEmpty

        End If

    Else
        ' /* only close snarls with same creating app */

        g_Debug "CloseMultiple(): closing all Snarls from " & g_HexStr(AppToken) & "..."

        If mActive.CountItems Then

            ' /* tell them they're being closed */
            For i = 1 To mActive.CountItems
                With uAsNotification(mActive.TagAt(i))
                    If .Info.ClassObj.App.Token = AppToken Then _
                        .Window.bNotifyWasClosed

                End With
            Next i

            ' /* fade all related notifications out together */
            For n = 255 To 0 Step -(Val(g_ConfigGet("step_size")) * 8)
                For i = 1 To mActive.CountItems
                    With uAsNotification(mActive.TagAt(i))
                        If .Info.ClassObj.App.Token = AppToken Then _
                            .Window.bSetAlpha n

                    End With
                Next i
            Next n

            ' /* make sure they're at zero alpha */

            For i = 1 To mActive.CountItems
                With uAsNotification(mActive.TagAt(i))
                    If .Info.ClassObj.App.Token = AppToken Then _
                        .Window.bSetAlpha 0

                End With
            Next i
            
            For i = mActive.CountItems To 1 Step -1
                If uAsNotification(mActive.TagAt(i)).Info.ClassObj.App.Token = AppToken Then
                    uAsNotification(mActive.TagAt(i)).Quit
                    uRemove i

                End If
            Next i
        End If
    End If

End Sub

Public Function SetAttribute(ByVal Id As Long, ByVal WhatToChange As SNARL_ATTRIBUTES, ByVal Value As String) As M_RESULT
Dim pn As TNotification

    ' /* must return a melon error for arcane legacy reasons */

    If Not Find(Id, "", "", "", pn) Then
        g_Debug "TNotificationRoster.SetAttribute(): notification " & CStr(Id) & " not found", LEMON_LEVEL_CRITICAL
        SetAttribute = M_NOT_FOUND
        Exit Function

    End If

    ' /* */

    SetAttribute = M_OK

    Select Case WhatToChange

    Case SNARL_ATTRIBUTE_TITLE
        pn.SetTitle Value

    Case SNARL_ATTRIBUTE_TEXT
        pn.SetText Value

    Case SNARL_ATTRIBUTE_ICON
        pn.SetIcon Value

    Case SNARL_ATTRIBUTE_TIMEOUT
        pn.SetTimeout Val(Value)

    Case SNARL_ATTRIBUTE_SOUND
        g_Debug "TNotificationRoster.SetAttribute(): SNARL_ATTRIBUTE_SOUND cannot be changed", LEMON_LEVEL_WARNING
        SetAttribute = M_NOT_IMPLEMENTED

    Case SNARL_ATTRIBUTE_ACK
        pn.SetCallback Value

    Case SNARL_ATTRIBUTE_MENU
        g_Debug "TNotificationRoster.SetAttribute(): SNARL_ATTRIBUTE_MENU is now deprecated", LEMON_LEVEL_WARNING
        pn.SetMenu Value

    Case Else
        g_Debug "TNotificationRoster.SetAttribute(): WhatToChange(" & CStr(WhatToChange) & ") is unknown", LEMON_LEVEL_CRITICAL
        SetAttribute = M_INVALID_ARGS

    End Select

End Function

Public Function RealMissedCount() As Long

    RealMissedCount = mMissedCount2

End Function

Public Function HaveMissedNotifications() As Boolean

    HaveMissedNotifications = (mMissedCount2 > 0)

End Function

'Public Sub ResetMissed()
'
'    mMissedCount2 = 0
'    frmAbout.bSetTrayIcon
'
'End Sub

'Private Sub uResetMissedList()
'
'    ' /* this tags the existing missed notifications as seen and resets
'    '    the manual missed count tracker - it does NOT clear the list */
'
'    mActualMissed = 0
'
''    ' /* mark existing notifications as seen */
''
''Dim pItem As TNotification
''
''    With mMissedX
''        .Rewind
''        Do While .GetNextTag(pItem) = B_OK
''            pItem.MarkAsSeen
''
''        Loop
''
''    End With
'
'    ' /* tell the tray icon */
'
'    frmAbout.bSetTrayIcon
'
'End Sub

Public Function HitTest(ByVal x As Long, ByVal y As Long) As Long

    If mActive.CountItems = 0 Then _
        Exit Function

Dim pn As TNotification
Dim px As BPoint
Dim i As Long

    Set px = new_BPoint(x, y)

    With mActive
        For i = 1 To .CountItems
            If uAsNotification(.TagAt(i)).Window.Frame.Contains(px) Then
                HitTest = i
                Exit Function

            End If
        Next i

    End With

End Function

Public Function NotificationAt(ByVal Index As Long) As CSnarlWindow

    If (Index > 0) And (Index <= mActive.CountItems) Then _
        Set NotificationAt = uAsNotification(mActive.TagAt(Index)).Window

End Function

Private Function uGetClassFlagsStr(ByRef Class As TAlert, ByVal ForcePriority As Boolean) As String
'Dim dw As E_ALERT_CLASS_FLAGS
'
'    If Not (Class Is Nothing) Then
'        dw = Class.GetFlags()
'
'        If (dw And E_CLASS_PRIORITY_NOTIFICATION) Then _
'            uGetClassFlagsStr = uGetClassFlagsStr & "p"
'
'        If (dw And E_CLASS_NOTIFICATION_ONLY) Then _
'            uGetClassFlagsStr = uGetClassFlagsStr & "n"
'
'    End If
'
'    If ForcePriority Then _
'        uGetClassFlagsStr = uGetClassFlagsStr & "p"

End Function

Private Function uFindbyStyleName(ByVal Style As String, ByRef Notification As TNotification) As Boolean

    ' /* only searches active list */

    If (mActive.CountItems = 0) Or (Style = "") Then _
        Exit Function

    Style = LCase$(Style)

Dim pn As TNotification

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            If pn.StyleNameLCase() = Style Then
                Set Notification = pn
                uFindbyStyleName = True
                Exit Function

            End If
        Loop

    End With

End Function

Private Function uGetNotificationPosition(ByRef Window As CSnarlWindow, ByVal Position As SN_START_POSITIONS, ByRef Frame As BRect, ByRef Column As Long) As Boolean

    If (Window Is Nothing) Then _
        Exit Function

Dim pv As mfxView

    Set pv = Window.GetView()
    If (pv Is Nothing) Then _
        Exit Function

    Column = 1

Dim ps As SIZEL

    With pv
        ps.cx = .Width
        ps.cy = .Height

    End With

    ' /* TODO: needs to take into account styles which veto the dropshadow */

    ps.cx = ps.cx + (2 * Val(g_ConfigGet("margin_spacing")))
    ps.cy = ps.cy + (2 * Val(g_ConfigGet("margin_spacing")))

Dim rcwa As RECT
Dim pmi As T_MONITOR_INFO
Const MARGIN = 4

    ' /* R2.4.2 DR3: get the workarea of the specified monitor */

    g_CountMonitors

    If g_GetMonitorInfo(Val(g_ConfigGet("default_screen")), pmi) Then
        LSet rcwa = pmi.rcWorkArea

    Else
        ' /* if it's not available use the primary monitor */
        g_GetWorkArea rcwa

    End If

    g_RectInset rcwa, MARGIN, MARGIN

    ' /* determine starting position */

    Set Frame = new_BRect(0, 0, ps.cx - 1, ps.cy - 1)

    If Position = SN_SP_DEFAULT_POS Then _
        Position = Val(g_ConfigGet("default_position"))

    Select Case Position

    Case SN_SP_TOP_LEFT
        Frame.OffsetBy rcwa.Left, rcwa.Top

    Case SN_SP_TOP_RIGHT
        Frame.OffsetBy rcwa.Right - ps.cx, rcwa.Top

    Case SN_SP_BOTTOM_LEFT
        Frame.OffsetBy rcwa.Left, rcwa.Bottom - ps.cy

    Case Else
        ' /* bottom right */
        Frame.OffsetBy rcwa.Right - ps.cx, rcwa.Bottom - ps.cy

    End Select

    ' /* loop looking for free space */

Dim p As Long
Const STEP_SIZE = 8

    Do While Not uIsFreeSpace(Frame)

        If (Position = SN_SP_TOP_LEFT) Or (Position = SN_SP_TOP_RIGHT) Then
            ' /* notifications run from top to bottom */

            Frame.OffsetBy 0, STEP_SIZE
            If Frame.Bottom > rcwa.Bottom Then
                ' /* notification would extend past bottom of work area so move to top of next column... */

                Frame.Top = rcwa.Top
                Frame.Bottom = Frame.Top + ps.cy

                If Position = SN_SP_TOP_RIGHT Then
                    ' /* ...which is to the left */
                    p = uFindExtent(Column, True) - STEP_SIZE
                    Frame.Right = p
                    Frame.Left = Frame.Right - ps.cx

                    If Frame.Left < rcwa.Left Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                Else
                    ' /* ...which is to the right */
                    p = uFindExtent(Column, False) + STEP_SIZE
                    Frame.Left = p
                    Frame.Right = Frame.Left + ps.cx

                    If Frame.Right > rcwa.Right Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                End If

                Column = Column + 1     ' // column count always increments

            End If

        Else
            ' /* notifications run from bottom to top */

            Frame.OffsetBy 0, -STEP_SIZE
            If Frame.Top < rcwa.Top Then
                ' /* notification would extend past top of work area so move to bottom of next column... */

                Frame.Top = rcwa.Bottom - ps.cy
                Frame.Bottom = Frame.Top + ps.cy

                If Position = SN_SP_BOTTOM_RIGHT Then
                    ' /* ...which is to the left */
                    p = uFindExtent(Column, True) - STEP_SIZE
                    Frame.Right = p
                    Frame.Left = Frame.Right - ps.cx

                    If Frame.Left < rcwa.Left Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                Else
                    ' /* ...which is to the right */
                    p = uFindExtent(Column, False) + STEP_SIZE
                    Frame.Left = p
                    Frame.Right = Frame.Left + ps.cx

                    If Frame.Right > rcwa.Right Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                End If

                Column = Column + 1

            End If

        End If
    Loop

    ' /* for 2.1+ style can veto position */

Dim pt As POINTAPI

    pt.x = Frame.Left
    pt.y = Frame.Top

    Window.AdjustPosition pt.x, pt.y

    Frame.Left = pt.x
    Frame.Top = pt.y

    uGetNotificationPosition = True

End Function

'Private Sub uSafeGetStyle(ByVal StyleToUse As String, ByRef Style As TStyle, ByRef Instance As IStyleInstance)
'
'    ' /* try to create a new TStyle from the provided Style name */
'
'    If Not (g_StyleRoster Is Nothing) Then _
'        g_StyleRoster.Find style_GetStyleName(StyleToUse), Style
'
'    ' /* didn't work? use a default one */
'
'    If (Style Is Nothing) Then _
'        Set Style = New TStyle
'
'    ' /* create a new instance from the provided Style */
'
'    Set Instance = Style.CreateInstance()
'
'    ' /* didn't work? use an internal style instance */
'
'    If (Instance Is Nothing) Then _
'        Set Instance = New TInternalStyle
'
'End Sub

Private Function uFindMergeCandidate(ByRef Class As TAlert, ByVal Title As String, ByRef Notification As TNotification) As Boolean

    ' /* only active (visible) notifications can be considered here */

    If (Class Is Nothing) Or (mActive.CountItems = 0) Then _
        Exit Function

Dim pn As TNotification

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            If pn.CanMerge(Class, Title) Then
                Set Notification = pn
                uFindMergeCandidate = True
                Exit Function

            End If
        
        Loop

    End With

End Function

Public Function AddAction(ByVal Token As Long, ByRef Data As BPackedData) As Long

    g_Debug "TNotificationRoster.AddAction()", LEMON_LEVEL_PROC_ENTER

    ' /* param checking... */

    If (Data Is Nothing) Then
        g_Debug "missing data", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_FAILED
        Exit Function

    End If

    If (Data.ValueOf("label") = "") Or (Data.ValueOf("cmd") = "") Then
        g_Debug "TNotificationRoster.AddAction(): missing label and/or command", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

Dim pn As TNotification

    ' /* uFind sets lasterror for us */

    If Find(Token, Data.ValueOf("uid"), Data.ValueOf("app-sig"), Data.ValueOf("password"), pn) Then _
        AddAction = CLng(pn.AddAction(Data.ValueOf("label"), Data.ValueOf("cmd")))

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function ClearActions(ByVal Token As Long, ByRef Data As BPackedData) As Long
Dim pn As TNotification

    ' /* uFind sets lasterror for us */

    If Find(Token, Data.ValueOf("uid"), Data.ValueOf("app-sig"), Data.ValueOf("password"), pn) Then
        pn.ClearActions
        ClearActions = -1

    End If

End Function

Private Sub uAddToMissedList(ByRef Info As T_NOTIFICATION_INFO, ByRef Args As BPackedData)
Dim n As SN_THREE_STATE

    ' /* check class setting */

    n = g_SafeLong(Info.ClassObj.GetValue("log_in_history"))

    Select Case n
    Case SN_TS_NEVER
        Exit Sub

    Case SN_TS_APP_DECIDES
        If NOTNULL(Args) Then
            If Args.Exists("log") Then
                If Args.ValueOf("log") = "0" Then _
                    Exit Sub
            End If
        End If

    Case Else
        ' /* assume "always" */

    End Select

    If Info.Token = 0 Then
        ' /* allocate a token now */
        mToken = mToken + 1
        Info.Token = mToken

    End If

    ' /* if 1000 items in the missed list, remove the oldest */

    If mMissedX.CountItems > 999 Then _
        mMissedX.Remove 1000

    ' /* create the new node and add it */

Dim pn As TNotification

    Set pn = New TNotification
    pn.SetTo SN_NS_MISSED, Info
    mMissedX.Add pn

    mMissedCount2 = mMissedCount2 + 1
    frmAbout.bUpdateMissedList

    ' /* tell the tray icon */

    frmAbout.bSetTrayIcon

'    mNotificationCentre.AddItem Info.ClassObj.App, _
                                Info.Title, _
                                Info.Text, _
                                load_image_obj(g_TranslateIconPath(Info.IconPath, "")), _
                                Info.Token

End Sub

Public Function Update(ByVal Token As Long, ByRef Args As BPackedData) As Boolean
Dim pn As TNotification

    If Find(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"), pn) Then _
        Update = pn.UpdateOrMerge(Args, False)

End Function

Public Function Merge(ByVal Token As Long, ByRef Args As BPackedData) As Boolean
Dim pn As TNotification

    If Find(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"), pn) Then _
        Merge = pn.UpdateOrMerge(Args, True)

End Function

Public Sub CheckMissed()
Static iToken As Long

    If mMissedCount2 > 0 Then
        ' /* we have some missed notifications, so figure out what to do */

        Select Case g_SafeLong(g_ConfigGet("show_missed_notifications"))
        Case 1
            ' /* do nothing */

        Case 2
            ' /* show notification */

            If iToken Then _
                Me.Hide iToken, "", "", gSnarlPassword

            iToken = g_PrivateNotify(SNARL_CLASS_GENERAL, "While you were away...", _
                                     "You missed " & CStr(mMissedCount2) & " notification" & IIf(mMissedCount2 > 1, "s", "") & ".", _
                                     -1, _
                                     g_MakePath(App.Path) & "etc\icons\snarl.png", _
                                     , _
                                     "!missed")

        Case 3
            ' /* show notification centre */
            mNotificationCentre.Show

        End Select

        mMissedCount2 = 0               ' // reset missed count...
        frmAbout.bSetTrayIcon           ' // update the tray icon

    End If

End Sub

Private Function uAsNotification(ByRef Obj As BTagItem) As TNotification

    Set uAsNotification = Obj

End Function

Public Function MissedList() As BTagList

    Set MissedList = mMissedX

End Function

'Public Sub UpdateMissedPanel()
'
'    If Not (mMissedPanel Is Nothing) Then _
'        mMissedPanel.UpdateList
'
'End Sub

Public Sub CloseMostRecent()
Dim pn As TNotification

    If mActive.CountItems > 0 Then
        Set pn = mActive.TagAt(mActive.CountItems)
        pn.Remove

    End If

End Sub

Public Function ActiveList() As BTagList

    Set ActiveList = mActive

End Function

Public Function WaitList() As BTagList

    Set WaitList = mWait

End Function

Public Sub EmptyIconCache()
Dim pt As BTagItem

    With mIconCache
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            DeleteFile pt.Value

        Loop

        .MakeEmpty

    End With

End Sub

Public Function History() As BTagList

    Set History = mHistory

End Function

Private Sub uDoGlobalRedirect(ByRef Content As T_NOTIFICATION_INFO)
Dim fPres As SN_REDIRECTION_FLAGS
Dim pt As BTagItem

    With gGlobalRedirectList
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            fPres = Val(pt.Value)
            If ((fPres And SN_RF_WHEN_AWAY) And (g_IsAway())) Or ((fPres And SN_RF_WHEN_BUSY) And (g_IsDND())) Or ((fPres And SN_RF_WHEN_ACTIVE) And (g_IsUserActive())) Then
                g_StyleRoster.RedirectTo pt.Name, Content

            End If
        Loop

    End With

End Sub

Private Sub uAddToHistory(ByRef Notification As TNotification)
Dim n As SN_THREE_STATE

    ' /* check class setting */

    n = g_SafeLong(Notification.Info.ClassObj.GetValue("log_in_history"))

    Select Case n
    Case SN_TS_NEVER
        Exit Sub

    Case SN_TS_APP_DECIDES
        ' /* if notification includes "log=0" or its class includes "log=0" then we don't log */
        If Notification.PackedArgs.Exists("log") Then
            If g_SafeLong(Notification.PackedArgs.ValueOf("log")) = 0 Then _
                Exit Sub

        ElseIf Notification.Info.ClassObj.AppProvidedSettings.Exists("log") Then
            ' /* 44.51: class object can now provide log setting */
            If g_SafeLong(Notification.Info.ClassObj.AppProvidedSettings.ValueOf("log")) = 0 Then _
                Exit Sub

        End If

    Case Else
        ' /* assume "always" */
    
    End Select

    ' /* keep history below 256 items */

    If mHistory.CountItems > 255 Then _
        mHistory.Remove 1

    mHistory.Add Notification
    frmAbout.bUpdateHistoryList

End Sub

Public Function GetNextToken() As Long

    GetNextToken = mToken
    mToken = mToken + 1

End Function

'Public Function IconCache() As BTagList
'
'    Set IconCache = mIconCache
'
'End Function

'Private Sub uDoClassPassThrough(ByRef Info As T_NOTIFICATION_INFO)
'
'    If (Info.ClassObj.RedirectList Is Nothing) Then _
'        Exit Sub
'
'Dim pCopyInfo As T_NOTIFICATION_INFO
'Dim fConditions As SN_REDIRECTION_FLAGS
'Dim pt As BTagItem
'
'    With Info.ClassObj.RedirectList
'        .Rewind
'        Do While .GetNextTag(pt) = B_OK
'            fConditions = Val(pt.Value)
'
'            Debug.Print "TNotificationRoster.uDoClassPassThrough(): is_dnd=" & g_IsDND() & " is_away=" & g_IsAway() & " cond=" & g_HexStr(fConditions)
'
'            If (g_IsDND()) And ((fConditions And SN_RF_WHEN_BUSY) <> 0) Then
'                ' /* redirect when busy */
'                fConditions = -1
'
'            ElseIf (g_IsAway()) And ((fConditions And SN_RF_WHEN_AWAY) <> 0) Then
'                ' /* redirect when away */
'                fConditions = -1
'
'            ElseIf (fConditions And SN_RF_WHEN_ACTIVE) <> 0 Then
'                ' /* redirect when active */
'                fConditions = -1
'
'            Else
'                ' /* no shadowing */
'                fConditions = 0
'
'            End If
'
'            If fConditions <> 0 Then
'                g_Debug "TNotificationRoster.uDoClassPassThrough(): passing to '" & pt.Name & "'..."
'                LSet pCopyInfo = Info
'                pCopyInfo.StyleName = LCase$(style_GetStyleName(pt.Name))
'                pCopyInfo.SchemeName = LCase$(style_GetSchemeName(pt.Name))
'                uCopyTo pCopyInfo
'
'            End If
'        Loop
'
'    End With
'
'End Sub

Private Function uApplyPresence(ByRef Info As T_NOTIFICATION_INFO) As SN_PRESENCE_ACTIONS

    If g_ConfigGet("sticky_snarls") = "1" Then
        g_Debug "TNotificationRoster.uApplyPresence(): Sticky mode enabled"
        ' /* R2.4 DR8: *all* notifications are sticky */
        Info.Timeout = 0

    ElseIf g_IsDND() Then
        ' /* Do Not Disturb (Busy) mode */
        g_Debug "TNotificationRoster.uApplyPresence(): Busy (aka DND) mode enabled: priority=" & CStr(Info.Priority)
        If Info.Priority < 0 Then
            ' /* low priority = ignored */
            g_SetLastError SNARL_ERROR_DO_NOT_DISTURB
            uApplyPresence = SN_PA_DO_NOTHING

        ElseIf Info.Priority = 0 Then
            ' /* determine what to do with normal priority notifications */
            uApplyPresence = Val(Info.ClassObj.GetValue("override-busy")) - 1
            If uApplyPresence = SN_PA_DO_DEFAULT Then _
                uApplyPresence = Val(g_ConfigGet("busy_mode"))

        End If

    ElseIf g_IsAway() Then
        ' /* Away mode */
        g_Debug "TNotificationRoster.uApplyPresence(): Away mode enabled: priority=" & CStr(Info.Priority)
        If Info.Priority < 0 Then
            ' /* low priority = binned */
            g_SetLastError SNARL_ERROR_CLASS_BLOCKED
            uApplyPresence = SN_PA_DO_NOTHING

        ElseIf Info.Priority = 0 Then
            ' /* determine what to do with normal priority notifications */
            uApplyPresence = Val(Info.ClassObj.GetValue("override-away")) - 1
            Debug.Print "### away class setting is: " & uApplyPresence
            If uApplyPresence = SN_PA_DO_DEFAULT Then _
                uApplyPresence = Val(g_ConfigGet("away_mode"))

            Debug.Print "### away setting is: " & uApplyPresence

        Else
            ' /* high priority = logged as missed */
            uApplyPresence = SN_PA_LOG_AS_MISSED

        End If

    Else
        uApplyPresence = Val(Info.ClassObj.GetValue("override-normal")) - 1

    End If

End Function

'Public Sub MarkMissed()
''Dim pItem As TNotification
''
''    ' /* mark existing notifications as seen */
''
''    With mMissedX
''        .Rewind
''        Do While .GetNextTag(pItem) = B_OK
''            pItem.MarkAsSeen
''
''        Loop
''
''    End With
'
'End Sub

Public Function CountActiveBySignature(ByVal Signature As String) As Long
Dim pn As TNotification

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            If pn.Info.ClassObj.App.Signature = Signature Then _
                CountActiveBySignature = CountActiveBySignature + 1

        Loop
    End With

End Function

Public Function NC() As TNotificationCentre

    Set NC = mNotificationCentre

End Function

Private Function uShouldRedact(ByRef Args As BPackedData, ByRef Class As TAlert) As Boolean

    If ISNULL(Args) Then _
        Exit Function

Dim l As Long

    l = g_SafeLong(Args.ValueOf("sensitivity"))
    If l = 0 Then _
        Exit Function

    Debug.Print "uShouldRedact(): sensitivity: " & CStr(l) & "; redact at: " & Class.GetValue("redact-when")
    uShouldRedact = (l >= g_SafeLong(Class.GetValue("redact-when")))

End Function
