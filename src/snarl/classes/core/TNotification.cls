VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TNotification.cls
'/
'/  Description:    Represents a notification within the notification roster
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

'// "http://www.google.com/ig/images/weather/chance_of_rain.gif"

'    Window As CSnarlWindow
'    Column As Integer           ' // used to help determine placement of other notifications
'    StyleName As String         ' // R2.31: expanded
'    SchemeName As String        ' // R2.31: expanded
'    Info As T_NOTIFICATION_INFO ' // R2.4 DR7: full originating information

Dim mWindow As CSnarlWindow
Dim mColumn As Integer          ' // used to help determine placement of other notifications
'Dim mStyleName As String        ' // R2.31: expanded
'Dim mSchemeName As String       ' // R2.31: expanded

Dim mInfo As T_NOTIFICATION_INFO
Dim mTempIcon As String
Dim mArgs As BPackedData

Public Enum SN_NOTIFICATION_STATUS
    SN_NS_ACTIVE = 1
    SN_NS_WAITING = 2
    SN_NS_MISSED = 3

End Enum

Dim mState As SN_NOTIFICATION_STATUS
Dim mReplayed As Boolean
Dim mIcon As MImage

Dim WithEvents theIconRequest As BWebRequest
Attribute theIconRequest.VB_VarHelpID = -1

Dim mPacked As BPackedData

Implements BTagItem

Friend Function StartAsyncIconDownload(ByRef RemoteIcon As BTagItem, ByRef Info As T_NOTIFICATION_INFO, ByRef Args As BPackedData) As Boolean

    On Error Resume Next

    err.Clear
    Set theIconRequest = New BWebRequest
    If err.Number <> 0 Then
        g_Debug "TNotification.StartAsyncIconDownload(): couldn't create BWebRequest", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    LSet mInfo = Info
    mTempIcon = RemoteIcon.Value
    Set mArgs = Args
    mInfo.IntFlags = mInfo.IntFlags Or SN_NF_TEMP_ICON

    g_Debug "TNotification.uStartAsyncIconDownload(): saving icon as '" & mTempIcon & "'..."
    StartAsyncIconDownload = theIconRequest.GetFile(RemoteIcon.Name, mTempIcon)

End Function

Friend Sub SetTo(ByVal State As SN_NOTIFICATION_STATUS, ByRef Info As T_NOTIFICATION_INFO, Optional ByRef Window As CSnarlWindow, Optional ByVal Column As Integer, Optional ByRef Frame As BRect)   ', Optional ByVal StyleName As String, Optional ByVal SchemeName As String)

    LSet mInfo = Info
    uFixupContent

    Set mWindow = Window
    mColumn = Column
'    mStyleName = StyleName
'    mSchemeName = SchemeName

    mState = State
    Set mIcon = load_image_obj(g_TranslateIconPath(Info.IconPath, ""))

    If Not (mWindow Is Nothing) Then
        mWindow.Window.MoveTo Frame.Left, Frame.Top
        mWindow.Show

    End If

End Sub

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mInfo.Token)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mInfo.ClassObj.App.Signature & mInfo.CustomUID

End Function

Private Sub theIconRequest_Completed(ByVal WasSuccessful As Boolean)

    g_Debug "TNotification.theIconRequest_Completed()", LEMON_LEVEL_PROC_ENTER

    g_Debug "removing from wait list..."
    g_Debug "index is " & g_NotificationRoster.WaitList.IndexOf(CStr(mInfo.Token))
    g_NotificationRoster.WaitList.Remove g_NotificationRoster.WaitList.IndexOf(CStr(mInfo.Token))

Dim pIcon As MImage
Dim sz As String

    If Not WasSuccessful Then
        ' /* didn't download for some reason so use the class default icon, if one exists */
        g_Debug "failed: will use class default (if one exists)"
        If Not (mInfo.ClassObj Is Nothing) Then
            sz = mInfo.ClassObj.DefaultIcon

        Else
            ' /* to-do: allow user to specify an icon? */

        End If

    Else
        g_Debug "success: creating squared-up icon..."
        Set pIcon = style_MakeSquareImage(load_image_obj(mTempIcon))
        If NOTNULL(pIcon) Then
            g_Debug "success: writing as PNG image..."
            With New mfxView
                .SizeTo pIcon.Width, pIcon.Height
                .Clear
                .DrawScaledImage pIcon
                .WriteToFile mTempIcon, "image/png"

            End With
            sz = mTempIcon

        Else
            g_Debug "failed", LEMON_LEVEL_WARNING

        End If
    End If

    g_Debug "icon is now " & g_Quote(sz)
    mInfo.IconPath = sz
    
    g_Debug "fixing up content..."
    uFixupContent

    g_Debug "adding to notification roster..."
    g_NotificationRoster.Add mInfo, mArgs

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function Window() As CSnarlWindow

    Set Window = mWindow

End Function

Public Function Column() As Integer

    Column = mColumn

End Function

Friend Function Info() As T_NOTIFICATION_INFO

    LSet Info = mInfo

End Function

Public Function StyleName() As String

    StyleName = mInfo.StyleName

End Function

Public Function StyleNameLCase() As String

    StyleNameLCase = LCase$(mInfo.StyleName)

End Function

'Private Sub uFixupContent(ByRef Info As T_NOTIFICATION_INFO)
'
'    If (Info.ClassObj Is Nothing) Then
'        Debug.Print "########### uFixupContent missing class object ###########"
'        Exit Sub
'
'    End If
'
'    With Info
'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)
'
'    End With
'
'Dim ppn As BPackedData
'
'    ' /* new */
'    Set ppn = New BPackedData
'
'    With ppn
'        .Add "id", Info.ClassObj.Name
'        .Add "title", Info.Title
'        .Add "text", Info.Text
'        .Add "timeout", CStr(Info.Timeout)
'        .Add "icon", Info.IconPath
'        .Add "priority", CStr(Info.Priority)
'        .Add "callback", Info.DefaultAck
'        .Add "value", Info.Value
'
''        If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
'            .Add "merge", "1"
'
'    End With
'
'Dim ppd As BPackedData
'Dim szn As String
'Dim szv As String
'
'    ' /* add in all other content */
'
'    Set ppd = New BPackedData
'    With ppd
'        If .SetTo(Info.OriginalContent) Then
'            .Rewind
'            Do While .GetNextItem(szn, szv)
'                If Not ppn.Exists(LCase$(szn)) Then _
'                    ppn.Add szn, szv
'
'            Loop
'        End If
'    End With
'
'    Info.OriginalContent = ppn.AsString()
'
'End Sub

Public Sub Unset()

    Set mWindow = Nothing

End Sub

Public Sub Remove()

    If mState = SN_NS_ACTIVE Then _
        mWindow.Hide

End Sub

Public Sub Quit()

    If Not (mWindow Is Nothing) Then _
        mWindow.Remove

End Sub

Public Function AppSigAndUID() As String

    AppSigAndUID = mInfo.ClassObj.App.Signature & mInfo.CustomUID

End Function

'Public Sub SetState(ByVal State As SN_NOTIFICATION_STATUS)
'
'    mState = State
'
'End Sub

Public Function State() As SN_NOTIFICATION_STATUS

    State = mState

End Function

Public Function GetExtent(ByVal Column As Integer, ByVal RightToLeft As Boolean, ByVal x As Long) As Long

    GetExtent = x

    ' /* in the correct column? */

    If mColumn <> Column Then _
        Exit Function

    If RightToLeft Then
        If mWindow.Frame.Left < x Then _
            GetExtent = mWindow.Frame.Left

    Else
        If mWindow.Frame.Right > x Then _
            GetExtent = mWindow.Frame.Right

    End If

End Function

Private Sub uFixupContent()

    g_KludgeNotificationInfo mInfo, mPacked

'    ' /* translates current T_NOTIFICATION_INFO content into a BPackedString
'    '    and then stores that back into T_NOTIFICATION_INFO->OriginalContent
'    '    this is so V42 styles can access the entire request */
'
'    If (mInfo.ClassObj Is Nothing) Then
'        g_Debug "TNotification.uFixupContent(): missing ClassObj", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    End If
'
'Dim ppd As BPackedData
'
'    Set ppd = New BPackedData
'
'    With mInfo
'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)
'
'        ppd.Add "id", .ClassObj.Name
'        ppd.Add "title", .Title
'        ppd.Add "text", .Text
'        ppd.Add "timeout", CStr(.Timeout)
'        ppd.Add "icon", .IconPath
'        ppd.Add "priority", CStr(.Priority)
'        ppd.Add "callback", .DefaultAck
'        ppd.Add "value", .Value
'
''        If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
'            .Add "merge", "1"
'
'    End With
'
'Dim ppx As BPackedData
'Dim szn As String
'Dim szv As String
'
'    ' /* add in all other custom content */
'
''    If Info.OriginalContent <> "" Then _
'        MsgBox Info.OriginalContent
'
'    Set ppx = New BPackedData
'    With ppx
'        If .SetTo(Info.OriginalContent) Then
'            .Rewind
'            Do While .GetNextItem(szn, szv)
'                If Not ppd.Exists(LCase$(szn)) Then _
'                    ppd.Add szn, szv
'
'            Loop
'        End If
'    End With
'
'    mInfo.OriginalContent = ppd.AsString()

End Sub

Public Function SetTitle(ByVal Title As String) As Boolean

    mInfo.Title = Title
    uFixupContent
    SetTitle = True

    Select Case mState
    Case SN_NS_ACTIVE
        mWindow.Update mInfo

    Case SN_NS_MISSED

    Case SN_NS_WAITING

    End Select

End Function

Public Function SetText(ByVal Text As String) As Boolean

    mInfo.Text = Text
    uFixupContent
    SetText = True

    Select Case mState
    Case SN_NS_ACTIVE
        mWindow.Update mInfo

    Case SN_NS_MISSED

    Case SN_NS_WAITING

    End Select

End Function

Public Function SetIcon(ByVal Icon As String) As Boolean

    mInfo.IconPath = Icon
    uFixupContent
    SetIcon = True

    If (mWindow Is Nothing) Then _
        Exit Function

    Select Case mState
    Case SN_NS_ACTIVE
        mWindow.Update mInfo

    Case SN_NS_MISSED

    Case SN_NS_WAITING

    End Select

End Function

Public Function SetTimeout(ByVal Timeout As Long) As Boolean

    Select Case mState
    Case SN_NS_ACTIVE
        mWindow.SetTimeout Timeout
        SetTimeout = True

    Case SN_NS_MISSED

    Case SN_NS_WAITING

    End Select

End Function

Public Function SetCallback(ByVal Callback As String) As Boolean

    Select Case mState
    Case SN_NS_ACTIVE
        mWindow.SetAck Callback
        SetCallback = True

    Case SN_NS_MISSED

    Case SN_NS_WAITING
        mInfo.DefaultAck = Callback

    End Select

End Function

Public Function SetMenu(ByVal Content As String) As Boolean

    Select Case mState
    Case SN_NS_ACTIVE
        mWindow.SetMenu Content
        SetMenu = True

    Case SN_NS_MISSED

    Case SN_NS_WAITING

    End Select

End Function

Public Function ClearActions() As Boolean

    ClearActions = True
    If mInfo.Actions.CountItems = 0 Then _
        Exit Function

    mInfo.Actions.MakeEmpty

    Select Case mState
    Case SN_NS_ACTIVE
        mWindow.RethinkActions

    Case SN_NS_MISSED
        ' /* need to update the missed notifications panel, if it's open */

    Case SN_NS_WAITING

    End Select

End Function

Public Function AddAction(ByVal Label As String, ByVal Command As String) As Boolean

    g_Debug "TNotification.AddAction()", LEMON_LEVEL_PROC_ENTER

Dim bCatch As Boolean
Dim dw As Long

    ' /* error checking */

    If g_SafeLeftStr(Command, 1) = "@" Then
        ' /* callback must be a signed 16-bit value */
        dw = g_SafeLong(g_SafeRightStr(Command, Len(Command) - 1), bCatch)
        If (bCatch) Or (dw < -32768) Or (dw > 32767) Or (Not g_IsNumeric(g_SafeRightStr(Command, Len(Command) - 1))) Then
            g_Debug "dynamic callback " & g_Quote(Command) & " must be UINT16", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
            g_SetLastError SNARL_ERROR_INVALID_ARG
            Exit Function

        End If
    End If

    If ISNULL(mInfo.Actions) Then _
        Set mInfo.Actions = new_BTagList()

    With mInfo.Actions
        If .IndexOf(Label) Then
            ' /* replace */
            g_Debug "replacing existing action " & g_Quote(Label) & " with " & g_Quote(Command)
            .Update Label, Command

        Else
            ' /* add new */
            g_Debug "adding new action " & g_Quote(Label) & " / " & g_Quote(Command)
            .Add new_BTagItem(Label, Command)

        End If

    End With

    Select Case mState
    Case SN_NS_ACTIVE
        g_Debug "state is SN_NS_ACTIVE: rethinking actions..."
        mWindow.RethinkActions

    Case SN_NS_MISSED
        ' /* need to update the missed notifications panel, if it's open */
        g_Debug "state is SN_NS_MISSED: updating missed list..."
        frmAbout.bUpdateMissedList

    Case SN_NS_WAITING
        g_Debug "state is SN_NS_WAITING: nothing to do"

    End Select

    AddAction = True

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function UpdateOrMerge(ByRef Args As BPackedData, ByVal ShouldMerge As Boolean) As Boolean

    Debug.Print "text was: " & mInfo.Text
    uUpdateContent Args, ShouldMerge
    Debug.Print "text now: " & mInfo.Text

    Select Case mState
    Case SN_NS_ACTIVE
        uFixupContent
        Debug.Print "text: " & mInfo.Text
        UpdateOrMerge = mWindow.Update(mInfo)


    Case SN_NS_MISSED
        ' /* refresh the missed panel if it's open */
        frmAbout.bUpdateMissedList
        UpdateOrMerge = True


    Case SN_NS_WAITING
        UpdateOrMerge = True

    End Select

End Function

Private Sub uUpdateContent(ByRef Args As BPackedData, Optional ByVal ShouldMerge As Boolean = False)
Dim sz As String

    ' /* update or merge provided information with existing
    '    content - ShouldMerge controls which */

    With mInfo

        ' /* only change something if the entry _exists_ in Args */

        If Args.Exists("title") Then
            sz = uUnescape(Args.ValueOf("title"), .Title)
            .Title = g_TrimLines(sz)

        End If

        sz = ""
        If Args.Exists("text") Then
            If ShouldMerge Then _
                sz = mInfo.Text & vbCrLf

            sz = sz & uUnescape(Args.ValueOf("text"), .Text)
            .Text = g_TrimLines(sz)

        End If

        If Args.Exists("icon") Then
            .IconPath = Args.ValueOf("icon")

        ElseIf Args.Exists("icon-base64") Then
            ' /* R2.4b4 */
            .IconPath = g_GetBase64Icon(Args.ValueOf("icon-base64"))

        ElseIf Args.Exists("icon-phat64") Then
            ' /* R2.4.2 DR3 */
            .IconPath = g_GetPhat64Icon(Args.ValueOf("icon-phat64"))

        End If

        If Args.Exists("timeout") Then _
            .Timeout = g_SafeLong(Args.ValueOf("timeout"))          ' // g_SafeLong() returns 0 on error - not good...

        ' /* following added in R2.4 Beta 4 */

        If Args.Exists("callback") Then _
            .DefaultAck = Args.ValueOf("callback")

        If Args.Exists("reply-to") Then _
            .hWndReply = g_SafeLong(Args.ValueOf("reply-to"))

        If Args.Exists("reply") Then _
            .uReplyMsg = g_SafeLong(Args.ValueOf("reply"))

        ' /* clearly this can only be set if referencing the notification by its token */

        If Args.Exists("uid") Then _
            .CustomUID = Args.ValueOf("uid")

        ' /* still TBC: Priority, Actions, Value */

        .LastUpdated = Now()


        .OriginalContent = Args.AsString()

    End With

End Sub

Public Function CanMerge(ByRef Class As TAlert, ByVal Title As String) As Boolean

    If Not (mWindow Is Nothing) Then _
        CanMerge = mWindow.IsMergeCandidate(Class, Title)

End Function

'Public Sub MarkAsReplayed()
'
'    mReplayed = True
'
'End Sub

Private Function uUnescape(ByVal Text As String, ByVal Replacement As String) As String
Dim sz As String

    sz = Text
    sz = Replace$(sz, "\+", Replacement)
    sz = Replace$(sz, "\user", g_GetUserName())
    sz = Replace$(sz, "\host", g_GetComputerName())
    sz = Replace$(sz, "\info", "Snarl " & App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ")")
    sz = Replace$(sz, "\n", vbCrLf)

    uUnescape = sz

End Function

'Friend Function Init2(ByRef Info As T_NOTIFICATION_INFO) As Long
'
'    LSet mInfo = Info
'
'    ' /* do some setting up */
'
'    With mInfo
'
'        ' /* create an empty list if no actions defined */
'        If (.Actions Is Nothing) Then _
'            Set .Actions = new_BTagList()
'
'        ' /* R2.4.2: erase callback if url detection enabled */
'        If (g_ConfigGet("no_callback_urls") = "1") And (g_IsURL(.DefaultAck)) Then _
'            .DefaultAck = ""
'
'        ' /* only record the tick of the last _generated_ notification */
'        .ClassObj.LastNotificationGeneratedTick = GetTickCount()
'
'        ' /* R2.3: set datestamp */
'        .DateStamp = Now()
'
'        ' /* R2.4: overflow trimming */
'        .Title = g_TrimLines(.Title)
'        .Text = g_TrimLines(.Text)
'
'        ' /* some bounds checking */
'        If (.Priority < 0) And (.Timeout = 0) Then
'            ' /* R2.4: low-priority notifications cannot be sticky */
'            .Timeout = Val(g_ConfigGet("default_duration"))
'
'        ElseIf .Timeout = -1 Then
'            ' /* V41: if timeout is -1 then use global default timeout */
'            .Timeout = Val(g_ConfigGet("default_duration"))
'
'        End If
'
'        ' /* R2.4: generate the "OriginalContent" content */
'        uFixupContent
'
'    End With
'
'    ' /* 40.7 - icon can be an http:// or https:// url - if it is we initiate an
'    '    async download of the icon and put the notification onto a stack while
'    '    we wait for the download to complete */
'
'    Init2 = uCheckIcon()
'    If Init2 <> 0 Then
'        Exit Function
'
'
'
'
'
'
'End Function
'
'Private Function uCheckIcon() As Long
'Dim pt As BTagItem
'
'    ' /* returns allocated token or zero if no async download required */
'
'    With mInfo
'        If g_IsURL(.IconPath) Then
'            g_Debug "TNotification.uCheckIcon(): remote icon '" & .IconPath & "'"
'
'            ' /* to-do: allow user to suppress this by specifying a maximum download wait
'            '    of zero seconds - in which case use the class default icon */
'
'            If g_NotificationRoster.IconCache.Find(.IconPath, pt) Then
'                ' /* icon is already cached */
'                g_Debug "TNotification.uCheckIcon(): icon found in cache (" & pt.Value & ")", LEMON_LEVEL_INFO
'                .IconPath = pt.Value
'
'            Else
'                ' /* icon isn't already cached */
'                Set pt = new_BTagItem(Info.IconPath, g_GetTempPath(True) & LCase$(g_CreateGUID(True)) & ".png")
'                g_Debug "TNotification.uCheckIcon(): starting async download to '" & pt.Value & "'..."
'                g_NotificationRoster.IconCache.Add pt
'
'                If pnx.StartAsyncIconDownload(pt, Args) Then
'
'                    mState = SN_NS_WAITING
'
'                    ' /* assign the token now */
'                    .Token = g_NotificationRoster.GetNextToken()
'                    uCheckIcon = .Token
'                    Exit Function
'
'                Else
'                    g_Debug "TNotification.uCheckIcon(): remote icon specified but couldn't start icon download", LEMON_LEVEL_WARNING
'
'                End If
'            End If
'
'        ElseIf g_IsFileURI(.IconPath) Then
'            ' /* R2.4.2 DR3: support for "file://" protocol */
'             .IconPath = g_PathFromURI(.IconPath)
'
'        End If
'
'    End With
'
'End Function

'Public Function ApplyPresence() As SN_PRESENCE_ACTIONS
'
'    If g_ConfigGet("sticky_snarls") = "1" Then
'        g_Debug "TNotification.ApplyPresence(): Sticky mode enabled"
'        ' /* R2.4 DR8: *all* notifications are sticky */
'        mInfo.Timeout = 0
'
'    ElseIf g_IsDND() Then
'        ' /* Do Not Disturb (Busy) mode */
'        g_Debug "TNotification.ApplyPresence(): Busy (aka DND) mode enabled: priority=" & CStr(mInfo.Priority)
'        If mInfo.Priority < 0 Then
'            ' /* low priority = ignored */
'            g_SetLastError SNARL_ERROR_DO_NOT_DISTURB
'            ApplyPresence = SN_PA_LAST_ERROR_SET
'
'        ElseIf Info.Priority = 0 Then
'            ' /* determine what to do with normal priority notifications */
'            ApplyPresence = Val(mInfo.ClassObj.GetValue("override-busy")) - 1
'            If ApplyPresence = SN_PA_DO_DEFAULT Then _
'                ApplyPresence = Val(g_ConfigGet("busy_mode"))
'
'        End If
'
'    ElseIf g_IsAway() Then
'        ' /* Away mode */
'        g_Debug "TNotification.ApplyPresence(): Away mode enabled: priority=" & CStr(Info.Priority)
'        If mInfo.Priority < 0 Then
'            ' /* low priority = ignored */
'            ApplyPresence = SN_PA_DO_NOTHING
'
'        ElseIf Info.Priority = 0 Then
'            ' /* determine what to do with normal priority notifications */
'            ApplyPresence = Val(mInfo.ClassObj.GetValue("override-away")) - 1
'            Debug.Print "### away class setting is: " & ApplyPresence
'            If ApplyPresence = SN_PA_DO_DEFAULT Then _
'                ApplyPresence = Val(g_ConfigGet("away_mode"))
'
'            Debug.Print "### away setting is: " & ApplyPresence
'
'        Else
'            ' /* high priority = logged as missed irrespective of class setting */
'            ApplyPresence = SN_PA_LOG_AS_MISSED
'
'        End If
'
'    Else
'        ' /* Normal/Active */
'        ApplyPresence = Val(Info.ClassObj.GetValue("override-normal")) - 1
'
'    End If
'
'    ' /* post-processing */
'
'    Select Case ApplyPresence
'    Case SN_PA_MAKE_STICKY
'        mInfo.Timeout = 0
'
'    Case SN_PA_DISPLAY_URGENT
'        mInfo.Priority = 1
'
'    End Select
'
'End Function
'
'Public Sub ValidateStyle()
'
'    ' /* no style name?  that's okay, use the default */
'
'    With mInfo
'        If .StyleName = "" Then
'            .StyleName = style_GetStyleName(g_ConfigGet("default_style"))
'            .SchemeName = style_GetSchemeName(g_ConfigGet("default_style"))
'
'        End If
'
'    End With
'
'End Sub

Public Function AppNameAndTitle() As String

    AppNameAndTitle = mInfo.ClassObj.App.Name
    If mInfo.Title <> "" Then _
        AppNameAndTitle = AppNameAndTitle & ": " & mInfo.Title

End Function

Public Function PackedArgs() As BPackedData

    Set PackedArgs = mPacked

End Function

Public Function AppName() As String

    AppName = mInfo.ClassObj.App.Name

End Function

Public Sub Replay()

    g_NotificationRoster.Add mInfo, Nothing, False, True
    mReplayed = True

End Sub

Public Function WasReplayed() As Boolean

    WasReplayed = mReplayed

End Function

Public Function Icon() As MImage

    Set Icon = mIcon

End Function
