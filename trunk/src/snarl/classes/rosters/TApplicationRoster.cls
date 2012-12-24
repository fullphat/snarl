VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TApplicationRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TApplicationRoster.cls
    '/
    '/  Description:    Manages registered applications
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mApp() As TApp
Dim mCount As Long
Dim mIndex As Long

Dim mRefs As Long
Dim mAppToken As Long       ' // V41
Dim mSnarlAppMsg As Long
Dim mRegLock As Long        ' // R2.4.2: was global

Dim mRemoteAppList As BTagList

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT

    mRefs = mRefs - 1
    If mRefs = 0 Then
        ' /* run tidyup code */

        Me.Unregister gSnarlToken, gSnarlPassword, False
        SendToAll SNARL_BROADCAST_QUIT

        ReDim mApp(0)
        mCount = 0

    End If

End Function

Private Sub MLibrary_Initialize()

    ' /* not currently used */

End Sub

Private Function MLibrary_Magic() As Long

    MLibrary_Magic = &H23232323

End Function

Private Function MLibrary_Open() As melon.M_RESULT
Dim pAppInfo As T_SNARL_APP
Dim pSnarlApp As TApp

    g_Debug "TApplicationRoster.Open()", LEMON_LEVEL_PROC

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */
        mSnarlAppMsg = RegisterWindowMessage("SnarlAppMessage")     ' // V40.25
        mAppToken = (&H1E0 + (Rnd * 80)) And &HFFFFFFFC
        Set mRemoteAppList = new_BTagList()

        ' /* register ourselves with, er, ourself! */

        g_IsRunning = True
        g_Debug "TApplicationRoster.Open(): registering ourself..."

        ' /* V42: password protect Snarl */
'        gSnarlPassword = uGenerateSnarlPassword()

        With pAppInfo
            .hWnd = frmAbout.hWnd
            .Icon = g_MakePath(App.Path) & "etc\icons\snarl.png"
'            .LargeIcon = .Icon
            .Name = App.Title
            .Password = gSnarlPassword
            .Pid = GetCurrentProcessId()
            .Signature = App.ProductName
            .uMsg = &H555

        End With

        gSnarlToken = uAddByStruct(pAppInfo, pSnarlApp)

        If gSnarlToken = 0 Then
            ' /* panic? */

        Else
            g_Debug "TApplicationRoster.Open(): registered as " & g_HexStr(gSnarlToken) & ", adding classes..."

            With pSnarlApp
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_GENERAL & "#?name::Welcome Message and other notifications#?password::" & gSnarlPassword)
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_APP_REG & "#?name::Application registered#?password::" & gSnarlPassword)
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_APP_UNREG & "#?name::Application unregistered#?password::" & gSnarlPassword)

                ' /* V38.97 */
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_JSON & "#?name::Incoming JSON notifications#?password::" & gSnarlPassword)

                ' /* V38.142 */
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_ANON_NET & "#?name::Anonymous remote notifications#?password::" & gSnarlPassword)

                ' /* V39.21 */
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_ANON & "#?name::Anonymous local notifications#?password::" & gSnarlPassword)

                ' /* R2.4 */
'                .AddClass g_newBPackedData("id::" & SNARL_CLASS_LOW_PRIORITY & "#?name::Low priority notifications#?password::" & gSnarlPassword)
'                .AddClass g_newBPackedData("id::" & SNARL_CLASS_SYSTEM & "#?name::Snarl Internal Notifications#?password::" & gSnarlPassword)

            End With
        End If

        g_IsRunning = False

    End If

    g_Debug "TApplicationRoster.Open() done", LEMON_LEVEL_PROC

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()

    ' /* not currently used */

End Sub

Private Function MRoster_CountItems() As Long

    MRoster_CountItems = mCount

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long
Dim i As Long

    If mCount Then
        Name = LCase$(Name)
        For i = 1 To mCount
            If LCase$(mApp(i).Name) = Name Then
                MRoster_FindItem = i
                Exit Function

            End If
        Next i

        g_Debug "TApplicationRoster.FindItem(): app '" & Name & "' not found", LEMON_LEVEL_WARNING

    Else
        g_Debug "TApplicationRoster.FindItem(): no apps registered!", LEMON_LEVEL_CRITICAL

    End If

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.mObject

    If (Index > 0) And (Index <= mCount) Then _
        Set MRoster_ItemAt = mApp(Index)

End Function

Private Function MRoster_NextItem(Item As melon.mObject) As Boolean

    If mIndex <= mCount Then
        Set Item = mApp(mIndex)
        mIndex = mIndex + 1
        MRoster_NextItem = True

    End If

End Function

Private Sub MRoster_Rewind()

    mIndex = 1

End Sub

Private Property Get MVersionInfo_Date() As String

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "application.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

Public Function Add41(ByRef Data As BPackedData, ByRef SenderSocket As CSocket, ByVal SenderPID As Long, ByVal RemoteHostName As String) As Long '// ByVal IsRemote As Boolean,
Dim szTitle As String
Dim sz As String

    g_Debug "TApplicationRoster.Add41()", LEMON_LEVEL_PROC_ENTER

    ' /* validate args */

    szTitle = Data.ValueOf("app-title")         ' // R2.4.2 DR3
    If szTitle = "" Then _
        szTitle = Data.ValueOf("title")

    sz = Data.ValueOf("app-sig")
    If sz = "" Then _
        sz = Data.ValueOf("id")

    ' /* R2.4.2 DR3: if no signature, craft it from the title */

    If (sz = "") And (szTitle <> "") Then
        sz = "application/x-" & LCase$(Replace$(szTitle, " ", "_"))
        g_Debug "created app-sig '" & sz & "'", LEMON_LEVEL_INFO

    End If

    ' /* error checking */

    If (szTitle = "") Or (sz = "") Then
        g_Debug "required arg missing", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    If InStr(sz, " ") Then
        g_Debug "application signature cannot contain spaces", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    ' /* 2.4.2 "only_allow_secure_apps" setting */

    If (g_ConfigGet("only_allow_secure_apps") = "1") And (Data.ValueOf("password") = "") Then
        g_Debug "user settings require app must provide a password", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_AUTH_FAILURE
        Exit Function

    End If

Dim i As Long

    ' /* check for duplicate and return existing token */
'And (mApp(i).RemoteHostName = RemoteHostName)

    If mCount Then
        For i = 1 To mCount
            If mApp(i).Signature = sz Then
                ' /* R2.4 DR7: check same password provided */
                If (mApp(i).Password <> "") And (mApp(i).Password <> Data.ValueOf("password")) Then
                    g_Debug "'" & sz & "' already registered: password mismatch", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
                    g_SetLastError SNARL_ERROR_AUTH_FAILURE
                    Add41 = 0
                    Exit Function

                Else
                    g_Debug "'" & sz & "' already registered (token=" & mApp(i).Token & ")", LEMON_LEVEL_PROC_EXIT

                    uUpdate i, Data

                    Add41 = mApp(i).Token
                    Exit Function

                End If
            End If
        Next i
    End If

Dim pInfo As T_SNARL_APP
Dim szIcon As String

    ' /* configure the item */

    With pInfo
        .Name = szTitle
        .Signature = sz

        If Data.Exists("app-icon") Then
            ' /* R2.5 Beta 2: so we can separate the two when using NOTIFY to register */
            .Icon = Data.ValueOf("app-icon")

        Else
            .Icon = Data.ValueOf("icon")

        End If

        .Pid = SenderPID                        ' // R2.4.2: added
        .RemoteHostName = RemoteHostName        ' // R2.4.2 DR3

        If (.Icon = "") And (Data.Exists("icon-base64")) Then
            ' /* R2.4 Beta 4 - support for Base64-encoded icons */
            .Icon = g_GetBase64Icon(Data.ValueOf("icon-base64"))

        ElseIf (.Icon = "") And (Data.Exists("icon-phat64")) Then
            ' /* R2.4.2 DR3: phat64 encoding */
            .Icon = g_GetPhat64Icon(Data.ValueOf("icon-phat64"))

        ElseIf g_IsURL(.Icon) Then
            ' /* R2.4.2: initiate an async download (user will get either previously cached, or "no icon" icon in the meanwhile) */
            sz = g_GetUserFolderPathStr(False) & "etc\app-cache\" & g_MakeFilename(.Signature) '& g_GetExtension(.Icon)
            uStartRemoteIconDownload .Icon, sz
            .Icon = sz

        ElseIf g_IsFileURI(.Icon) Then
            ' /* R2.4.2 DR3: support for "file://" protocol */
            .Icon = g_PathFromURI(.Icon)


        ElseIf (.Icon <> "") And (g_GetExtension(.Icon) = "") Then
            ' /* filename with no extension - could be any format... */
'            MsgBox .Icon
            If uGetBestIcon(.Icon, szIcon) Then _
                .Icon = szIcon

        End If

'        .LargeIcon = .Icon                       ' // no need for two icons: V41 spec recommends 128x128px icon

        ' / R2.4.1 Beta 3 - if (still) no icon, try to use the sender exe's one */



        If (.Icon = "") And (.RemoteHostName <> "") Then
            ' /* truly remote apps get this icon, all others get the normal one */
            .Icon = g_MakePath(App.Path) & "etc\icons\no_icon-remote.png"

        ElseIf (.Icon = "") And (SenderPID <> 0) Then
            sz = g_GetExeFullPath(SenderPID)
            If sz <> "" Then
                If uGetBestIcon(sz, szIcon) Then _
                    pInfo.Icon = szIcon

            End If
        End If

        Debug.Print "--> icon = " & .Icon

        If Data.Exists("reply-to") Then
            ' /* V42.22: new name */
            .hWnd = g_SafeLong(Data.ValueOf("reply-to"))

        Else
            .hWnd = g_SafeLong(Data.ValueOf("hwnd"))

        End If

        If Data.Exists("reply-with") Then
            ' /* R2.4.2 DR3: new new name :( */
            .uMsg = g_SafeLong(Data.ValueOf("reply-with"))

        ElseIf Data.Exists("reply") Then
            ' /* V42.22: new name */
            .uMsg = g_SafeLong(Data.ValueOf("reply"))

        Else
            .uMsg = g_SafeLong(Data.ValueOf("umsg"))

        End If

        .Flags = Val(Data.ValueOf("flags"))
        .Password = Data.ValueOf("password")

        ' /* R2.4.2 DR3: if "hint" is specified this will be displayed when the
        '    "Details" button in [Apps] is clicked */

        If Data.Exists("hint") Then _
            .Hint = Data.ValueOf("hint")

'        ' /* R2.4.2 DR3: if "app-flags" is specified and we have a valid reply-to
'        '    window, this can be used in place of the unituitive "flags" parameter. */
'
'        If (Data.Exists("app-flags")) And ((IsWindow(.hWnd) <> 0) Or (Not (SenderSocket Is Nothing))) Then
'            With New BPackedData
'                If .SetTo(Data.ValueOf("app-flags"), "|", "") Then
'                    If .Exists("prefs") Then _
'                        pInfo.Flags = pInfo.Flags Or (SNARLAPP_HAS_PREFS Or SNARLAPP_IS_WINDOWLESS)
'
'                    If .Exists("about") Then _
'                        pInfo.Flags = pInfo.Flags Or (SNARLAPP_HAS_ABOUT Or SNARLAPP_IS_WINDOWLESS)
'
'                End If
'            End With
'        End If

        ' /* R2.4.2 DR3: if "config-tool" is specified then it gets included
        '    in the tray sub-menu */

        If Data.Exists("config-tool") Then
            .Tool = Data.ValueOf("config-tool")
            .IncludeInMenu = True

        End If

        ' /* R2.4.2 DR3: if "app-daemon" is "1" then it gets included in the
        '    tray sub-menu and is marked as V43 */

        If Data.ValueOf("app-daemon") = "1" Then
            ' /* V43 windowless apps are assumed to support all features */
            If IsWindow(.hWnd) <> 0 Then
                .Flags = SNARLAPP_HAS_ABOUT Or SNARLAPP_HAS_PREFS Or SNARLAPP_IS_DAEMON
                .IncludeInMenu = True

            ElseIf Not (SenderSocket Is Nothing) Then
                .IncludeInMenu = True
            
            End If

        ElseIf IsWindow(.hWnd) <> 0 Then
            ' /* pre-V43 support: Win32 API only */
            .IncludeInMenu = (GetProp(.hWnd, "snarl_app") = 1) Or ((.Flags And SNARLAPP_IS_WINDOWLESS) <> 0)

        End If

        ' /* R2.4.2 DR3 - tweak remote settings for TCP protocols: if the app is being registered
        '    by a remote source, it cannot have dynamic About/Prefs callbacks */

'        .IsRemote = IsRemote
        Set .Socket = SenderSocket
'        If Not (.Socket Is Nothing) Then _
            .IsRemote = (InStr(get_ip_address_table(), .Socket.RemoteHostIP) = 0)

        ' /* R2.4.2 DR3 - remotely registered apps cannot have these */

        If .RemoteHostName <> "" Then _
            .Flags = .Flags And (Not (SNARLAPP_HAS_ABOUT Or SNARLAPP_HAS_PREFS Or SNARLAPP_IS_WINDOWLESS))

        g_Debug "name='" & .Name & "' signature='" & .Signature & "' flags=0x" & g_HexStr(.Flags)


        ' /* R2.5.1 */

        If Data.Exists("keep-alive") Then _
            .KeepAlive = (Data.ValueOf("keep-alive") = "1")

    End With

    Add41 = uAddByStruct(pInfo)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function uAddByStruct(ByRef Info As T_SNARL_APP, Optional ByRef NewAppObject As TApp) As Long
Dim pa As TApp

    ' /* assign token and bump counter on */

    Info.Token = mAppToken
    mAppToken = mAppToken + 4

    Set pa = New TApp

    ' /* unlikely, but bInit() can fail... */

    If Not pa.bInit(Info) Then _
        Exit Function

Dim bShow As Boolean

    ' /* add it to our roster */

    uAlphaAdd pa
    bShow = True

    ' /* don't show if it's us or we're locked */

    If (pa.Name = App.Title) Or (mRegLock) > 0 Then
        bShow = False

    ElseIf g_ConfigGet("notify_on_first_register") = "1" Then
        ' /* if notify_on_first_register is set, we need to take account of the IsNewApp setting */
        bShow = pa.IsNew()

    End If

Dim lFlags As SN_NOTIFICATION_FLAGS
Dim i As Long

    If bShow Then

        ' /* if app has a password or was registered via some IP protocol
        '    add the relevant system flags here */

'        If pa.Password <> "" Then _
            lFlags = lFlags Or SN_NF_SECURE

'        If pa.IsRemote Then _
            lFlags = lFlags Or SN_NF_REMOTE

        ' /* display the 'app registered' notification */

        i = g_PrivateNotify(SNARL_CLASS_APP_REG, _
                            IIf(pa.IsRemote, "Remote a", "A") & "pplication registered", _
                            pa.NameEx & " registered successfully with Snarl" & IIf(pa.IsBanned, vbCrLf & vbCrLf & "NOTE: Application has been blocked by Snarl", ""), _
                            -1, _
                            IIf(pa.Icon = "", g_MakePath(App.Path) & "etc\icons\no_icon.png", pa.Icon), _
                            , , , _
                            lFlags)

        If i Then
            ' /* add the action(s) */
            g_QuickAddAction i, "Manage Notifications...", "!notifications " & pa.Signature
            If pa.HasConfig Then _
                g_QuickAddAction i, pa.Name & " Preferences...", "!app_settings " & pa.Signature

        End If
    End If

    frmAbout.bUpdateAppList
    Set NewAppObject = pa
    uAddByStruct = pa.Token

End Function

Public Function OldAdd(ByVal Name As String, ByVal hWnd As Long, ByVal uMsg As Long, ByVal Icon As String, ByVal LargeIcon As String, ByVal Pid As Long) As M_RESULT

    g_Debug "TApplicationRoster.OldAdd()", LEMON_LEVEL_PROC_ENTER

Dim pInfo As T_SNARL_APP
Dim pa As TApp

    If Me.FindByWindow(hWnd, pa) <> 0 Then
        g_Debug "config window is already registered", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        OldAdd = M_ALREADY_EXISTS
        Exit Function

    End If

    ' /* configure the item */

    With pInfo
        .Name = Name
        .hWnd = hWnd
        .uMsg = uMsg
        .Icon = IIf(LargeIcon <> "", LargeIcon, Icon)
'        .LargeIcon = LargeIcon
        .Signature = Name
        ' /* V38.119 - for V39: can unregister an app using it's name and pid */
        .Pid = Pid

    End With

    If uAddByStruct(pInfo) <> 0 Then
        ' /* success (token was returned) */
        OldAdd = M_OK

    Else
        OldAdd = M_FAILED

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

'Private Sub uAdd(ByRef theApp As TApp)          ', ByVal IsNewApp As Boolean
'
'    uAlphaAdd theApp
'
'    ' /* check version - deprecated as of V41 (wasn't ever used?) */
'
''    If (theApp.hWnd <> 0) And (theApp.uMsg <> 0) Then
''        i = SendMessage(theApp.hWnd, theApp.uMsg, SNARL_ASK_APPLET_VER, ByVal 0&)
''        If i <> 0 Then
''            g_Debug "TApplicationRoster.uAdd(): app responded to SNARL_ASK_APPLET_VER with " & CStr(i), LEMON_LEVEL_INFO
''
''        Else
''            g_Debug "TApplicationRoster.uAdd(): app ignored SNARL_ASK_APPLET_VER"
''
''        End If
''
''    End If
'
'Dim bShow As Boolean
'
'    bShow = True
'
'    ' /* don't show if it's us or we're locked */
'
''Debug.Print "(): " & gPrefs.notify_on_first_register & " >> " & mRegLock
'
'    If (theApp.Name = App.Title) Or (mRegLock) > 0 Then
'        bShow = False
'
'    ElseIf g_ConfigGet("notify_on_first_register") = "1" Then
'        ' /* if notify_on_first_register is set, we need to take account of the IsNewApp setting */
'        bShow = theApp.IsNew()
'
'    End If
'
'Dim lFlags As SNRL_NOTIFICATION_FLAGS
'Dim i As Long
'
'    If bShow Then
'
'        ' /* if app has a password or was registered via some IP protocol
'        '    add the relevant system flags here */
'
'        If theApp.Password <> "" Then _
'            lFlags = lFlags Or SNRL_NOTIFICATION_SECURE
'
'        If theApp.IsRemote Then _
'            lFlags = lFlags Or SNRL_NOTIFICATION_REMOTE
'
'        ' /* create the notification */
'
'        i = g_PrivateNotify(SNARL_CLASS_APP_REG, _
'                            "Application registered", _
'                            theApp.Name & " registered successfully with Snarl", _
'                            -1, _
'                            IIf(theApp.LargeIcon = "", g_MakePath(App.Path) & "etc\icons\snarl.png", theApp.LargeIcon), , , _
'                            SNARL41_NOTIFICATION_AUTO_DISMISS, lFlags)
'
'        If i Then
'            ' /* add the action(s) */
'            sn42AddAction i, theApp.Name & " Settings...", "!snarl cfg " & theApp.Signature
'            If theApp.HasConfig Then _
'                sn42AddAction i, "Configure " & theApp.Name & "...", "!snarl manage " & theApp.Signature
'
''            szMenu = "Notification Settings...#?!snarl cfg " & theApp.Signature
''            If theApp.HasConfig Then _
'                szMenu = szMenu & "||Configure " & theApp.Name & "...#?!snarl manage " & theApp.Signature
'
''            snChangeAttribute i, SNARL_ATTRIBUTE_MENU, szMenu
'
'        End If
'
'    End If
'
'    frmAbout.bUpdateAppList
'
'End Sub

Private Sub uAlphaAdd(ByRef theApp As TApp)
Static i As Long
Static j As Long

    ' /* add it alpha-sorted */

    If mCount Then
        For i = 1 To mCount
            If LCase$(theApp.Name) < LCase$(mApp(i).Name) Then
                mCount = mCount + 1
                ReDim Preserve mApp(mCount)
                For j = mCount To (i + 1) Step -1
                    Set mApp(j) = mApp(j - 1)

                Next j

                Set mApp(i) = theApp
                Exit Sub

            End If
        Next i
    End If
    
    ' /* drop through here if no other apps or can be added to end of list */

    mCount = mCount + 1
    ReDim Preserve mApp(mCount)
    Set mApp(mCount) = theApp

End Sub

Public Function Register(ByVal Name As String, ByVal hWnd As Long, ByVal uMsg As Long, ByVal Icon As String, ByVal LargeIcon As String, ByVal Pid As Long) As M_RESULT

    ' /* for snRegisterApp() - to be introduced in V39 */
    g_Debug "TApplicationRoster.Register()", LEMON_LEVEL_PROC_ENTER
    g_Debug "Name='" & Name & "' pid=" & CStr(Pid) & " hWnd=0x" & g_HexStr(hWnd) & " ReplyMsg=0x" & g_HexStr(uMsg)

    If Pid = 0 Then
        Register = M_INVALID_ARGS
        g_Debug "pid can't be null", LEMON_LEVEL_CRITICAL

    ElseIf uFindByPid(Pid) <> 0 Then
        Register = M_ALREADY_EXISTS
        g_Debug "app with pid " & CStr(Pid) & " already registered", LEMON_LEVEL_CRITICAL

    Else
        ' /* carry on from here */
        Register = OldAdd(Name, hWnd, uMsg, Icon, LargeIcon, Pid)

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function OldRemove(ByVal hWnd As Long) As Boolean

    If mCount = 0 Then _
        Exit Function

Dim i As Long

    For i = 1 To mCount
        If mApp(i).hWnd = hWnd Then
            OldRemove = (Remove(i) = M_OK)
            Exit Function

        End If

    Next i

    g_Debug "TApplicationRoster.OldRemove(): app (window) " & g_HexStr(hWnd) & " not found", LEMON_LEVEL_WARNING

End Function

Public Function FindByWindow(ByVal hWnd As Long, ByRef App As TApp) As Boolean

    If (hWnd = 0) Or (mCount = 0) Then _
        Exit Function

Dim i As Long

    For i = 1 To mCount
        If mApp(i).hWnd = hWnd Then
            Set App = mApp(i)
            FindByWindow = True
            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.FindByWindow(): window " & g_HexStr(hWnd) & " not found", LEMON_LEVEL_CRITICAL

End Function

Public Function FindByToken(ByVal Token As Long, ByRef App As TApp, ByVal Password As String) As Boolean

    If Token = 0 Then
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    If mCount = 0 Then
        g_SetLastError SNARL_ERROR_NOT_REGISTERED
        Exit Function

    End If

Static i As Long

    For i = 1 To mCount
        If mApp(i).Token = Token Then
            ' /* V42: validate password, if the app has one */
            If (mApp(i).Password <> "") And (mApp(i).Password <> Password) Then
                g_Debug "TApplicationRoster.FindByToken(): incorrect password", LEMON_LEVEL_CRITICAL
                g_SetLastError SNARL_ERROR_AUTH_FAILURE

            Else
                g_SetLastError SNARL_SUCCESS
                Set App = mApp(i)
                FindByToken = True

            End If

            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.FindByToken(): " & g_HexStr(Token) & " not found", LEMON_LEVEL_CRITICAL
    g_SetLastError SNARL_ERROR_NOT_REGISTERED

End Function

Public Function FindBySignature(ByVal Signature As String, ByRef App As TApp, ByVal Password As String) As Boolean

    If Signature = "" Then
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    If mCount = 0 Then
        g_SetLastError SNARL_ERROR_NOT_REGISTERED
        Exit Function

    End If
    
Static i As Long

    For i = 1 To mCount
        If (mApp(i).Signature = Signature) Then                         ' // And (mApp(i).RemoteHostName = RemoteIP)
            ' /* V42: validate password, if the app has one */
            If (mApp(i).Password <> "") And (mApp(i).Password <> Password) Then
                g_Debug "TApplicationRoster.FindBySignature(): incorrect password", LEMON_LEVEL_CRITICAL
                g_SetLastError SNARL_ERROR_AUTH_FAILURE

            Else
                g_SetLastError SNARL_SUCCESS
                Set App = mApp(i)
                FindBySignature = True

            End If

            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.FindBySignature(): " & Signature & " not found", LEMON_LEVEL_CRITICAL
    g_SetLastError SNARL_ERROR_NOT_REGISTERED

End Function

Public Function Find(ByVal Name As String, ByRef App As TApp) As Boolean
Dim i As Long

    i = MRoster_FindItem(Name)
    If i > 0 Then
        Set App = mApp(i)
        Find = True

    End If

End Function

Public Function AppAt(ByVal Index As Long) As TApp

    Set AppAt = MRoster_ItemAt(Index)

End Function

Public Sub SendToAll(ByVal EventId As SNARL_GLOBAL_EVENTS)
Static i As Long

    If mCount = 0 Then _
        Exit Sub

    g_Debug "TApplicationRoster.SendToAll(): sending event ID '" & CStr(EventId) & "' to all registered apps..."

    For i = mCount To 1 Step -1
        If (IsWindow(mApp(i).hWnd) <> 0) And (mApp(i).uMsg <> 0) Then _
            PostMessage mApp(i).hWnd, mApp(i).uMsg, EventId, ByVal 0&

    Next i

End Sub

Public Function CountApps() As Long

    CountApps = mCount

End Function

'Public Sub WriteConfig()
'Dim i As Long
'
'    If mCount Then
'        For i = 1 To mCount
'            mApp(i).UpdateConfig
'
'        Next i
'    End If
'
'End Sub

Public Sub ResetAlerts()
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            mApp(i).ResetAlerts

        Next i
    End If

End Sub

Public Function OldUnregister(ByVal Pid As Long) As M_RESULT

    If Pid = 0 Then
        g_Debug "TApplicationRoster.OldUnregister(): invalid arg", LEMON_LEVEL_CRITICAL
        OldUnregister = M_INVALID_ARGS
        Exit Function

    End If

    If mCount = 0 Then
        g_Debug "TApplicationRoster.OldUnregister(): no applications registered", LEMON_LEVEL_CRITICAL
        OldUnregister = M_FAILED
        Exit Function

    End If

Dim sz As String
Dim i As Long

    i = uFindByPid(Pid)
    If i Then
        sz = mApp(i).Name
        OldUnregister = Remove(i)
        g_Debug "TApplicationRoster.OldUnregister(): app " & CStr(Pid) & " (" & sz & ") unregistered okay"

    Else

        g_Debug "TApplicationRoster.OldUnregister(): app #" & CStr(Pid) & " not found", LEMON_LEVEL_CRITICAL
        OldUnregister = M_NOT_FOUND

    End If

End Function

Public Function Remove(ByVal Index As Long) As M_RESULT

    If (Index < 1) Or (Index > mCount) Then
        g_Debug "TApplicationRoster.Remove(): bad index #" & CStr(Index), LEMON_LEVEL_CRITICAL
        Remove = M_INVALID_ARGS
        Exit Function

    End If

Dim szIcon As String
Dim sz As String
Dim i As Long

    sz = mApp(Index).Name
    szIcon = IIf(mApp(Index).Icon = "", g_MakePath(App.Path) & "etc\icons\no_icon.png", mApp(Index).Icon)

    If Index < mCount Then
        For i = Index To (mCount - 1)
            Set mApp(i) = mApp(i + 1)

        Next i
    End If

    mCount = mCount - 1
    ReDim Preserve mApp(mCount)

    ' /* don't show if it's us unregistering ;-) */

    If (sz <> App.Title) And (mRegLock < 1) And (g_ConfigGet("notify_on_first_register") = "0") Then _
        g_PrivateNotify SNARL_CLASS_APP_UNREG, "Application unregistered", _
                     sz & " unregistered with Snarl", , _
                     szIcon

    frmAbout.bUpdateAppList

    Remove = M_OK

End Function

Private Function uFindByPid(ByVal Pid As Long) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If (mApp(i).Pid = Pid) Then
                uFindByPid = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function FindByPid(ByVal Pid As Long, ByRef App As TApp) As Boolean
Dim i As Long

    i = uFindByPid(Pid)
    If i Then
        Set App = mApp(i)
        FindByPid = True

    End If

End Function

Public Function IndexOfPidAndName(ByVal Pid As Long, ByVal Name As String) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If (mApp(i).Pid = Pid) And (mApp(i).Name = Name) Then
                IndexOfPidAndName = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function IndexOf(ByVal Title As String) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If mApp(i).Name = Title Then
                IndexOf = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function IndexOfSig(ByVal Signature As String) As Long
Static i As Long

    If mCount Then
        For i = 1 To mCount
            If mApp(i).Signature = Signature Then
                IndexOfSig = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function UnregisterBySig(ByVal Signature As String, ByVal Password As String) As Long
Static i As Long

    i = IndexOfSig(Signature)
    If i = 0 Then
        g_Debug "TApplicationRoster.UnregisterBySig(): '" & Signature & "' not found", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_NOT_REGISTERED

    Else
        UnregisterBySig = uUnregister(i, Password, False)

    End If

End Function

Public Function Unregister(ByVal Token As Long, ByVal Password As String, ByVal RemoveSettings As Boolean) As Long
Static i As Long

    i = uIndexOfToken(Token)
    If i = 0 Then
        g_Debug "TApplicationRoster.Unregister(): '" & CStr(Token) & "' not found", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_NOT_REGISTERED

    Else
        Unregister = uUnregister(i, Password, RemoveSettings)

    End If

End Function

Private Function uUnregister(ByVal Index As Long, ByVal Password As String, ByVal RemoveSettings As Boolean) As Long

    ' /* WARNING: index is _NOT_ bounds-checked */

    ' /* V42: password must match */
    If mApp(Index).Password <> "" Then
        If mApp(Index).Password <> Password Then
            g_SetLastError SNARL_ERROR_AUTH_FAILURE
            Exit Function

        End If
    End If

    If RemoveSettings Then _
        mApp(Index).RemoveSettings

    Remove Index
    uUnregister = -1

End Function

Private Function uIndexOfToken(ByVal Token As Long) As Long

    If (mCount < 1) Or (Token = 0) Then _
        Exit Function

Static i As Long

    For i = 1 To mCount
        If mApp(i).Token = Token Then
            uIndexOfToken = i
            Exit Function

        End If
    Next i

End Function

Public Function Update(ByVal Token As Long, ByRef Data As BPackedData) As Long

    If (Data Is Nothing) Then
        g_Debug "TApplicationRoster.Update(): arg missing", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

Static i As Long

    i = uIndexOfToken(Token)
    If i = 0 Then
        g_Debug "TApplicationRoster.Update(): '" & CStr(Token) & "' not found", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_NOT_REGISTERED
        Exit Function

    End If

    ' /* V42: password must match */

    If mApp(i).Password <> "" Then
        If mApp(i).Password <> Data.ValueOf("password") Then
            g_Debug "TApplicationRoster.Update(): '" & CStr(Token) & "' password mismatch", LEMON_LEVEL_CRITICAL
            g_SetLastError SNARL_ERROR_AUTH_FAILURE
            Exit Function

        End If
    End If

'        If Data.Exists("title") Then _
        mApp(i)Name = Data.ValueOf("title")

    If Data.Exists("icon") Then _
        mApp(i).SetIcon Data.ValueOf("icon")

    frmAbout.bUpdateAppList
    Update = -1

End Function

'Friend Sub bAddClass(ByRef Class As TAlert)
'
'    mClasses = mClasses + 1
'    ReDim Preserve mClass(mClasses)
'    Set mClass(mClasses) = Class
'
'End Sub
'
'Public Function FindClass(ByVal Token As Long, ByRef Class As TAlert) As Boolean
'
'    If mClasses = 0 Then _
'        Exit Function
'
'Dim i As Long
'
'    For i = 1 To mClasses
'        If mClass(i).Token = Token Then
'            Set Class = mClass(i)
'            FindClass = True
'            Exit Function
'
'        End If
'    Next i
'
'End Function

'Public Function SnarlAppsMenu() As OMMenu
'Dim pm As OMMenu
'Dim i As Long
'
'    Set SnarlAppsMenu = New OMMenu
'
'    With SnarlAppsMenu
'        If mCount Then
'            For i = 1 To mCount
'                If mApp(i).IncludeInMenu Then _
'                    .AddItem .CreateItem("!" & CStr(i), mApp(i).Name)
'
'            Next i
'        End If
'
'        If .CountItems = 0 Then _
'            .AddItem .CreateItem("", "None", , False)
'
'    End With
'
'End Function

Private Function uGenerateSnarlPassword() As String
Dim i As Integer
Dim c As Integer

    Randomize Timer
    c = (Rnd * 24) + 8      ' // variable length between 8 and 32 chars

    For i = 1 To c
        Randomize Timer
        uGenerateSnarlPassword = uGenerateSnarlPassword & g_HexStr(Rnd * 255, 2)

    Next i

End Function

Private Sub uUpdate(ByVal Index As Long, ByRef Data As BPackedData)

    If (Index < 1) Or (Index > mCount) Or (Data Is Nothing) Then _
        Exit Sub

    ' /* can only change certain things */

Dim szPath As String
Dim pi As BIcon
Dim sz As String
Dim rm As Long
Dim rt As Long

    With mApp(Index)

        If Data.Exists("icon") Then
            sz = Data.ValueOf("icon")

            If (sz = "") And (Data.Exists("icon-base64")) Then
                ' /* R2.4 Beta 4 - support for Base64-encoded icons */
                sz = g_GetBase64Icon(Data.ValueOf("icon-base64"))
    
            ElseIf (sz = "") And (Data.Exists("icon-phat64")) Then
                ' /* R2.4.2 DR3: phat64 encoding */
                sz = g_GetPhat64Icon(Data.ValueOf("icon-phat64"))
    
            ElseIf g_IsURL(sz) Then
                ' /* R2.4.2: initiate an async download (user will get either previously cached, or "no icon" icon in the meanwhile) */
                szPath = g_GetUserFolderPathStr(False) & "etc\icons\app-cache\" & g_MakeFilename(.Signature) '& g_GetExtension(sz)
                uStartRemoteIconDownload sz, szPath
                sz = szPath
    
            ElseIf g_IsFileURI(sz) Then
                ' /* R2.4.2 DR3: support for "file://" protocol */
                sz = g_PathFromURI(sz)
    
            End If

            ' / R2.4.1 Beta 3 - if (still) no icon, try to use the sender exe's one */

            If (sz = "") And (.IsRemote) Then
                sz = g_MakePath(App.Path) & "etc\icons\no_icon-remote.png"

            ElseIf (sz = "") And (.Pid <> 0) Then
                szPath = g_GetExeFullPath(.Pid)
                If szPath <> "" Then
                    With new_BIconContent(szPath)
                        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pi) Then
                            With New mfxView
                                .SizeTo pi.Width, pi.Height
                                .DrawScaledImage pi.Render, , new_BPoint(pi.Width, pi.Height)
                                sz = g_GetSafeTempIconPath()
                                .WriteToFile sz, "image/png"
    
                            End With
                        End If
                    End With
                End If
            End If

            .SetIcon sz

        End If


        If Data.Exists("reply-to") Then
            rt = g_SafeLong(Data.ValueOf("reply-to"))

        Else
            rt = g_SafeLong(Data.ValueOf("hwnd"))

        End If

        If Data.Exists("reply-with") Then
            ' /* R2.4.2 DR3 */
            rm = g_SafeLong(Data.ValueOf("reply-with"))

        ElseIf Data.Exists("reply") Then
            ' /* V42.22: new name */
            rm = g_SafeLong(Data.ValueOf("reply"))

        Else
            rm = g_SafeLong(Data.ValueOf("umsg"))

        End If

        .SetReply rt, rm

    End With

End Sub

Public Sub LockRegister()

    mRegLock = mRegLock + 1

End Sub

Public Sub UnlockRegister()

    mRegLock = mRegLock - 1

End Sub

Private Sub uStartRemoteIconDownload(ByVal URL As String, ByVal Path As String)
Dim Pid As TIconDownloader

    Set Pid = New TIconDownloader
    Pid.Go URL, Path
    mRemoteAppList.Add Pid

End Sub

Public Sub RemoveIconDownloader(ByVal Guid As String)

    g_Debug "TApplicationRoster.RemoveIconDownloader(): guid=" & Guid & " index=" & mRemoteAppList.IndexOf(Guid)
    mRemoteAppList.Remove mRemoteAppList.IndexOf(Guid)

End Sub

Public Function PrivateFindBySignature(ByVal Signature As String, ByRef App As TApp) As Boolean

    If (Signature = "") Or (mCount = 0) Then _
        Exit Function
    
Static i As Long

    For i = 1 To mCount
        If mApp(i).Signature = Signature Then
            Set App = mApp(i)
            PrivateFindBySignature = True
            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.PrivateFindBySignature(): " & Signature & " not found", LEMON_LEVEL_CRITICAL

End Function

Public Function PrivateFindByToken(ByVal Token As Long, ByRef App As TApp) As Boolean

    If (Token = 0) Or (mCount = 0) Then _
        Exit Function
    
Static i As Long

    For i = 1 To mCount
        If mApp(i).Token = Token Then
            Set App = mApp(i)
            PrivateFindByToken = True
            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.PrivateFindByToken(): " & CStr(Token) & " not found", LEMON_LEVEL_CRITICAL

End Function

Public Function CountSnarlApps() As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If mApp(i).IncludeInMenu Then _
                CountSnarlApps = CountSnarlApps + 1

        Next i
    End If

End Function

Private Function uGetBestIcon(ByVal Path As String, ByRef Destination As String) As Boolean
Dim pi As BIcon

    With new_BIconContent(Path)
        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pi) Then
            With New mfxView
                .SizeTo pi.Width, pi.Height
                .DrawScaledImage pi.Render, , new_BPoint(pi.Width, pi.Height)
                Destination = g_GetSafeTempIconPath()
                .WriteToFile Destination, "image/png"
                uGetBestIcon = True

            End With
        End If
    End With

End Function
