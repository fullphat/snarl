VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TApp.cls
    '/
    '/  Description:    Represents a single registered application
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mInfo As T_SNARL_APP        ' // app info
Dim mAlert() As TAlert          ' // class(es) registered against this app
Dim mAlerts As Long

Dim mSettings As ConfigFile     ' // persistent settings
Dim mAppConfig As ConfigSection ' // persistent app settings
Dim mIsUnknown As Boolean       ' // new app (or, one without persistent settings anyway)

    ' /* R2.4.2 */
Dim mIsBanned As Boolean        ' // app has attempted a DoS (or otherwise been banned)
Dim mLastNotificationTick As Long
Dim mDoSCount As Long

Dim mCurrentLP As Long          ' // DR3: token of current low priority notification

Public PopupFrame As BRect

Dim mCachedIcon As mfxBitmap
Dim mRealIconPath As String

    ' /* R2.5.1 */

Dim mExtension As String
'Dim mExtensionConfigurable As Boolean

Implements mObject
Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mInfo.Signature

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mInfo.Name

End Function

Private Sub Class_Initialize()

    mInfo.Name = "Snarl"
    mInfo.Signature = App.ProductName

End Sub

Private Property Get MObject_Type() As String
End Property

Friend Function bInit(ByRef Info As T_SNARL_APP) As Boolean
Dim sz As String
Dim i As Long
Dim f As Boolean

    g_Debug "TApp.bInit()", LEMON_LEVEL_PROC_ENTER

    ' /* R2.5.1: if admin setting "BlockUnknownApps" is set, check that the class file exists and fail if it doesn't */

    If (gSysAdmin.BlockUnknownApps) And (Not g_Exists(gPrefs.SnarlConfigPath & "etc\" & g_MakeFilename(Info.Signature) & ".class")) Then
        g_SetLastError SNARL_ERROR_ACCESS_DENIED
        g_Debug "blocked by admininstrative setting", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    ' /* store registration time/date */

    g_Debug "timestamp"
    Info.Timestamp = Now()

    g_Debug "info"
    LSet mInfo = Info
    mIsUnknown = False

    ' /* R2.5.1 Beta 3: if it's an extension, snag it now */

    g_Debug "checking extension"
    If NOTNULL(gCurrentExtension) Then _
        mExtension = gCurrentExtension.Class

    ' /* do we have a config file for this app? */

    g_Debug "config file"
    Set mSettings = New ConfigFile
    With mSettings
        ' /* R2.31: we know that SnarlConfigPath is valid */
        .File = gPrefs.SnarlConfigPath & "etc\" & g_MakeFilename(Info.Signature) & ".class"
        .Load

        ' /* V40.25: must have a "_cfg" section */

        i = .FindSection("_cfg")
        If i = 0 Then
            g_Debug "no config"
            Set mAppConfig = .AddSectionObj("_cfg")
            mAppConfig.Add "banned", "0"
            uWriteSettings

        Else
            g_Debug "has config"
            Set mAppConfig = .SectionAt(i)
            mIsBanned = (mAppConfig.GetValueWithDefault("banned", "0"))

        End If

        ' /* must have an "_all" section */

        mIsUnknown = Not (.SectionExists("_all"))
        g_Debug "_all >> " & CStr(mIsUnknown)

        If mIsUnknown Then _
            .AddSection "_all"

        ' /* R2.4.2 DR3: always write the application name to the config file */

        If mAppConfig.IndexOf("name") = 0 Then
            mAppConfig.Add "name", Info.Name

        Else
            mAppConfig.Update "name", Info.Name

        End If

        g_Debug "writing settings"
        uWriteSettings

    End With

    g_Debug "icon path [" & mInfo.Icon & "]"
    mRealIconPath = g_TranslateIconPath(mInfo.Icon, "")
    g_Debug "real icon path [" & mRealIconPath & "]"

    If Not g_Exists(mRealIconPath) Then _
        mRealIconPath = g_MakePath(App.Path) & "etc\icons\no_icon.png"

    g_Debug "icon"
    mInfo.Icon = mRealIconPath
    Set mCachedIcon = load_image_obj(mInfo.Icon)

    ' /* all apps get this class */

    g_Debug "adding _all"
    AddAlert "_all", ""
    bInit = True

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function AddAlert(ByVal Class As String, ByVal Description As String, Optional ByRef ClassObject As TAlert, Optional ByVal Flags As SN_CLASS_FLAGS) As M_RESULT

    g_Debug "TApp.AddAlert()", LEMON_LEVEL_PROC_ENTER

    ' /* check the class isn't already registered */
    If Me.FindAlert(Class, Nothing) Then
        g_Debug "class '" & Class & "' already exists", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_CLASS_ALREADY_EXISTS
        AddAlert = M_ALREADY_EXISTS
        Exit Function

    End If

Dim bNeedUpdate As Boolean
Dim pClassSettings As ConfigSection
Dim i As Long

    ' /* load up persistent settings if any exist */
    g_Debug "checking persistent settings..."

    With mSettings
        i = .FindSection(Class)
        If i = 0 Then
            g_Debug "no config for '" & Class & "', creating section now..."
            Set pClassSettings = .AddSectionObj(Class)
            uWriteSettings

        Else
            g_Debug "got settings for '" & Class & "'"
            Set pClassSettings = .SectionAt(i)

        End If

    End With

    ' /* add and intialize it */
    g_Debug "creating class object..."
    Set ClassObject = New TAlert
    ClassObject.bInit Class, Description, Me, pClassSettings, Flags

    mAlerts = mAlerts + 1
    ReDim Preserve mAlert(mAlerts)
    Set mAlert(mAlerts) = ClassObject
    AddAlert = M_OK

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function IndexOf(ByVal Class As String) As Long

    If (mAlerts = 0) Or (Class = "") Then _
        Exit Function

Dim i As Long

    For i = 1 To mAlerts
        If mAlert(i).Name = Class Then
            IndexOf = i
            Exit Function

        End If
    Next i

End Function

Public Function FindAlert(ByVal Class As String, ByRef Alert As TAlert) As Boolean
Dim i As Long

    i = IndexOf(Class)
    If i Then
        Set Alert = mAlert(i)
        Alert.Index = i
        FindAlert = True

    End If

End Function

Public Function CountAlerts() As Long

    CountAlerts = mAlerts

End Function

Public Function AlertAt(ByVal Index As Long) As TAlert

    If (Index > 0) And (Index <= mAlerts) Then
        Set AlertAt = mAlert(Index)
        AlertAt.Index = Index

    End If

End Function

Public Sub WriteConfig()

    uWriteSettings

End Sub

Public Sub ResetAlerts()
Dim i As Long

    If mAlerts Then
        For i = 1 To mAlerts
            mAlert(i).Reset

        Next i

    End If

End Sub

Public Function Name() As String

    Name = mInfo.Name

End Function

Public Function hWnd() As Long

    hWnd = mInfo.hWnd

End Function

Public Function uMsg() As Long

    uMsg = mInfo.uMsg

End Function

Public Function Pid() As Long

    Pid = mInfo.Pid

End Function

Public Function Icon() As String

    Icon = mInfo.Icon

End Function

'Public Function LargeIcon() As String
'
'    LargeIcon = mInfo.LargeIcon
'
'End Function

Public Function UseSameSettingsForAllClasses() As Boolean

    If Not (mAppConfig Is Nothing) Then _
        UseSameSettingsForAllClasses = Val(mAppConfig.GetValueWithDefault("use_same_settings_for_all_classes", "0"))

End Function

Public Sub SetUseSameSettingsForAllClasses(ByVal Enabled As Boolean)

    If Not (mAppConfig Is Nothing) Then
        mAppConfig.Update "use_same_settings_for_all_classes", IIf(Enabled, "1", "0")
        If Not (mSettings Is Nothing) Then _
            uWriteSettings

    End If

End Sub

Public Function Show(ByVal Class As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Timeout As Long, Optional ByVal Icon As String, Optional ByVal hWndReply As Long, Optional ByVal uMsg As Long, Optional ByVal SoundFile As String, Optional ByVal Style As String, Optional ByVal Ack As String, Optional ByVal APIVersion As Long) As M_RESULT
Dim pc As TAlert

    ' /* V39.86: wasn't taking account of the "Use Same Settings..." option */

    If Me.UseSameSettingsForAllClasses Then
        Set pc = mAlert(1)

    ElseIf Not Me.FindAlert(Class, pc) Then
        ' /* is the class registered? */
        g_Debug "TApp.Show(): class '" & Class & "' not registered, using default...", LEMON_LEVEL_WARNING
        If mAlerts > 0 Then _
            Set pc = mAlert(1)              ' // default/catch-all...

    End If

    ' /* here, we must have a class... */

    If (pc Is Nothing) Then
        g_Debug "TApp.Show(): failed - no class", LEMON_LEVEL_CRITICAL
        Show = M_FAILED
        Exit Function

    End If

Dim pInfo As T_NOTIFICATION_INFO

    With pInfo
        ' /* alert class won't know about these two */

        .hWndReply = hWndReply
        If .hWndReply = 0 Then _
            .hWndReply = mInfo.hWnd

        .uReplyMsg = uMsg
        If .uReplyMsg = 0 Then _
            .uReplyMsg = mInfo.uMsg

        .Pid = mInfo.Pid
        .Title = Title
        .Text = Text
        .Timeout = Timeout
        .IconPath = Icon
        .SndFile = SoundFile
        .StyleName = style_GetStyleName(Style)
        .SchemeName = style_GetSchemeName(Style)
        .DefaultAck = Ack
        .APIVersion = APIVersion

    End With

    ' /* pass it to the class... */

    Show = pc.ShowNotification(pInfo, Nothing)
    If Show = 0 Then _
        Show = M_FAILED

'    Show = pc.ShowNotification(mInfo.pid, Title, Text, Timeout, Icon, hWndReply, uMsg, SoundFile, Style, Ack)

End Function

Public Function Show41(ByVal ClassId As String, ByRef pData As BPackedData, Optional ByRef ReplySocket As CSocket, Optional ByVal IntFlags As SN_NOTIFICATION_FLAGS) As Long '//, Optional ByVal RemoteHost As String) As Long

    ' /* this is where we transition from a BPackedData into a T_NOTIFICATION_INFO struct */

    If (pData Is Nothing) Then _
        Exit Function

    ' /* R2.4.2: check app is not banned */

    If uCheckDOS() Then _
        Exit Function

Dim pClass As TAlert

    ' /* V39.86: wasn't taking account of the "Use Same Settings..." option */

    If Me.UseSameSettingsForAllClasses Then
        ' /* use the All/Other class definitions */
        Set pClass = mAlert(1)

    ElseIf Not Me.FindAlert(ClassId, pClass) Then
        ' /* is the class registered? */
        g_Debug "TApp.Show(): class '" & ClassId & "' not registered, using default...", LEMON_LEVEL_WARNING
        If mAlerts > 0 Then _
            Set pClass = mAlert(1)              ' // default/catch-all...

        ' /* to-do: block notifications with invalid class names, especially those
        '    sourcing from remotely registered apps */

    End If

    ' /* here, we must have a class... */

    If (pClass Is Nothing) Then
        g_Debug "TApp.Show41(): couldn't find a suitable class", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

    ' /* R2.4.2: if "Don't show if foreground app?" is set, check now */

    If pClass.GetValue("block_if_foreground") = "1" Then
        If (mInfo.Pid <> 0) And (uGetForegroundPID() = mInfo.Pid) Then
            g_Debug "TApp.Show41(): discarded: application is foreground", LEMON_LEVEL_INFO
            g_SetLastError SNARL_ERROR_DISCARDED
            Exit Function

        End If
    End If


    ' /* NEW: for long-term work - construct a TNotificationInfo object */

'Dim pni As TNotificationInfo
'
'    Set pni = New TNotificationInfo
'    pni.SetTo Me, pClass, pData, ReplySocket, IntFlags


    ' /* convert the packed data into a T_NOTIFICATION_INFO struct */

Dim pInfo As T_NOTIFICATION_INFO
Dim i As Long

    With pInfo

        ' /* R2.4 DR7: "callback" is now preferred over "ack" */

        If pData.Exists("callback") Then
            .DefaultAck = pData.ValueOf("callback")

        Else
            .DefaultAck = pData.ValueOf("ack")

        End If

        ' /* R2.4 DR7: new "uid" arg */

        If pData.Exists("uid") Then _
            .CustomUID = pData.ValueOf("uid")

        ' /* R2.4 DR7: new "value" arg */

        .Value = pData.ValueOf("value")

        ' /* R2.4 DR7: internal-only args (can only be set at g_DoNotify() level) */

        .IntFlags = IntFlags

        ' /* R2.4 DR7: static copy of sender BUT only if it's not local */

'        Debug.Print "++ " & RemoteHost & " : " & get_ip_address_table()
'        If InStr(get_ip_address_table(), RemoteHost) = 0 Then _
'            .RemoteHostName = RemoteHost

        ' /* flags: deprecated */

        If pData.Exists("flags") Then
            i = Val("&H" & pData.ValueOf("flags"))
            .Flags = (i And &HFFFF&)                ' // only keep user flags

        End If

        ' /* R2.4 DR7: "reply-to" is now preferred over "hwnd" */

        .hWndReply = Val(pData.ValueOf("reply-to"))
        If .hWndReply = 0 Then _
            .hWndReply = Val(pData.ValueOf("hwnd"))

        If .hWndReply = 0 Then _
            .hWndReply = mInfo.hWnd                 ' // if still no reply window, use app-defined value (which may still be 0)

        
        If pData.Exists("icon-base64") Then
            ' /* R2.4.2 DR3: use Base64 encoded icon data, if there is any */
            .IconPath = g_GetBase64Icon(pData.ValueOf("icon-base64"))

        ElseIf pData.Exists("icon-phat64") Then
            ' /* R2.4.2 DR3: phat64 encoding? */
            .IconPath = g_GetPhat64Icon(pData.ValueOf("icon-phat64"))

        Else
            .IconPath = pData.ValueOf("icon")

        End If
        

        .Priority = Val(pData.ValueOf("priority"))
        Set .Socket = ReplySocket

        .Text = Replace$(pData.ValueOf("text"), "\n", vbCrLf)

        .Timeout = -1                               ' // pre-set with default
        If pData.Exists("timeout") Then _
            .Timeout = Val(pData.ValueOf("timeout"))

        .Title = Replace$(pData.ValueOf("title"), "\n", vbCrLf)

        ' /* R2.4 DR7: "reply" is now preferred over "umsg" */

        .uReplyMsg = Val(pData.ValueOf("reply"))
        If .uReplyMsg = 0 Then _
            .uReplyMsg = Val(pData.ValueOf("umsg"))

        If .uReplyMsg = 0 Then _
            .uReplyMsg = mInfo.uMsg                 ' // if still no reply message, use app-defined value (which may still be 0)


        ' /* R2.4.2: caller can now request a particular style - if security option allows it */

        If (pData.Exists("style")) And (g_ConfigGet("ignore_style_requests") = "0") Then
            .StyleName = style_GetStyleName(pData.ValueOf("style"))
            .SchemeName = style_GetSchemeName(pData.ValueOf("style"))

        End If

        ' /* 2.4.2 DR3 - handle "sound" */

        If pData.Exists("sound") Then _
            .SndFile = pData.ValueOf("sound")

        ' /* 2.4.2 DR3 - store API version */

        .APIVersion = (IntFlags And SN_NF_API_MASK)


        ' /* R2.5 Beta 2 - script callbacks */

        If (pData.Exists("callback-script")) And (pData.Exists("callback-script-type")) Then
            .ScriptFilename = pData.ValueOf("callback-script")
            .ScriptLanguage = pData.ValueOf("callback-script-type")

        End If

        ' /* V42 styles need the unabridged content */

        .OriginalContent = Replace$(pData.AsString(), "\n", vbCrLf)

    End With

    ' /* pass it to the class... */

    Show41 = pClass.ShowNotification(pInfo, pData)

End Function


'Friend Function ShowNew(ByVal Class As String, ByRef Info As T_NOTIFICATION_INFO) As Long
'Dim pc As TAlert
'
'    ' /* V39.86: wasn't taking account of the "Use Same Settings..." option */
'
'    If Me.UseSameSettingsForAllClasses Then
'        Set pc = mAlert(1)
'
'    ElseIf Not Me.FindAlert(Class, pc) Then
'        ' /* is the class registered? */
'        g_Debug "TApp.Show(): class '" & Class & "' not registered, using default...", LEMON_LEVEL_WARNING
'        If mAlerts > 0 Then _
'            Set pc = mAlert(1)              ' // default/catch-all...
'
'    End If
'
'    ' /* here, we must have a class... */
'
'    If (pc Is Nothing) Then
'        g_Debug "TApp.Show(): couldn't find a suitable class", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_FAILED
'        Exit Function
'
'    End If
'
'    ' /* alert class won't know about these two */
'
'    If Info.hWndReply = 0 Then _
'        Info.hWndReply = mInfo.hWnd
'
'    If Info.uReplyMsg = 0 Then _
'        Info.uReplyMsg = mInfo.uMsg
'
'    ' /* secure */
'
''    If mInfo.Password <> "" Then _
'        Info.Flags = Info.Flags Or SNARL42_NOTIFICATION_SECURE
'
'    ' /* pass it to the class... */
'
'    ShowNew = pc.ShowNotification(Info)
'
'End Function

Public Function Token() As Long

    Token = mInfo.Token

End Function

Public Sub SetIcon(ByVal Path As String)

    mInfo.Icon = Path

End Sub

Public Function AddClass(ByRef Data As BPackedData) As Long

'//ByVal Name As String, ByVal Flags As E_ALERT_CLASS_FLAGS, Optional ByRef ClassObject As TAlert

    ' /* validate some key stuff */
    If (Data Is Nothing) Then
        g_Debug "TApp.AddClass(): config error", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_FAILED
        Exit Function

    End If

Dim szId As String

    ' /* R2.5 Beta 2: support for adding classes via NOTIFY */
    If Data.Exists("class-id") Then
        szId = Data.ValueOf("class-id")

    ElseIf Data.Exists("id") Then
        szId = Data.ValueOf("id")

    End If

    ' /* arg checking */

    If szId = "" Then
        g_Debug "TApp.AddClass(): required arg 'class-id/id' missing", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

Dim szName As String

    ' /* R2.5 Beta 2: support for adding classes via NOTIFY */
    If Data.Exists("class-name") Then
        szName = Data.ValueOf("class-name")

    ElseIf Data.Exists("name") Then
        szName = Data.ValueOf("name")

    End If

    ' /* name isn't mandatory so set it from id */

    If szName = "" Then _
        szName = szId

    ' /* enabled/disabled (backwards compatability with pre-release editions) */

Dim lFlags As SN_CLASS_FLAGS

    If Data.Exists("enabled") Then
        If Data.ValueOf("enabled") = "0" Then _
            lFlags = lFlags Or SN_CF_DISABLED

    End If

    ' /* Use AddAlert() to actually add the class to our app - LastError will be set on failure */

Dim pClass As TAlert

    If AddAlert(szId, szName, pClass, lFlags) <> M_OK Then _
        Exit Function

    ' /* do V41 post-initialization */
    pClass.Augment Data

    frmAbout.bUpdateClassList mInfo.Token
    AddClass = -1

End Function

Public Function RemClass(ByRef Data As BPackedData, Optional ByVal RemoveAll As Boolean = False) As Long
Static i As Long
Static j As Long

    ' /* validate some key stuff */

    If (Data Is Nothing) Then
        g_Debug "TApp.RemClass(): config error", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_FAILED
        Exit Function

    End If

    ' /* special case: if the 'all' tag is present and it's set to "1" then
    '    remove all classes */

    If (Data.ValueOf("all") = "1") Or (RemoveAll) Then
        g_Debug "TApp.RemClass(): removing all classes..."
        mAlerts = 0
        ReDim mAlert(0)
        AddAlert "_all", ""
        RemClass = -1
        frmAbout.bUpdateClassList mInfo.Token
        g_SetLastError SNARL_SUCCESS

    ElseIf Not Data.Exists("id") Then
        g_Debug "TApp.RemClass(): required arg 'id' missing", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_ARG_MISSING

    ElseIf Data.ValueOf("id") = "_all" Then
        g_Debug "TApp.RemClass(): cannot remove '_all' class", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_FAILED

    Else
        i = uIndexOf(Data.ValueOf("id"))
        If i = 0 Then
            g_Debug "TApp.RemClass(): class '" & Data.ValueOf("id") & "' not found", LEMON_LEVEL_CRITICAL
            g_SetLastError SNARL_ERROR_CLASS_NOT_FOUND

        Else

            If i < mAlerts Then
                For j = i To (mAlerts - 1)
                    Set mAlert(j) = mAlert(j + 1)

                Next j

            End If

            mAlerts = mAlerts - 1
            ReDim Preserve mAlert(mAlerts)
            frmAbout.bUpdateClassList mInfo.Token
            RemClass = -1
            g_SetLastError SNARL_SUCCESS

        End If
    End If

End Function

Public Function Signature() As String

    Signature = mInfo.Signature

End Function

Public Function Flags() As SNARLAPP_FLAGS

    Flags = mInfo.Flags

End Function

Public Function IncludeInMenu() As Boolean

    IncludeInMenu = mInfo.IncludeInMenu

End Function

Public Function HasConfig() As Boolean
Dim pe As TExtension

    If mExtension <> "" Then
        If g_ExtnRoster.Find(mExtension, pe) Then
            HasConfig = pe.IsConfigurable()

        Else
            g_Debug "TApp.HasConfig(): extension " & g_Quote(mExtension) & " not found", LEMON_LEVEL_CRITICAL

        End If

    Else
        HasConfig = (mInfo.Tool <> "") Or _
                    ((Not (mInfo.Socket Is Nothing)) And (mInfo.IncludeInMenu)) Or _
                    ((mInfo.Flags And SNARLAPP_HAS_PREFS) <> 0)

    End If

End Function

Public Function HasAbout() As Boolean
Dim pe As TExtension

    If mExtension <> "" Then
        If g_ExtnRoster.Find(mExtension, pe) Then _
            HasAbout = True

    ElseIf mInfo.RemoteHostName <> "" Then
        HasAbout = True

    Else
        HasAbout = (mInfo.Token = gSnarlToken) Or (mInfo.Hint <> "") Or _
                   ((Not (mInfo.Socket Is Nothing)) And (mInfo.IncludeInMenu)) Or _
                   ((mInfo.Flags And SNARLAPP_HAS_ABOUT) <> 0)

    End If

End Function

Public Function IsQuittable() As Boolean

    IsQuittable = (IsDaemon()) Or (mExtension <> "")

End Function

Public Function IsDaemon() As Boolean

    IsDaemon = ((mInfo.Flags And SNARLAPP_IS_DAEMON) <> 0)

End Function

Public Function IsSecure() As Boolean

    IsSecure = (mInfo.Password <> "")

End Function

Public Sub DoAbout()
Dim fApp As SN_NOTIFICATION_FLAGS
Dim pe As TExtension

    If mInfo.Token = gSnarlToken Then
        ' /* ourselves */
        frmAbout.Show

    ElseIf mExtension <> "" Then
        '/ * extension */
        If g_ExtnRoster.Find(mExtension, pe) Then _
            pe.ShowDetails

    ElseIf mInfo.RemoteHostName <> "" Then

        g_PrivateNotify , _
                        mInfo.Name & " on " & mInfo.RemoteHostName, _
                        mInfo.Hint, , _
                        mInfo.Icon, , , , _
                        SN_NF_REMOTE, , _
                        "_about"


    ElseIf mInfo.Hint <> "" Then
        ' /* display the hint text */
        g_Debug "TApp.DoAbout(): displaying static hint text..."

'        If mInfo.IsRemote Then _
'            fApp = fApp Or SN_NF_REMOTE

        If mInfo.Password <> "" Then _
            fApp = fApp Or SN_NF_SECURE

        g_PrivateNotify , _
                        mInfo.Name, _
                        mInfo.Hint, , _
                        mInfo.Icon, , , , _
                        fApp, , _
                        "_about"

    ElseIf (Not (mInfo.Socket Is Nothing)) Then
        ' /* SNP-registered (V43 only) */
        g_Debug "TApp.DoAbout(): local SNP3 (assumed) app..."
        g_SNP3SendCallback mInfo.Socket, SNARL_NOTIFY_APP_DO_ABOUT, "DoAbout", "", ""

    ElseIf (IsWindow(mInfo.hWnd) <> 0) And ((mInfo.Flags And SNARLAPP_HAS_ABOUT)) Then
        ' /* Win32-registered */
        g_Debug "TApp.DoAbout(): posting SNARLAPP_DO_ABOUT to reply-to window..."
        PostMessage mInfo.hWnd, WM_NULL, 0, ByVal 0&
        PostMessage mInfo.hWnd, snAppMsg(), SNARLAPP_DO_ABOUT, ByVal 0&

        ' /* just show what we can */
'        g_PrivateNotify "", _
'                        mInfo.Name, _
'                        "Registered: " & Format$(mInfo.Timestamp, "d mmm yyyy hh:mm:ss") & _
'                        IIf(mInfo.Password <> "", "\nPassword protected", "") & _
'                        IIf(mInfo.IsRemote, "\nSource is remote", "") & _
'                        IIf(mIsBanned, "\n\nIs banned", ""), , _
'                        mInfo.Icon, , , , , , _
'                        "_about"

    End If

End Sub

Public Sub DoSettings(ByVal hWndPanel As Long)

    g_Debug "TApp.DoSettings()", LEMON_LEVEL_PROC_ENTER

Dim pe As TExtension

    If mExtension <> "" Then
        g_Debug "is an extension...", LEMON_LEVEL_INFO

        If hWndPanel <> 0 Then
            If g_ExtnRoster.Find(mExtension, pe) Then
                If pe.IsConfigurable Then
                    g_Debug "showing config..."
                    pe.DoPrefs hWndPanel
    
                Else
                    g_Debug "extension " & g_Quote(mExtension) & " is not configurable", LEMON_LEVEL_WARNING

                End If
    
            Else
                g_Debug "extension " & g_Quote(mExtension) & " not found", LEMON_LEVEL_CRITICAL
            
            End If
        
        Else
            g_Debug "can't show extension GUI without owner window", LEMON_LEVEL_CRITICAL
        
        End If

    ElseIf mInfo.Tool <> "" Then
        ' /* run the tool */
        g_Debug "running tool '" & mInfo.Tool & "'..."
        WinExec mInfo.Tool, SW_SHOW

    ElseIf Not (mInfo.Socket Is Nothing) Then
        ' /* SNP-registered (V43 only) */
        g_Debug "local SNP3 (assumed) app..."
        g_SNP3SendCallback mInfo.Socket, SNARL_NOTIFY_APP_DO_PREFS, "DoPrefs", "", ""

    ElseIf (IsWindow(mInfo.hWnd) <> 0) And ((mInfo.Flags And SNARLAPP_HAS_PREFS)) Then
        ' /* Win32-registered */
        g_Debug "posting SNARLAPP_DO_PREFS to reply-to window..."
        PostMessage mInfo.hWnd, WM_NULL, 0, ByVal 0&
        PostMessage mInfo.hWnd, snAppMsg(), SNARLAPP_DO_PREFS, ByVal 0&

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Sub DoQuitPlease()
Dim pe As TExtension

    If mExtension <> "" Then
        If g_ExtnRoster.Find(mExtension, pe) Then
            pe.SetEnabled False
            frmAbout.bUpdateExtList

        End If

    ElseIf (IsWindow(mInfo.hWnd) <> 0) And ((mInfo.Flags And SNARLAPP_IS_DAEMON)) Then
        ' /* Win32-registered */
        g_Debug "TApp.DoAbout(): posting SNARLAPP_QUIT_REQUESTED to reply-to window..."
        PostMessage mInfo.hWnd, WM_NULL, 0, ByVal 0&
        PostMessage mInfo.hWnd, snAppMsg(), SNARLAPP_QUIT_REQUESTED, ByVal 0&

    End If

End Sub

Public Sub Activated()

    If mInfo.Tool <> "" Then
        g_Debug "TApp.Activated(): running static tool '" & mInfo.Tool & "'..."
        If g_SafeLeftStr(mInfo.Tool, 1) = "!" Then
            g_ProcessAck mInfo.Tool

        Else
            WinExec mInfo.Tool, SW_SHOW

        End If

    ElseIf Not (mInfo.Socket Is Nothing) Then
        g_Debug "TApp.DoSettings(): local SNP3 (assumed) app..."
        g_SNP3SendCallback mInfo.Socket, SNARL_NOTIFY_APP_ACTIVATED, "Activated", "", ""

    ElseIf (mInfo.Flags And SNARLAPP_IS_DAEMON) <> 0 Then
        ' /* V43 */
        g_Debug "TApp.DoSettings(): posting SNARLAPP_ACTIVATED to reply-to window..."
        PostMessage mInfo.hWnd, WM_NULL, 0, ByVal 0&
        PostMessage mInfo.hWnd, snAppMsg(), SNARLAPP_ACTIVATED, ByVal 0&

    Else
        ' /* pre-V43 */
        g_Debug "TApp.DoSettings(): posting SNARLAPP_DO_PREFS to reply-to window..."
        PostMessage mInfo.hWnd, WM_NULL, 0, ByVal 0&
        PostMessage mInfo.hWnd, snAppMsg(), SNARLAPP_DO_PREFS, ByVal 0&

    End If

End Sub

Private Function uIndexOf(ByVal Id As String) As Long

    If mAlerts = 0 Then _
        Exit Function

Dim i As Long

    For i = 1 To mAlerts
        If mAlert(i).Name = Id Then
            uIndexOf = i
            Exit Function

        End If
    Next i

End Function

Private Sub uWriteSettings()

    If (mSettings Is Nothing) Then _
        Exit Sub

    If gSysAdmin.TreatSettingsAsReadOnly Then
        g_Debug "TApp.uWriteSettings(): blocked by admin setting"

    Else
        mSettings.Save

    End If

End Sub

Public Function Password() As String

    Password = mInfo.Password

End Function

Public Function IsRemote() As Boolean

    ' /* returns TRUE if the remote application is _truly_ remote */
    
    If mInfo.RemoteHostName <> "" Then _
        IsRemote = Not g_IsLocalAddress(mInfo.RemoteHostName)

End Function

Public Function IsNew() As Boolean

    IsNew = mIsUnknown

End Function

'Public Sub SetLastError(ByVal err As Long)
'
'    mLastError = err
'
'End Sub
'
'Public Function LastError() As Long
'
'    LastError = mLastError
'
'End Function

Public Sub SetReply(ByVal ReplyTo As Long, ByVal ReplyMessage As Long)

    If ReplyTo <> 0 Then _
        mInfo.hWnd = ReplyTo

    If ReplyMessage <> 0 Then _
        mInfo.uMsg = ReplyMessage

End Sub

Private Function uGetForegroundPID() As Long
Dim h As Long
Dim p As Long

    h = GetForegroundWindow()
    If h = 0 Then _
        Exit Function

    If GetWindowThreadProcessId(h, p) <> 0 Then _
        uGetForegroundPID = p

End Function

Public Function IsBanned() As Boolean

    IsBanned = mIsBanned

End Function

Private Function uCheckDOS() As Boolean

    If mIsBanned Then
        g_Debug "TApp.uCheckDOS(): '" & mInfo.Signature & "' is banned", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_ACCESS_DENIED
        uCheckDOS = True

    ElseIf g_NotificationRoster.CountActiveBySignature(mInfo.Signature) > 9 Then
        g_Debug "TApp.uCheckDOS(): '" & mInfo.Signature & "' has too many active notifications", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_FLOODING
        uCheckDOS = True

    Else
        g_SetLastError SNARL_SUCCESS

    End If

'    ' /* R2.5 Beta 2: ignore if app's "never_ban" setting is enabled */
'
'    If (g_ConfigGet("block_dos_attempt") = "0") Or (mAppConfig.GetValueWithDefault("never_ban", "0") = "1") Then _
'        Exit Function
'
'Const BAN_NUMBER = 10                           ' // this many...
'Const BAN_LIMIT = 500                           ' // ...within this time (in ms) = a ban
'Const BAN_QUANTUM = BAN_LIMIT / BAN_NUMBER      ' // so this is the per-notification trigger
'
'    If mIsBanned Then
'        g_Debug "TApp.uCheckDOS(): '" & mInfo.Signature & "' is banned", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_ACCESS_DENIED
'
'    Else
'        ' /* if we get BAN_NUMBER within BAN_QUANTUM, that's a ban */
'
'        If (GetTickCount() - mLastNotificationTick) < BAN_QUANTUM Then
'            mDoSCount = mDoSCount + 1
'            g_Debug "TApp.uCheckDOS(): notification within ban quantum (" & GetTickCount() - mLastNotificationTick & ") count=" & mDoSCount
'
'            If mDoSCount >= BAN_NUMBER Then
'                g_Debug "TApp.uCheckDOS(): possible DoS attack by '" & mInfo.Signature & "' detected", LEMON_LEVEL_CRITICAL
'                If gDebugMode Then _
'                    g_PrivateNotify SNARL_CLASS_GENERAL, "Application blocked", mInfo.Name & " (" & mInfo.Signature & ") generated a large number of notifications within a short period of time and has been blocked"
'
'                mIsBanned = True
'                mAppConfig.Update "banned", "1"
'                uWriteSettings
'
'            End If
'
'        Else
'            ' /* reset */
'            mDoSCount = 0
'
'        End If
'
'        mLastNotificationTick = GetTickCount()
'
'    End If
'
'    uCheckDOS = mIsBanned

End Function

Public Sub SetBlocked(ByVal IsBlocked As Boolean)

    mIsBanned = IsBlocked
    mAppConfig.Update "banned", IIf(IsBlocked, "1", "0")
    uWriteSettings

End Sub

Public Sub SetCurrentLowPriority(ByVal Token As Long)

    ' /* R2.4.2 DR3: each app can have only one low priority notification on screen at once */

    If mCurrentLP <> 0 Then _
        g_NotificationRoster.PrivateHide mCurrentLP

    mCurrentLP = Token

End Sub

Public Sub RemoveSettings()

    If (mSettings Is Nothing) Then _
        Exit Sub

    g_Debug "TApp.RemoveSettings(): deleting '" & mSettings.File & "'..."
    DeleteFile mSettings.File

End Sub

Public Function CachedIcon() As MImage

    Set CachedIcon = mCachedIcon

End Function

Public Function AppType() As SN_APP_TYPES

    AppType = mInfo.AppType

End Function

Public Function NameEx() As String

    NameEx = mInfo.Name

    If mInfo.RemoteHostName = "" Then _
        Exit Function

    If Not g_IsLocalAddress(mInfo.RemoteHostName) Then _
        NameEx = NameEx & " on " & mInfo.RemoteHostName

End Function

Public Function IsLocal() As Boolean

    ' /* returns TRUE if app is network-based but registered locally */

    If Not (mInfo.Socket Is Nothing) Then _
        IsLocal = Not Me.IsRemote

End Function

Public Function RemoteHostName() As String

    RemoteHostName = mInfo.RemoteHostName

End Function

Friend Sub PrivateGetInfo(ByRef Info As T_SNARL_APP)

    LSet Info = mInfo

End Sub

Public Function KeepAlive() As Boolean

    KeepAlive = mInfo.KeepAlive

End Function

Public Function RealIconPath() As String

    RealIconPath = mRealIconPath

End Function

Public Sub TestNotification(ByRef Class As TAlert)
Dim ppd As BPackedData

    Set ppd = New BPackedData
    With ppd
        .Add "title", Class.Description
        .Add "text", "Test Notification"

    End With
    
    Me.Show41 Class.Name, ppd, , App.Major


'Dim pInfo As T_NOTIFICATION_INFO
'
'    With pInfo
'        .Title = Me.Description
'        .Text = "Test Notification"
'        .Timeout = -1
'        .APIVersion = VB.App.Major
'        .IntFlags = VB.App.Major
'        Set .ClassObj = Me
'.ClassObj.App.Show41
'
'    End With
'
'    Me.ShowNotification pInfo, Nothing
End Sub
