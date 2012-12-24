Attribute VB_Name = "mSnarl"
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           mSnarl.bas
    '/
    '/  Description:    Global functions and declarations
    '/
    '/  © 2004-2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* these are used by the deprecated SNARL_GET_VERSION and for GNTP responses */
Public Const APP_VER = 3
Public Const APP_SUB_VER = 0
Public Const APP_SUB_SUB_VER = 0

Public Const SNP_VERSION = "3.0"

Public Const GNTP_DEFAULT_PORT = 23053
Public Const SNP_DEFAULT_PORT = 9887
Public Const JSON_DEFAULT_PORT = 9889
Public Const MELON_DEFAULT_PORT = 5233

    ' /* private V43 Snarl App flags */
Public Const SNARLAPP_IS_DAEMON = &H4000&


Public Enum SOS_ERRORS
    SOS_UNSPECIFIED_FAILURE = &H40          '// reserved
    SOS_BAD_COPYDATA
    SOS_MISSING_ROSTER                      '// one of the rosters is unavailable
    SOS_SPURIOUS_MANAGE                     '// reserved for unhandled WM_MANAGESNARL
    SOS_SPURIOUS_TEST                       '// unhandled WM_SNARLTEST
    SOS_SPURIOUS_COMMAND                    '// reserved for unhandled WM_SNARL_COMMAND

    SOS_FILE_NOT_FOUND = &H50               '// critical file error
    SOS_PATH_NOT_FOUND                      '// critical path error

End Enum


Public Declare Sub CoFreeUnusedLibrariesEx Lib "ole32" (ByVal dwUnloadDelay As Long, ByVal dwReserved As Long)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function LockWorkStation Lib "user32.dll" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Sub ShellAbout Lib "SHELL32.DLL" Alias "ShellAboutA" (ByVal hWndOwner As Long, ByVal lpszAppName As String, ByVal lpszMoreInfo As String, ByVal hIcon As Long)
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Const WINDOW_CLASS = "w>Snarl"

Public Const WM_SNARL_INIT = WM_USER + 1
Public Const WM_SNARL_QUIT = WM_USER + 2
Public Const WM_SNARL_TRAY_ICON = WM_USER + 3
'Public Const WM_SNARL_NOTIFY_RUNNING = WM_USER + 4

'Public Const WM_REMOTENOTIFY = WM_USER + 9              ' // frmAbout: remote notifications
Public Const WM_INSTALL_SNARL = WM_USER + 12            ' // frmAbout: snarl update available


Public Const WM_SNARL_COMMAND = WM_USER + 80
    ' /* WM_SNARL_COMMAND expanded in R2.3 (renamed in R2.5 Beta 2) to use wParam and lParam, as follows:
    '       wParam                                                          lParam
    '          0        Display Preferences Panel                           0
    '          1        Install style or extension                          atom of registered string containing style or extension
    '          2        Configure style or extension                        ("")
    ' */

Public Enum SN_DO_PREFS
    SN_DP_DO_PREFS
    SN_DP_INSTALL
    SN_DP_CONFIGURE
    SN_DP_RESTART
    SN_DP_UNLOAD
    SN_DP_LOAD
    SN_DP_RESTART_STYLE_ROSTER
    ' /* R2.5.1 */
    SN_DP_SHOW_ABOUT
    SN_DP_SHOW_INFO

End Enum


Public Const TIMER_UPDATES = 32

    ' /* Snarl app class id's */

Public Const SNARL_CLASS_GENERAL = "_WLC"
Public Const SNARL_CLASS_APP_UNREG = "_APU"
Public Const SNARL_CLASS_APP_REG = "_APR"
Public Const SNARL_CLASS_JSON = "_ANJ"
Public Const SNARL_CLASS_ANON_NET = "_ANN"
Public Const SNARL_CLASS_ANON = "_ANL"
'Public Const SNARL_CLASS_LOW_PRIORITY = "_LOW"
'Public Const SNARL_CLASS_SYSTEM = "_SYS"
Public Const SNARL_CLASS_SYSTEM = "_SYS"

    ' /* internal notification flags */

Public Enum SN_NOTIFICATION_FLAGS
    SN_NF_REMOTE = &H80000000
    SN_NF_SECURE = &H40000000
    SN_NF_IS_GNTP = &H20000000         ' // R2.4.2: GNTP-based notification
    SN_NF_IS_SNP3 = &H10000000         '// R2.4.2 DR3: SNP3
    SN_NF_FORWARD = &H8000000           ' // R2.5 Beta 2: notification has been forwarded
    SN_NF_TEMP_ICON = &H4000000         ' // R2.6A4: icon is temporary

    SN_NF_MERGE = &H1000

    ' /* bottom 8 bits used for api version (V42 onwards) */
    SN_NF_API_MASK = &HFF&

End Enum

    ' /* master notification structure, as used by the notification roster */

Public Type T_NOTIFICATION_INFO
    Pid As Long
    Title As String
    Text As String
    Timeout As SN_NOTIFICATION_DURATION
    IconPath As String
    hWndReply As Long
    uReplyMsg As Long
    SndFile As String
'    StyleToUse As String
    StyleName As String                 ' // Split
    SchemeName As String                ' // Split
    DefaultAck As String                ' // Known as "callback" from R2.4 DR7
    Position As SN_START_POSITIONS
    Token As Long
    ' /* V41 */
    Priority As Long                    ' // V41: <0 = low, 0 = normal, >0 = high
    Value As String                     ' // V41: freeform value which will negate the need to use the Text field
                                        '         thoughts are the value can encapsulate the format it's sent in
                                        '         e.g. 45%, 2.3466, $5.00, etc. it's up to the style to determine
                                        '         how/if it's displayed
    DateStamp As Date                   ' // V41: when it was added to the Notification Roster
'    Icon As mfxBitmap                   ' // V41 (R2.31): note it's an mfxBitmap, not an MImage!
'    Sender As String
'    Class As String
    ' /* V42 */
    Flags As SNARL41_NOTIFICATION_FLAGS ' // V41 (R2.4): new flags
    OriginalContent As String           ' // V41 (R2.4): as passed from external source
    LastUpdated As Date                 ' // time last changed
    Socket As CSocket                   ' // reply socket (SNP2.0 native only)
    IntFlags As SN_NOTIFICATION_FLAGS    ' // internal notification flags
'    RemoteHostName As String            ' // sender (as string) for remote connections that do not have reply sockets
'                                        ' // R2.4.2 DR3: only set if sender is truly remote (i.e. not in our IP table)
    ClassObj As TAlert                  ' // object
    CustomUID As String                 ' // R2.4 DR7: custom UID (set during <notify>)
    Actions As BTagList                 ' // R2.4 DR7: should have been here all along
    APIVersion As Long                  ' // R2.4.1: will be 42 for V42, 0 for everything prior to it
    ' /* V43 */
    ScriptFilename As String
    ScriptLanguage As String
    ' /* V44 */
    AckButtonLabel As String
'    Content As BPackedData

End Type

Public Type T_SNARL_STYLE_ENGINE_INFO
    Name As String
    Version As Long
    Revision As Long
    Date As String
    Path As String                  ' // path to the engine's dll
'    Copyright As String
    Description As String
    Obj As IStyleEngine
    Flags As Long                   ' // bit 31 set means internal

End Type

Public Enum SN_APP_TYPES
    SN_AT_UNKNOWN
    SN_AT_WIN32
    SN_AT_SNP
    SN_AT_GNTP
    SN_AT_GROWLNET

End Enum

    ' /* internal registered application structure */

Public Type T_SNARL_APP
    Name As String
    hWnd As Long
    uMsg As Long
    Pid As Long                 ' // V38 (for V39)
    Icon As String              ' // R1.6 - path to application icon (if empty we use window icon)
'    LargeIcon As String         ' // V38 (private for now) - path to large icon
    Token As Long               ' // V41
    Signature As String         ' // V41 - MIME string
    Flags As SNARLAPP_FLAGS     ' // V41
    Password As String          ' // V42 - non-persistent (so the app can generate a new one each time)
'    IsRemote As Boolean         ' // V42 - remotely registered
    Timestamp As Date           ' // V42.21: set when added
    Tool As String              ' // R2.4.2 DR3: path to static configuration tool
    Hint As String              ' // R2.4.2 DR3: text to display in Details...
    Socket As CSocket           ' // R2.4.2 DR3: sender's socket
    IncludeInMenu As Boolean    ' // R2.4.2 DR3: should be included in "Apps" submenu
    AppType As SN_APP_TYPES     ' // R2.4.2 DR3: set during init()
    RemoteHostName As String    ' // R2.4.2 DR3: static, used by NameEx()
    ' /* R2.5.1 */
    KeepAlive As Boolean        ' // for Win32, don't kill if PID disappears; for SNP, don't kill if socket unregisters

End Type

    ' /* internal Snarl admin structure */

Public Type T_SNARL_ADMIN
    HideIcon As Boolean                 ' // hides the tray icon (over-rules undoc'd setting in .snarl file)
    InhibitPrefs As Boolean             ' // completely blocks access to prefs panel
    InhibitQuit As Boolean              ' // can't quit Snarl using menu
    InhibitMenu As Boolean              ' // right-click tray icon does nothing
    TreatSettingsAsReadOnly As Boolean  ' // don't write settings
    BlockUnknownApps As Boolean         ' // don't allow new apps to register
    NoRunStartupSequence As Boolean
    NoDebugMode As Boolean

End Type

Public gSysAdmin As T_SNARL_ADMIN
Public gExtDetailsToken As Long
Public gStyleEngineDetailsToken As Long

'Public Type T_SNARL_ICON_THEME
'    Name As String
'    Path As String
'    IconFile As String
'
'End Type
'
'Public gIconTheme() As T_SNARL_ICON_THEME
'Public gIconThemes As Long

Private Const SPI_GETFONTSMOOTHING = 74
Private Const SPI_GETFONTSMOOTHINGTYPE = 8202
Private Const FE_FONTSMOOTHINGSTANDARD = 1
Private Const FE_FONTSMOOTHINGCLEARTYPE = 2


Public Const HWND_SNARL = &H534E524C Or &H80000000

'Public bm_Menu As MImage
Public bm_CloseGadget As MImage
Public bm_ActionsGadget As MImage
Public bm_HasActions As MImage
Public bm_Remote As MImage
Public bm_Secure As MImage
Public bm_IsSticky As MImage
Public bm_Priority As MImage
Public bm_Forward As MImage

Public bm_Button As mfxBitmap
Public bm_CallbackButton As mfxBitmap

Public Enum SN_START_POSITIONS
    ' /* IMPORTANT!! These have now changed under V41 */
    SN_SP_DEFAULT_POS = 0
    SN_SP_TOP_LEFT
    SN_SP_TOP_RIGHT
    SN_SP_BOTTOM_LEFT
    SN_SP_BOTTOM_RIGHT

End Enum

    ' /* these only apply if the new E_CLASS_CUSTOM_DURATION is not set */

Public Enum SN_NOTIFICATION_DURATION
    ' /* IMPORTANT!! These have now changed under V41 */
'    SN_ND_DEFAULT = 0
    SN_ND_APP_DECIDES = 1
    SN_ND_CUSTOM           ' // "custom_timeout" contains value in seconds

End Enum

    ' /* master controls */
Public g_IsRunning  As Boolean
Public g_IsQuitting As Boolean

    ' /* rosters */
Public g_ExtnRoster As TExtensionRoster
Public g_StyleRoster As TStyleRoster
Public g_AppRoster As TApplicationRoster
Public g_NotificationRoster As TNotificationRoster
Public g_SubsRoster As TNetworkRoster

'Public Enum E_FONTSMOOTHING
'    E_MELONTYPE
'    E_NONE
'    E_ANTIALIAS
'    E_CLEARTYPE
'    E_WINDOWS_DEFAULT
'
'End Enum

'Private Const SNARL_XXX_GLOBAL_MSG = "SnarlGlobalEvent"

Public Type T_CONFIG

'    run_on_logon As Boolean
'    font_smoothing As E_FONTSMOOTHING
'    suppress_delay As Long          ' // in ms
'    hotkey_prefs As Long            ' // MAKELONG(mods,key)
'    last_sound_folder As String
'    use_hotkey As Boolean
'    UserDnD As Boolean      ' // not persitent: user-controlled DND setting
'    SysDnDCount As Long             ' // set using WM_MANAGE_SNARL
'    MissedCountOnDnD As Long
'    use_dropshadow As Boolean
    last_update_check As Date
    AgreeBetaUsage As Boolean

    ' /* R2.31 */
    SnarlConfigPath As String           ' // path (UNC or other) to Snarl config folder (should contain /etc and other folders)
    SnarlConfigFile As String           ' // not persistent; just a handy copy

End Type

Public gPrefs As T_CONFIG

Dim mSettings As ConfigFile         ' // V40.25 - new way of managing persistent settings
Dim mConfig As ConfigSection        ' // V40.25 - the actual config section
Dim mDefaults As BPackedData        ' // V40.25 - new way of managing persistent settings
Dim mConfigLocked As Boolean
Dim mWriteConfigOnUnlock As Boolean

'Public g_Settings As ConfigFile
Dim m_Alerts As ConfigSection

'Public g_IgnoreLock As Long         ' // if >0 don't alert when app registers - overrides class setting
'Public gSelectedClass As TAlert
Public gDebugMode As Boolean
'Public mAwayCount As Long           ' // R2.4 DR8: renamed and reimplemented

Public gLastNotification As Date    ' // V41.47 - last notification timestamp

'Public Type G_REMOTE_COMPUTER
'    IsHostName As Boolean
'    HostNameOrIp As String
'
'End Type

Public ghWndMain As Long
Public gUpdateFilename As String        ' // name of the update file to download

Dim mUpdateCheck As TAutoUpdate
Dim mBetaUpdateCheck As TAutoUpdate

'Public gCurrentLowPriority As T_NOTIFICATION_INFO       ' // only one can be on-screen at any one time
Public gSnarlToken As Long              ' // when Snarl registers with itself
Public gSnarlPassword As String         ' // created on the fly

'Dim mDoNotDisturbLock As Long           ' // >0 means enabled, <=0 means disabled
'Public gNotificationMenuOpen As Boolean

    ' /* R2.4 DR8 */

Public Enum SN_PRESENCE_FLAGS
    ' /* Away flags occupy bottom 16 bits */
    SN_PF_AWAY_USER_IDLE = 1
    SN_PF_AWAY_COMPUTER_LOCKED = 2
    SN_PF_AWAY_SCREENSAVER_ACTIVE = 8
    SN_PF_AWAY_SOS = &H10
    SN_PF_AWAY_MASK = &HFFFF&

    ' /* DnD flags occupy top 16 bits */
    SN_PF_DND_FULLSCREEN_APP = &H10000
    SN_PF_DND_USER = &H20000                       ' // from the tray icon menu
    SN_PF_DND_EXTERNAL = &H40000                   ' // for future use
    SN_PF_DND_MASK = &HFFFF0000

End Enum

Dim mPresFlags As SN_PRESENCE_FLAGS

Public Enum SN_PRESENCE_ACTIONS
'    SN_PA_LAST_ERROR_SET = -1
    SN_PA_DO_DEFAULT = 0
    SN_PA_LOG_AS_MISSED = 1
    SN_PA_MAKE_STICKY = 2
    SN_PA_DO_NOTHING = 3
    SN_PA_DISPLAY_NORMAL = 4
    SN_PA_DISPLAY_URGENT = 5
'    SN_PA_FORWARD = 6

End Enum

    ' /* RunFile style stuff - shouldn't be global */
Public gRunFiles As BTagList    '// BTagItem->Name = filename, BTagItem->Value = version
                                '// version 1 = filename only, static template
                                '// version 2 = filename with variable template
                                '// version 3 = unabridged content in &/= format (usable by HeySnarl for example)

Public gStyleDefaults As CConfFile
Public gStartTime As Date

    ' /* new three-state option - easier to use with prefs kit which requires a 1-based value */

Public Enum SN_THREE_STATE
    SN_TS_ALWAYS = 1
    SN_TS_NEVER = 2
    SN_TS_APP_DECIDES = 3

End Enum

    ' /* R2.5 Beta 2 */
Dim mFlags As SNARL_SYSTEM_FLAGS
Public gCurrentExtension As TExtension

Public Enum SN_REDIRECTION_FLAGS
    SN_RF_WHEN_BUSY = 1
    SN_RF_WHEN_AWAY = 2
    SN_RF_WHEN_ACTIVE = 4
    SN_RF_ALWAYS = SN_RF_WHEN_AWAY Or SN_RF_WHEN_BUSY Or SN_RF_WHEN_ACTIVE
    SN_RF_NEVER = 0

End Enum

Public gGlobalRedirectList As BTagList

    ' /* requesters */
Dim mReq() As TRequester
Dim mReqs As Long
Public gRequestId As Long

Public Sub Main()
Dim l As Long

    ' /* get comctl.dll loaded up for our XP manifest... */

    g_InitComCtl

    ' /* is Snarl already running? */

    l = FindWindow(WINDOW_CLASS, "Snarl")

'    MsgBox Command$

Dim fQuitWhenDone As Boolean
Dim pArgs As BTagList
Dim sz As String
Dim i As Long

    Set pArgs = g_MakeArgList(Command$)

    If Command$ <> "" Then
        ' /* command specified, but which? */
        Debug.Print "command line: " & Command$
'        MsgBox Command$


        i = 1
        
        ' /* commands:
        '       -c | --configure
        '       -debug
        '       -do
        '       -l | --load
        '       -p | --parse
        '       -q | --quit | -quit
        '       -u | --unload
        '
        ' */

        Do While i <= pArgs.CountItems
        
'            MsgBox pArgs.TagAt(i).Name
        
            Select Case LCase$(pArgs.TagAt(i).Name)
            Case "-quit", "-q", "--quit"
                ' /* if we have an already running instance, tell it to quit */
                If l <> 0 Then _
                    SendMessage l, WM_CLOSE, 0, ByVal 0&
    
                ' /* don't process any further */
                Exit Sub

            Case "-debug"
                ' /* enable debug mode */
                gDebugMode = True

            Case "-p", "--parse"
                ' /* -r <request> - process standard request */
                i = i + 1
                If i <= pArgs.CountItems Then
                    snDoRequest pArgs.TagAt(i).Name
                    fQuitWhenDone = True

                Else
                    ' /* arg missing: stop processing */
                    MsgBox "Missing argument: use '-p|--parse <request>'", vbExclamation Or vbOKOnly, App.Title
                    Exit Sub

                End If

            Case "-do"
                ' /* -do <snarl:request> - URL handler */
                i = i + 1
                If i <= pArgs.CountItems Then
                    sz = pArgs.TagAt(i).Name
                    If g_BeginsWith(sz, "snarl:", False) Then
                        ' /* from the URL handler */
                        MsgBox "URL: " & sz
                        fQuitWhenDone = True
                        
                    Else
                        ' /* invalid */
                        MsgBox "Invalid argument: use '-do <command>'", vbExclamation Or vbOKOnly, App.Title
                        Exit Sub

                    End If
                Else
                    ' /* arg missing: stop processing */
                    MsgBox "Missing argument: use '-do <command>'", vbExclamation Or vbOKOnly, App.Title
                    Exit Sub

                End If

            Case "-u", "--unload"
                ' /* -u <thing> - unload extension or styleengine
                i = i + 1
                If i <= pArgs.CountItems Then
                    ' /* send the request */
                    sz = "snarl?cmd=unload&what=" & pArgs.TagAt(i).Name
                    snDoRequest sz
                    fQuitWhenDone = True

                Else
                    ' /* arg missing: stop processing */
                    MsgBox "Missing argument: use '--unload <AddOn>'", vbInformation Or vbOKOnly, App.Title
                    Exit Sub

                End If

            Case "-l", "--load"
                ' /* -l <thing> - load extension or styleengine
                i = i + 1
                If i <= pArgs.CountItems Then
                    ' /* send the request */
                    snDoRequest "snarl?cmd=load&what=" & pArgs.TagAt(i).Name
                    fQuitWhenDone = True

                Else
                    ' /* arg missing: stop processing */
                    MsgBox "Missing argument: use '--load <AddOn>'", vbInformation Or vbOKOnly, App.Title
                    Exit Sub

                End If

            Case "-c", "--configure"
                ' /* -c <thing> - configure extension, app or style
                i = i + 1
                If i <= pArgs.CountItems Then
                    ' /* send the request */
                    snDoRequest "snarl?cmd=configure&what=" & pArgs.TagAt(i).Name
                    fQuitWhenDone = True

                Else
                    ' /* arg missing: stop processing */
                    MsgBox "Missing argument: use '--configure <Extension|App|Style>'", vbInformation Or vbOKOnly, App.Title
                    Exit Sub

                End If



            Case Else

                ' /* check to see if a particular file was dropped */
                Select Case g_GetExtension(pArgs.TagAt(i).Name, True)
                Case "webforward"
                    fQuitWhenDone = True
'                    MsgBox "webforward"
'                    g_CopyToAppData Command$, "styles\webforward"
'                    If l Then _
'                        PostMessage l, WM_SNARL_COMMAND, SN_DP_RESTART_STYLE_ROSTER, ByVal 0&
'
'                    Exit Sub

                Case "rsz"
                    ' /* packed runnable style: if it installs ok, tell the running instance to restart the engine */
                    fQuitWhenDone = True
                    If g_InstallRSZ(pArgs.TagAt(i).Name, False) Then _
                        snDoRequest "snarl?cmd=reload&what=runnable.styleengine"

                Case "ssz"
                    ' /* packed scripted style: if it installs ok, tell the running instance to restart the engine */
                    fQuitWhenDone = True
                    If g_InstallSSZ(pArgs.TagAt(i).Name, False) Then _
                        snDoRequest "snarl?cmd=reload&what=script.styleengine"

                Case Else
                    MsgBox "unknown argument " & g_Quote(pArgs.TagAt(i).Name)
                    Exit Sub

                End Select
            End Select
            i = i + 1

        Loop

'        Case "-install"
'            ' /* must have one further arg: style engine or extension to install */
'            If (szArgs <> "") And (l <> 0) Then
'                PostMessage l, WM_SNARL_COMMAND, 1, ByVal RegisterClipboardFormat(szArgs)
'                Exit Sub
'
'            End If
'
'        Case "-configure"
'            ' /* must have one further arg: style or extension to install */
'            If (szArgs <> "") And (l <> 0) Then
'                PostMessage l, WM_SNARL_COMMAND, 2, ByVal RegisterClipboardFormat(szArgs)
'                Exit Sub
'
'            End If
'
'        Case "-do"
'            ' /* from the URL handler */
'            If g_BeginsWith(szArgs, "snarl:", False) Then
'                szArgs = g_SafeRightStr(szArgs, Len(szArgs) - 6)
'                i = InStr(szArgs, "?")
'                If i Then
'                    szCmd = g_SafeLeftStr(szArgs, i - 1)
'                    szArgs = g_SafeRightStr(szArgs, Len(szArgs) - i)
'
'                Else
'                    szCmd = szArgs
'                    szArgs = ""
'
'                End If
'
'                Select Case szCmd
'                Case "about"
'                    If l <> 0 Then _
'                        PostMessage l, WM_SNARL_COMMAND, SN_DP_SHOW_INFO, ByVal 0
'
'                Case Else
'                    MsgBox "Unknown command '" & szCmd & "' Args: '" & szArgs & "'", vbExclamation Or vbOKOnly, App.Title
'
'                End Select
'
'
'            Else
'                MsgBox "Incorrect URL format supplied", vbExclamation Or vbOKOnly, App.Title
'
'            End If
'
'            Exit Sub
'
'        Case Else
'


    End If

    If (l <> 0) And (NOTNULL(pArgs)) Then
        ' /* Snarl is already running (and no useful command-line arg specified) */
        If pArgs.CountItems = 0 Then _
            PostMessage l, WM_SNARL_COMMAND, 0, ByVal 0&    ' // tell the running instance to show its ui...

        Exit Sub

    ElseIf fQuitWhenDone Then
        Exit Sub

    End If

    App.TaskVisible = False

    ' /* V38.133 - enable debug mode if switch present */
    ' /* V40.18 - or if either CTRL key is held down */
    ' /* V42.122 - if beta release */
    ' /* V43.71 - if "D" pressed, not CTRL */

    gDebugMode = gDebugMode Or (g_IsPressed(vbKeyD)) Or (uIsDebugBuild())

    ' /* reset debug mode if admin setting blocks it */

    If gSysAdmin.NoDebugMode Then _
        gDebugMode = False


    If gDebugMode Then
        ' /* start logging */
        l3OpenLog "%APPDATA%\full phat\snarl\snarl.log", True
        g_Debug "** Snarl " & App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ") **"
        g_Debug "** " & App.LegalCopyright
        g_Debug ""

    End If

    ' /* need to check exec.library first */

    If Not melonGetExec() Then
        MsgBox "Snarl requires exec.library V46 or greater", vbCritical, "Snarl Initialisation Failed"
        GoTo noexec

    End If

    ' /* NOTE! We no longer need V46 graphics but older extensions will rely on it */

    If open_library("graphics.library", 46) <> M_OK Then
        MsgBox "Snarl requires graphics.library V46 or greater", vbCritical, "Snarl Initialisation Failed"
        GoTo nographics

    End If

    If open_library("libnitro1_2", 41) <> M_OK Then
        MsgBox "Snarl requires Nitro R1.2 V41 or greater", vbCritical, "Snarl Initialisation Failed"
        GoTo nonitro

    End If

    If open_library("openmenulite.library", 46) <> M_OK Then
        MsgBox "Snarl requires openmenulite.library V46 or greater", vbCritical, "Snarl Initialisation Failed"
        GoTo noopenmenu

    End If

    ' /* check storage kit */

    If Not melonCheckKit("storage") Then
        MsgBox "Snarl needs the Storage Kit!", vbCritical, "Snarl Initialisation Failed"
        GoTo nostoragekit

    End If

Dim pName As String

    ' /* resources */

    If Not melonCheckLibOrResource("icon_resource", 0, pName) Then
        MsgBox "icon.resource is damaged or not installed", vbCritical, "Snarl Early Startup Error"
        GoTo nostoragekit

    Else
        g_Debug "Main(): got " & pName

    End If

    If Not melonCheckLibOrResource("web_resource", 0, pName) Then
        MsgBox "web.resource is damaged or not installed", vbCritical, "Snarl Early Startup Error"
        GoTo nostoragekit

    Else
        g_Debug "Main(): got " & pName

    End If

    If Not melonCheckLibOrResource("misc_resource", 0, pName) Then
        MsgBox "misc.resource is damaged or not installed", vbCritical, "Snarl Early Startup Error"
        GoTo nostoragekit

    Else
        g_Debug "Main(): got " & pName

    End If



    ' /* -------------- end of early startup ------------- */

    ' /* the first thing we should do now is create the message handling window */

    If Not EZRegisterClass(WINDOW_CLASS) Then
        g_Debug "main(): couldn't register window class", LEMON_LEVEL_CRITICAL
        GoTo nostoragekit

    End If

    ghWndMain = EZ4AddWindow(WINDOW_CLASS, New TMainWindow, "Snarl", 0, 0)
    If ghWndMain = 0 Then
        g_Debug "main(): couldn't create window", LEMON_LEVEL_CRITICAL
        EZUnregisterClass WINDOW_CLASS
        GoTo nostoragekit

    End If

    ' /* V41: set our system version and revision as properties */

    SetProp ghWndMain, "_version", App.Major
    SetProp ghWndMain, "_revision", App.Revision

    ' /* R2.31: set our flags */

Dim dwFlags As SNARL_SYSTEM_FLAGS

    If gDebugMode Then _
        dwFlags = dwFlags Or SNARL_SF_DEBUG_MODE

    g_SetSystemFlags dwFlags

    ' /* R2.31: init last notification timestamp */
    
    gLastNotification = Now()

    ' /* intialize the IP forwarding subsystem */

'    g_ForwardInit

    ' /* R2.31 - pre-set our config folder */

'    gPrefs.SnarlConfigPath = App.Path                ' // fail-safe

'Dim sz As String

    ' /* R2.31 - look for a local sysconfig.ssl and get its target  */

Dim pSysConfig As CConfFile

    Set pSysConfig = New CConfFile
    With pSysConfig
        If .SetTo(g_MakePath(App.Path) & "sysconfig.ssl", True) Then
            g_Debug "Main: local sysconfig.ssl exists, querying..."
            sz = g_RemoveQuotes(.ValueOf("target"))

            ' /* R2.31 - quick check on folder structure */

            If g_IsFolder(g_MakePath(sz) & "etc") Then
                gPrefs.SnarlConfigPath = sz

            Else
                g_Debug "Main: config path '" & sz & "' is invalid", LEMON_LEVEL_CRITICAL

            End If
        End If

    End With



    If gPrefs.SnarlConfigPath = "" Then
        g_Debug "Main(): no pre-defined configuration path"

        If g_GetUserFolderPath(sz) Then
            gPrefs.SnarlConfigPath = sz                  ' // standard location
            uCreateUserSettings

        Else
            g_Debug "Main(): problem setting user-specific config path", LEMON_LEVEL_CRITICAL
            gPrefs.SnarlConfigPath = App.Path
    
        End If

    End If




    gPrefs.SnarlConfigPath = g_MakePath(gPrefs.SnarlConfigPath)
    g_Debug "Main: config path is '" & gPrefs.SnarlConfigPath & "'"

    ' /* fix up the .snarl path */

    gPrefs.SnarlConfigFile = gPrefs.SnarlConfigPath & "etc\config41.snarl"
    g_Debug "Main: .snarl path is '" & gPrefs.SnarlConfigFile & "'"

    ' /* R2.31: register our config path as a global atom and store the atom as a window property */

    SetProp ghWndMain, "_config_path", RegisterWindowMessage(gPrefs.SnarlConfigPath)

    ' /* do we have a snarl.admin file to load? */

Dim szName As String
Dim szData As String

    Set pSysConfig = New CConfFile
    If pSysConfig.SetTo(gPrefs.SnarlConfigPath & "etc\snarl.admin", True) Then

        g_Debug "Main: loaded admin settings from '" & gPrefs.SnarlConfigPath & "snarl.admin" & "'"

        With pSysConfig
            .Rewind
            Do While .GetEntry(szName, szData)
                g_Debug "Main: '" & szName & "'='" & szData & "'"

            Loop

        End With

        With gSysAdmin
            .HideIcon = (pSysConfig.ValueOf("HideIcon") = "1")
            .InhibitPrefs = (pSysConfig.ValueOf("InhibitPrefs") = "1")
            .TreatSettingsAsReadOnly = (pSysConfig.ValueOf("TreatSettingsAsReadOnly") = "1")
            .InhibitMenu = (pSysConfig.ValueOf("InhibitMenu") = "1")
            .InhibitQuit = (pSysConfig.ValueOf("InhibitQuit") = "1")
            ' /* R2.5.1 */
            .BlockUnknownApps = (pSysConfig.ValueOf("BlockUnknownApps") = "1")
            .NoRunStartupSequence = (pSysConfig.ValueOf("NoRunStartupSequence") = "1")
            .NoDebugMode = (pSysConfig.ValueOf("NoDebugMode") = "1")

        End With

    Else
        g_Debug "Main: no admin settings file"

    End If

    ' /* get settings */

    If Not g_ConfigInit() Then
        g_Debug "main(): new/clean installation..."
        g_ConfigSet "step_size", "1"

    End If

    ' /* load up some required bits and bobs */

    g_LoadIconTheme


'    ' /* R2.4: managed style settings */
'
'    Set gStyleSettings = New CConfFile
'    gStyleSettings.SetTo gPrefs.SnarlConfigPath & "etc\.stylesettings"

    ' /* main start */

    SendMessage ghWndMain, WM_SNARL_INIT, 0, ByVal 0&

    Load frmAbout           ' // keeps us open...

    ' /* R2.4.2 DR3: start network (renamed in 2.6 from subscriber) roster */

    g_Debug "Main(): Starting network roster..."
    Set g_SubsRoster = New TNetworkRoster
    melonLibInit g_SubsRoster
    melonLibOpen g_SubsRoster

    ' /* start notification roster */

    g_Debug "Main(): Starting notifications roster..."
    Set g_NotificationRoster = New TNotificationRoster
    melonLibInit g_NotificationRoster
    melonLibOpen g_NotificationRoster

    ' /* start app roster */

    g_Debug "Main(): Starting app roster..."
    Set g_AppRoster = New TApplicationRoster
    melonLibInit g_AppRoster
    melonLibOpen g_AppRoster

'    g_Debug "Main(): Setting auto-run state..."
'    g_SetAutoRun2

    ' /* get style packs */

    g_Debug "Main(): Starting style roster..."
    Set g_StyleRoster = New TStyleRoster
    melonLibInit g_StyleRoster
    melonLibOpen g_StyleRoster

    ' /* get icon themes */

'    g_GetIconThemes

    ' /* set master running flag */

    g_SetRunning True, False
    gStartTime = Now()

    ' /* display welcome message */

    If (g_ConfigGet("show_msg_on_start") = "1") Or (gDebugMode) Then
        i = g_PrivateNotify(SNARL_CLASS_GENERAL, "Welcome to Snarl!", _
                            "Snarl " & g_Version() & vbCrLf & App.LegalCopyright & vbCrLf & "http://www.getsnarl.info" & IIf(gDebugMode, vbCrLf & vbCrLf & "Debug mode enabled", "") & IIf(g_IsAlphaBuild, vbCrLf & "Alpha build", ""), , _
                            g_MakePath(App.Path) & "etc\icons\snarl.png")

        If i Then
            g_QuickAddAction i, "User Guide", "http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=User_Guide"
            g_QuickAddAction i, "Release Notes", "http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=New_Features#" & Replace$(App.Comments, " ", "_")
        
        End If

    End If

    ' /* get extensions */

    g_Debug "Main(): Starting extensions roster..."
    Set g_ExtnRoster = New TExtensionRoster
    melonLibInit g_ExtnRoster
    melonLibOpen g_ExtnRoster


    Set mUpdateCheck = New TAutoUpdate
    If g_ConfigGet("auto_update") = "1" Then
        g_Debug "Main(): Doing auto-update check..."
        If mUpdateCheck.Check(False, "http://www.fullphat.net/updateinfo/snarl.updateinfo") Then _
            g_Debug "Main(): auto-update check initiated"

    Else
        g_Debug "Main(): Auto-update is disabled"

    End If

    ' /* R2.4.1: check for beta release as well */

    Set mBetaUpdateCheck = New TAutoUpdate

    If g_ConfigGet("auto_beta_update") = "1" Then
        g_Debug "Main(): Doing auto-update (beta) check..."
        If mBetaUpdateCheck.Check(False, "http://snarlwin.svn.sourceforge.net/viewvc/snarlwin/snarl-beta.updateinfo") Then _
            g_Debug "Main(): beta auto-update check initiated"

    Else
        g_Debug "Main(): Auto-update is disabled"

    End If


    ' /* notify ready to run */

    frmAbout.bReadyToRun

    If Not g_IsPressed(vbKeyS) Then
        ' /* launch startup-sequence apps */
        g_Debug "Main(): processing startup-sequence...", LEMON_LEVEL_INFO

        If g_GetUserFolderPath(sz, True) Then _
            uRunStartupSequence sz

        If g_GetUserFolderPath(sz) Then _
            uRunStartupSequence sz

    Else
        g_Debug "Main(): skipping startup-sequence (key held)", LEMON_LEVEL_INFO

    End If

'    g_Debug "Main(): garbage collection"
'    If g_IsWinXPOrBetter() Then _
'        CoFreeUnusedLibrariesEx 0, 0

    g_Debug "Main(): startup complete"

    With New BMsgLooper
        .Run

    End With

    g_Debug "main(): stopping..."

    If mReqs Then
        For l = mReqs To 1 Step -1
            mReq(l).Quit

        Next l

    End If

    Set mUpdateCheck = Nothing         ' // this will abort the request if it's still running...

Dim uSnarlGlobal As Long
Dim t As Long

    t = GetTickCount()
    g_IsQuitting = True

    ' /* broadcast SNARL_QUIT */

    g_Debug "main(): broadcasting SNARL_QUIT..."
    PostMessage HWND_BROADCAST, snSysMsg(), SNARL_BROADCAST_QUIT, ByVal CLng(App.Major)

    SendMessage ghWndMain, WM_SNARL_QUIT, 0, ByVal 0&
    Unload frmAbout

    ' /* stop various rosters - order *is* important (should be reverse of startup) */

    g_Debug "main(): stopping extension roster..."
    melonLibClose g_ExtnRoster
    melonLibUninit g_ExtnRoster

    g_Debug "main(): stopping style roster..."
    melonLibClose g_StyleRoster
    melonLibUninit g_StyleRoster

    g_Debug "main(): stopping application roster..."
    melonLibClose g_AppRoster
    melonLibUninit g_AppRoster

    g_Debug "main(): stopping notification roster..."
    melonLibClose g_NotificationRoster
    melonLibUninit g_NotificationRoster

    g_Debug "main(): stopping subscriptions roster..."
    melonLibClose g_SubsRoster
    melonLibUninit g_SubsRoster

    EZ4RemoveWindow ghWndMain
    EZUnregisterClass WINDOW_CLASS

    SOS_quit

    ' /* done */

    t = GetTickCount() - t
    g_Debug "main(): took " & t & " ms to complete closedown"

nostoragekit:
    close_library "openmenulite.library"

noopenmenu:
    close_library "Nitro R1.2"

nonitro:
    close_library "graphics.library"

nographics:
noexec:

End Sub

Public Function g_ConfigInit() As Boolean

    On Error Resume Next

    ' /* defaults */

    Set mDefaults = New BPackedData
    With mDefaults
        .Add "default_position", CStr(SN_SP_BOTTOM_RIGHT)
        .Add "show_msg_on_start", "1"
        .Add "run_on_logon", "1"

        ' /* R2.0 (V38.13) */
        .Add "default_style", "corporate/standard"    ' // as "<style>[/<scheme>]

        ' /* R2.0 (V38.32) */
        .Add "sticky_snarls", "0"
        .Add "log_only", "0"
        .Add "default_duration", "10"

        ' /* R2.04 (V38.82) - no longer used */
'        .Add "font_smoothing", CStr(E_MELONTYPE)
        .Add "melontype_contrast", "10"

        ' /* R2.1 (V39) */
        .Add "listen_for_json", "1"             ' // R2.5 Beta 2 now enabled by default
        .Add "listen_for_snarl", "1"            ' // R2.5 Beta 2 now enabled by default
        .Add "duplicates_quantum", "2000"
        .Add "hotkey_prefs", CStr(vbKeyF10)
        .Add "notify_on_first_register", "0"
        .Add "global_opacity", "100"
        .Add "last_sound_folder", g_GetSystemFolderStr(CSIDL_PERSONAL)
        .Add "show_tray_icon", "1"
        .Add "ignore_new_classes", "0"      ' // new alert classes are always enabled by default

        ' /* R2.2 */
        .Add "use_hotkey", "1"
        .Add "do_not_disturb", "0"
'        .Add "idle_timeout", "300"          ' // i.e. 5 minutes
        .Add "margin_spacing", "0"
        .Add "use_dropshadow", "1"
        .Add "dropshadow_strength", "88"    ' // is a %
        .Add "dropshadow_size", "10"
        .Add "icon_theme", ""

        ' /* R2.3 */
        .Add "auto_update", "1"
        .Add "enable_sounds", "1"
        .Add "use_style_sounds", "1"
        .Add "prefer_style_sounds", "0"
        .Add "default_normal_sound", ""
        .Add "default_priority_sound", ""
        .Add "use_style_icons", "1"
        .Add "auto_sticky_on_screensaver", "1"
        .Add "show_timestamp", "0"

        ' /* R2.4: style-usable settings are prefixed with 'style.' */
        
        .Add "style.overflow_limit", "7"

        ' /* R2.4 DR8 */

        .Add "away_when_locked", "1"
        .Add "away_when_fullscreen", "1"
        .Add "away_when_screensaver", "1"
        .Add "away_mode", "1"               ' // log missed
        .Add "busy_mode", "1"               ' // log missed

        ' /* R2.4 Beta 4 */

        .Add "idle_minutes", "4"            ' // i.e. 5 minutes
        .Add "include_host_name_when_forwarding", "0"

        ' /* R2.4.1 */

        .Add "allow_right_clicks", "0"
        .Add "auto_beta_update", "0"

        ' /* R2.4.2 */

        .Add "only_allow_secure_apps", "0"
        .Add "apps_must_register", "0"
        .Add "auto_detect_url", "0"
        .Add "no_callback_urls", "0"
        .Add "block_null_pid", "0"              ' // blocks WM_COPYDATA where wParam is NULL
        .Add "block_dos_attempt", "1"
        .Add "ignore_style_requests", "0"       ' // prevents notifications from requesting specific styles
        .Add "global_shadow_list", ""
'        .Add "auth_type", "md5"
'        .Add "auth_salt", ""
'        .Add "auth_key", ""
        .Add "auth_password", ""
        .Add "default_screen", "1"

        ' /* R2.5 Beta 2 */
        .Add "allow_subs", "1"
        .Add "include_icon_when_forwarding", "1"

        ' /* R2.5.1 */
        .Add "use_notification_hotkey", "1"

        ' /* R2.6 *
        .Add "scaling", "1"
        .Add "global_redirect", ""
        .Add "garbage_collection", "1"
        .Add "show_missed_notifications", "2"
        .Add "notify_when_subscriber_added", "1"
        .Add "nc-col-background", CStr(rgba(31, 33, 33))
        .Add "nc-col-text", CStr(rgba(255, 255, 255))
        .Add "nc-font-typeface", "Tahoma"
        .Add "nc-font-point", "9"
        .Add "nc-opacity-percent", "90"
        .Add "callback_as_button", "1"
        .Add "block_net_control", "0"

    End With

    ' /* attempt to load the config file */

'    MsgBox gSysAdmin.RemoteSnarlFile

Dim i As Long

    Set mSettings = New ConfigFile
    With mSettings
        .File = gPrefs.SnarlConfigFile
        .Load

        i = .FindSection("general")
        If i = 0 Then
            Set mConfig = .AddSectionObj("general")
            If Not gSysAdmin.TreatSettingsAsReadOnly Then _
                .Save

        Else
            Set mConfig = .SectionAt(i)

        End If

'        i = .FindSection("remote_computers")
'        If i = 0 Then
'            Set gRemoteComputers = .AddSectionObj("remote_computers")
'            If Not gSysAdmin.TreatSettingsAsReadOnly Then _
'                .Save
'
'        Else
'            Set gRemoteComputers = .SectionAt(i)
'
'        End If

    End With

    ' /* R2.4.2 DR3: style defaults */

    Set gStyleDefaults = New CConfFile

    With gStyleDefaults
        .SetTo g_MakePath(gPrefs.SnarlConfigPath) & "etc\styledefaults.conf"

        .AddIfMissing "background-colour", CStr(rgba(235, 235, 235))
        .AddIfMissing "background-colour-priority", CStr(rgba(235, 99, 5))
        .AddIfMissing "width", "250"

        .AddIfMissing "title-font", "name::Tahoma#?size::8#?bold::1#?italic::0#?underline::0#?strikeout::0"
        .AddIfMissing "title-colour", CStr(rgba(0, 0, 0))
        .AddIfMissing "title-colour-priority", CStr(rgba(0, 0, 0))
'        .AddIfMissing "title-weight", "1"
        .AddIfMissing "title-opacity", "70"

        .AddIfMissing "text-font", "name::Tahoma#?size::8#?bold::0#?italic::0#?underline::0#?strikeout::0"
        .AddIfMissing "text-colour", CStr(rgba(0, 0, 0))
        .AddIfMissing "text-colour-priority", CStr(rgba(0, 0, 0))
'        .AddIfMissing "text-weight", "1"
        .AddIfMissing "text-opacity", "60"

        ' /* R2.6 */
        .AddIfMissing "colour-tint", CStr(rgba(0, 0, 0))

    End With

    ' /* load up the global redirects list */

    Set gGlobalRedirectList = new_BTagList()

Dim sn As String
Dim sv As String

    With New BPackedData
        If .SetTo(g_ConfigGet("global_redirect")) Then
            .Rewind
            Do While .GetNextItem(sn, sv)
                gGlobalRedirectList.Add new_BTagItem(sn, sv)

            Loop

        End If
    End With








    g_ConfigInit = (Val(g_ConfigGet("step_size")) > 0)

End Function

Public Function g_ConfigGet(ByVal Name As String) As String

    ' /* pre-set with default */

    If Not (mDefaults Is Nothing) Then _
        g_ConfigGet = mDefaults.ValueOf(Name)

Dim sz As String

    If Not (mConfig Is Nothing) Then
        If mConfig.Find(Name, sz) Then _
            g_ConfigGet = sz

    End If

End Function

Public Sub g_ConfigSet(ByVal Name As String, ByVal Value As String)

    If (mConfig Is Nothing) Then _
        Exit Sub

    mConfig.Update Name, Value
    g_WriteConfig

End Sub

Public Sub g_WriteConfig()

'    Debug.Print "g_WriteConfig: " & gNoWriteConfig

    If mConfigLocked Then
        g_Debug "g_WriteConfig(): config is locked - request queued"
        mWriteConfigOnUnlock = True
        Exit Sub

    End If

    ' /* R2.31: only if admin says we can... */

    If gSysAdmin.TreatSettingsAsReadOnly Then
        g_Debug "g_WriteConfig(): configuration as been set as read only by system administrator", LEMON_LEVEL_WARNING

    Else
        g_Debug "g_WriteConfig(): writing to " & mSettings.File & "..."
        mSettings.Save

    End If

End Sub

Public Function g_Version() As String

    '& IIf(App.Revision <> 0, "." & App.Revision, "")
'    g_Version = App.Major & "." & App.Minor & IIf(App.Comments <> "", " " & App.Comments, "") & " (Build " & CStr(App.Revision) & ")"

    g_Version = App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ")"

End Function

Public Sub g_SetRunning(ByVal IsRunning As Boolean, Optional ByVal Broadcast As Boolean = True)

    If g_IsRunning = IsRunning Then _
        Exit Sub

Dim dw As Long

    If IsRunning Then
        ' /* set master flag *first* */
        g_IsRunning = True

        ' /* tell the extensions */
        If Not (g_ExtnRoster Is Nothing) Then _
            g_ExtnRoster.SendSnarlState True

        ' /* tell our applications we're starting */
        If Not (g_AppRoster Is Nothing) Then _
            g_AppRoster.SendToAll SNARL_BROADCAST_LAUNCHED

        SetWindowText ghWndMain, "Snarl"

        ' /* R2.4: broadcast a started message */

        PostMessage HWND_BROADCAST, snSysMsg(), SNARL_BROADCAST_STARTED, ByVal 0&


'        If Broadcast Then
'            ' /* send started broadcast */
'            g_Debug "g_SetRunning(): Broadcasting SNARL_LAUNCHED..."
'            SendMessageTimeout HWND_BROADCAST, snGetGlobalMsg(), SNARL_LAUNCHED, ByVal CLng(App.Major), SMTO_ABORTIFHUNG, 500, dw
'
'        End If

    Else

'        If Broadcast Then
'            ' /* send stopped broadcast */
'            g_Debug "g_SetRunning(): Broadcasting SNARL_QUIT..."
'            SendMessageTimeout HWND_BROADCAST, snGetGlobalMsg(), SNARL_QUIT, ByVal CLng(App.Major), SMTO_ABORTIFHUNG, 500, dw
'
'        End If

        ' /* set master flag */
        g_IsRunning = False
        SetWindowText ghWndMain, "Snarl-stopped"

        ' /* close all notifications */
        If Not (g_NotificationRoster Is Nothing) Then _
            g_NotificationRoster.CloseMultiple 0

        ' /* tell the extensions */
        If Not (g_ExtnRoster Is Nothing) Then _
            g_ExtnRoster.SendSnarlState False

        ' /* tell our applications we've stopped */
        If Not (g_AppRoster Is Nothing) Then _
            g_AppRoster.SendToAll SNARL_BROADCAST_QUIT

        ' /* R2.4: broadcast a stopped message */
        PostMessage HWND_BROADCAST, snSysMsg(), SNARL_BROADCAST_STOPPED, ByVal 0&

    End If

    ' /* update tray icon */

    frmAbout.bSetTrayIcon

End Sub

'Public Sub g_SetAutoRun2()
'Dim bAutoRun As Boolean
'
'    bAutoRun = CBool(g_ConfigGet("run_on_logon"))
'
'    If bAutoRun Then
'        add_registry_startup_item "Snarl", g_MakePath(App.Path) & LCase$(App.EXEName) & ".exe"
'
'    Else
'        rem_registry_startup_item "Snarl", g_MakePath(App.Path) & LCase$(App.EXEName) & ".exe"
'
'    End If
'
'End Sub

Public Function gfRegisterAlert(ByVal AppName As String, ByVal Class As String, ByVal Flags As Long) As M_RESULT
Dim pa As TApp

    g_Debug "gfRegisterAlert('" & AppName & "' '" & Class & "' #" & g_HexStr(Flags) & ")", LEMON_LEVEL_PROC

    ' /* find the app */

    If Not g_AppRoster.Find(AppName, pa) Then
        g_Debug "gfRegisterAlert(): App not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfRegisterAlert = M_NOT_FOUND
        Exit Function

    End If

    ' /* try to add it */

    gfRegisterAlert = pa.AddAlert(Class, "")

End Function

Public Function gfAddClass(ByVal Pid As Long, ByVal Class As String, ByVal Flags As Long, ByVal Description As String) As M_RESULT
Dim pa As TApp

    g_Debug "gfAddClass('" & CStr(Pid) & "' '" & Class & "' #" & g_HexStr(Flags) & ")", LEMON_LEVEL_PROC

    ' /* find the app */

    If Not g_AppRoster.FindByPid(Pid, pa) Then
        g_Debug "gfAddClass(): App not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfAddClass = M_NOT_FOUND
        Exit Function

    End If

    ' /* try to add it */

    gfAddClass = pa.AddAlert(Class, Description)

End Function

'Public Function globalEnableAlert(ByVal AppName As String, ByVal AlertName As String, ByVal Enabled As Boolean) As Boolean
'Dim i As Long
'Dim j As Long
'
'    i = globalFindAppByName(AppName)
'    If i Then
'        With g_Applet(i).Alerts
'            j = .IndexOf(AlertName)
'            If j Then _
'                .Update AlertName, IIf(Enabled, "1", "0")
'
'        End With
'
'    End If
'
'End Function

Public Function g_UTF8(ByVal str As String) As String

    g_UTF8 = trim(toUnicodeUTF8(g_utoa(str)))

End Function

Public Function g_GetUserFolderPath(ByRef Path As String, Optional ByVal AllUsers As Boolean = False) As Boolean
Dim sz As String

    If AllUsers Then
        If Not g_GetSystemFolder(CSIDL_COMMONAPPDATA, sz) Then _
            Exit Function

    Else
        If Not g_GetSystemFolder(CSIDL_APPDATA, sz) Then _
            Exit Function

    End If

    Path = g_MakePath(sz) & "full phat\snarl"
    g_GetUserFolderPath = True

End Function

Public Function g_GetUserFolderPathStr(Optional ByVal AllUsers As Boolean = False) As String
Dim sz As String

    If g_GetUserFolderPath(sz, AllUsers) Then _
        g_GetUserFolderPathStr = g_MakePath(sz)

End Function

Public Function g_GetUserFolder(ByRef Folder As storage_kit.Node, Optional ByVal AllUsers As Boolean = False, Optional ByVal PathToAdd As String) As Boolean
Dim sz As String

    If Not g_GetSystemFolder(IIf(AllUsers, CSIDL_COMMONAPPDATA, CSIDL_APPDATA), sz) Then _
        Exit Function

    Set Folder = New Node
    g_GetUserFolder = Folder.SetTo(g_MakePath(sz) & "full phat\snarl" & IIf(PathToAdd <> "", "\" & PathToAdd, ""))

End Function

'Public Function g_GetUserStylesFolder(ByRef Folder As storage_kit.Node, Optional ByVal AllUsers As Boolean = False) As Boolean
'Dim sz As String
'
'    If Not g_GetSystemFolder(IIf(AllUsers, CSIDL_COMMONAPPDATA, CSIDL_APPDATA), sz) Then _
'        Exit Function
'
'    Set Folder = New Node
'    g_GetUserStylesFolder = Folder.SetTo(g_MakePath(sz) & "full phat\snarl\styles")
'
'End Function

Public Function g_GetSystemFolderNode(ByVal Path As CSIDL_VALUES, ByRef Folder As storage_kit.Node) As Boolean
Dim sz As String

    If Not g_GetSystemFolder(Path, sz) Then _
        Exit Function

    Set Folder = New storage_kit.Node
    g_GetSystemFolderNode = Folder.SetTo(g_MakePath(sz))

End Function

Public Function g_GetAppFolderNode(ByRef Folder As storage_kit.Node, Optional ByVal PathToAdd As String) As Boolean

    Set Folder = New storage_kit.Node
    g_GetAppFolderNode = Folder.SetTo(g_MakePath(App.Path) & PathToAdd)

End Function

'Public Function gSetUpFontSmoothing(ByRef aView As mfxView, ByVal TextColour As Long, ByVal SmoothingColour As Long) As MFX_DRAWSTRING_FLAGS
'Dim dw As Long
'
'    aView.SetHighColour TextColour
'
'    Select Case gPrefs.font_smoothing
'    Case E_MELONTYPE
'        If SmoothingColour = 0 Then
'            ' /* calculate it (only really works for dark colours at present) */
'            aView.SetLowColour rgba(get_red(TextColour), _
'                                    get_green(TextColour), _
'                                    get_blue(TextColour), _
'                                    (Val(g_ConfigGet("melontype_contrast")) / 100) * 255)
'
'        Else
'            aView.SetLowColour SmoothingColour
'
'        End If
'
'        gSetUpFontSmoothing = MFX_SIMPLE_OUTLINE
'
'    Case E_NONE
'        aView.TextMode = MFX_TEXT_PLAIN
'
'    Case E_ANTIALIAS
'        aView.TextMode = MFX_TEXT_ANTIALIAS
'
'    Case E_CLEARTYPE
'        aView.TextMode = MFX_TEXT_CLEARTYPE
'
'    Case E_WINDOWS_DEFAULT
'        SystemParametersInfo SPI_GETFONTSMOOTHING, 0, dw, 0
'        If dw = 0 Then
'            ' /* none */
'            aView.TextMode = MFX_TEXT_PLAIN
'
'        Else
'            ' /* enabled - but which type? */
'            aView.TextMode = MFX_TEXT_ANTIALIAS     ' // assume antialias...
'            If g_IsWinXPOrBetter() Then
'                dw = 0
'                SystemParametersInfo SPI_GETFONTSMOOTHINGTYPE, 0, dw, 0
'
'                If dw = FE_FONTSMOOTHINGCLEARTYPE Then _
'                    aView.TextMode = MFX_TEXT_CLEARTYPE
'
''                FE_FONTSMOOTHINGSTANDARD and
'
'            End If
'
'        End If
'
'    End Select
'
'End Function

Public Function gfSetAlertDefault(ByVal Pid As Long, ByVal Class As String, ByVal Element As Long, ByVal Value As String) As M_RESULT
Dim pa As TApp
Dim pc As TAlert

    g_Debug "gfSetAlertDefault('" & Pid & "' '" & Class & "' #" & CStr(Element) & " '" & Value & "')", LEMON_LEVEL_PROC

    If (g_AppRoster Is Nothing) Then
        g_Debug "gfSetAlertDefault(): App not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_ABORTED
        Exit Function

    End If

    ' /* find the app */

    If Not g_AppRoster.FindByPid(Pid, pa) Then
        g_Debug "gfSetAlertDefault(): App '" & Pid & "' not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_NOT_FOUND
        Exit Function

    End If

    ' /* check the class  */

    If Not pa.FindAlert(Class, pc) Then
        g_Debug "gfSetAlertDefault(): alert class '" & Class & "' not found", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_NOT_FOUND
        Exit Function

    End If

    ' /* change the value */

    gfSetAlertDefault = M_OK

    Select Case Element

    Case SNARL_ATTRIBUTE_TITLE
        pc.AppProvidedSettings.Update "title", Value, True
'        pc.DefaultTitle = Value

    Case SNARL_ATTRIBUTE_TEXT
        pc.AppProvidedSettings.Update "text", Value, True
'        pc.DefaultText = Value

    Case SNARL_ATTRIBUTE_TIMEOUT
        pc.AppProvidedSettings.Update "duration", CStr(Value), True
'        pc.DefaultTimeout = Val(Value)

    Case SNARL_ATTRIBUTE_SOUND
        pc.AppProvidedSettings.Update "sound", Value, True
'        pc.DefaultSound = Value

    Case SNARL_ATTRIBUTE_ICON
        pc.DefaultIcon = Value

    Case SNARL_ATTRIBUTE_ACK
        pc.AppProvidedSettings.Update "callback", Value, True
'        pc.DefaultAck = Value

    Case Else
        g_Debug "gfSetAlertDefault(): unknown element '" & Element & "'", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_INVALID_ARGS

    End Select

End Function

Public Sub g_WriteToLog(ByVal Title As String, ByVal Text As String)
Dim sz As String
Dim n As Integer

    On Error Resume Next

    If Not g_GetUserFolderPath(sz) Then _
        Exit Sub

    n = FreeFile()
    Open g_MakePath(sz) & "snarl_log.txt" For Append As #n
    If err.Number <> 0 Then _
        Exit Sub

    Print #n, CStr(Now()) & vbTab & Replace$(Title, vbCrLf, "/n") & vbTab & Replace$(Text, vbCrLf, "/n")
    Close #n

End Sub

'Public Sub g_GetIconThemes()
'Dim pn As storage_kit.Node
'
'    ReDim gIconTheme(0)
'    gIconThemes = 0
'
'    If g_GetUserFolder(pn) Then _
'        uGetIconThemes pn
'
'    If g_GetUserFolder(pn, True) Then _
'        uGetIconThemes pn
'
'End Sub
'
'Private Sub uGetIconThemes(ByRef Folder As storage_kit.Node)
'
'    If Not (Folder.SetTo(g_MakePath(Folder.File) & "themes")) Then _
'        Exit Sub
'
'    If Not (Folder.IsFolder) Then _
'        Exit Sub
'
'Dim i As Long
'Dim c As Long
'
'    With Folder
'        .ReadContents
'        c = .CountNodes
'        If c Then
'            For i = 1 To c
'                If .NodeAt(i).IsFolder Then _
'                    uGetIconTheme .NodeAt(i)
'
'            Next i
'        End If
'    End With
'
'End Sub
'
'Private Sub uGetIconTheme(ByRef Folder As storage_kit.Node)
'Dim pn As storage_kit.Node
'
'    Set pn = New storage_kit.Node
'    If Not (pn.SetTo(g_MakePath(Folder.File) & "icons")) Then _
'        Exit Sub
'
'    If Not (pn.IsFolder) Then _
'        Exit Sub
'
'Static i As Long
'Static j As Long
'
'    ' /* add it alpha-sorted */
'
'    If gIconThemes Then
'        For i = 1 To gIconThemes
'            If LCase$(Folder.Filename) < LCase$(gIconTheme(i).Name) Then
'                ' /* make a gap */
'                gIconThemes = gIconThemes + 1
'                ReDim Preserve gIconTheme(gIconThemes)
'                For j = gIconThemes To (i + 1) Step -1
'                    LSet gIconTheme(j) = gIconTheme(j - 1)
'
'                Next j
'
'                ' /* insert here */
'                With gIconTheme(i)
'                    .Name = Folder.Filename
'                    .Path = pn.File
'                    .IconFile = g_MakePath(Folder.File) & "theme.png"
'
'                End With
'                Exit Sub
'            End If
'        Next i
'    End If
'
'    ' /* drop through here if no other themes */
'
'    gIconThemes = gIconThemes + 1
'    ReDim Preserve gIconTheme(gIconThemes)
'    With gIconTheme(gIconThemes)
'        .Name = Folder.Filename
'        .Path = pn.File
'        .IconFile = g_MakePath(Folder.File) & "theme.png"
'
'    End With
'
'End Sub
'
'Public Function g_GetIconThemePath(ByVal Name As String, ByRef Path As String) As Boolean
'Dim i As Long
'
'    If gIconThemes Then
'        For i = 1 To gIconThemes
'            If LCase$(gIconTheme(i).Name) = LCase$(Name) Then
'                Path = g_MakePath(gIconTheme(i).Path)
'                g_GetIconThemePath = True
'
'            End If
'        Next i
'    End If
'
'End Function

Public Function g_IsValidImage(ByRef Image As MImage) As Boolean

    If (Image Is Nothing) Then _
        Exit Function

    g_IsValidImage = ((Image.Width > 0) And (Image.Height > 0))

End Function

Public Function g_DoSchemePreview2(ByVal Name As String, ByVal Scheme As String, ByVal IsPriority As Boolean, ByVal Percent As Integer) As M_RESULT

    ' /* this handles external requests to Snarl to display a notification in a particular
    '    style and scheme - only the SNARL_PREVIEW_SCHEME message handler calls this */

    If (g_NotificationRoster Is Nothing) Or (g_StyleRoster Is Nothing) Then _
        Exit Function

Dim pStyle As TStyle

    ' /* find the style */

    If Not g_StyleRoster.Find(Name, pStyle) Then _
        Exit Function

    If Scheme = "" Then
        ' /* if no scheme, use "<Default>" */
        Scheme = "<Default>"

    Else
        ' /* otherwise, supplied scheme must exist */
        If pStyle.SchemeIndex(Scheme) = 0 Then _
            Exit Function

    End If
    
Dim pInfo As T_NOTIFICATION_INFO

    With pInfo
        .Title = pStyle.Name & IIf(Scheme = "<Default>", "", "/" & Scheme) & IIf(IsPriority, " (Priority)", "")
        .Text = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
        .Timeout = -1
        .IconPath = IIf(pStyle.IconPath = "", g_MakePath(App.Path) & "etc\icons\style.png", pStyle.IconPath)
        .StyleName = pStyle.Name
        .SchemeName = LCase$(Scheme)
        .Position = SN_SP_DEFAULT_POS
        .Priority = IIf(IsPriority, 1, 0)
        Set .ClassObj = New TAlert
        .CustomUID = "style-preview" & IIf(IsPriority, "-priority", "")
        .APIVersion = App.Major
        .IntFlags = App.Major

        If (pStyle.Flags And S_STYLE_V42_CONTENT) Then
            ' /* set up for V42 style */
            With New BPackedData
                If (Percent >= 0) And (Percent <= 100) Then _
                    .Add "value-percent", CStr(Percent)
    
                pInfo.OriginalContent = .AsString()

            End With

        Else
            ' /* set up for pre-V42 style */
            If (Percent > 0) And (Percent <= 100) Then _
                .Text = CStr(Percent)
        
        End If

        g_NotificationRoster.Hide 0, .CustomUID, App.ProductName, ""

    End With

    g_DoSchemePreview2 = (g_NotificationRoster.Add(pInfo, Nothing, False) <> 0)

End Function

Public Function g_GetSafeTempIconPath() As String
Dim sz As String
Dim c As Long

    sz = String$(MAX_PATH + 1, 0)
    GetTempPath MAX_PATH, sz
    sz = g_TrimStr(sz)
    If sz = "" Then _
        Exit Function

    sz = g_MakePath(sz)

    c = 1
    Do While g_Exists(sz & "snarl-icon" & CStr(c))
        c = c + 1

    Loop

    g_GetSafeTempIconPath = sz & "snarl-icon" & CStr(c)

End Function

Public Sub g_ConfigLock()

    mConfigLocked = True

End Sub

Public Sub g_ConfigUnlock(Optional ByVal IgnoreDelayedWrite As Boolean)

    mConfigLocked = False
    If (mWriteConfigOnUnlock) And (Not IgnoreDelayedWrite) Then _
        g_WriteConfig

    mWriteConfigOnUnlock = False

End Sub

Public Function g_ConfigIsLocked() As Boolean

    g_ConfigIsLocked = mConfigLocked

End Function

'Public Function new_Class(ByVal Priority As Boolean) As TAlert
'
'    Set new_Class = New TAlert
'    new_Class.bSpecialInit "_id", "_desc", Priority
'
'End Function

'Public Function g_StickyNotifications() As Boolean
'
'    g_StickyNotifications = (g_ConfigGet("sticky_snarls") = "1") 'Or (gIsAway)
'
'End Function

Public Function g_GetStylePath(ByVal StyleToUse As String) As String

    g_GetStylePath = g_MakePath(App.Path) & "etc\default_theme\"
    If (g_StyleRoster Is Nothing) Then _
        Exit Function

    If StyleToUse = "" Then _
        StyleToUse = g_ConfigGet("default_style")

    StyleToUse = style_GetStyleName(StyleToUse)
    If StyleToUse = "" Then _
        Exit Function                               ' // indicates a problem!

Dim pStyle As TStyle

    If g_StyleRoster.Find(style_GetStyleName(StyleToUse), pStyle) Then _
        g_GetStylePath = pStyle.Path

End Function

'Public Function g_RemoveForwarder(ByVal uID As Long)
'
'    Debug.Print "STUB: g_RemoveForwarder(" & CStr(uID) & ")"
'
'End Function

Public Function g_SettingsPath() As String

    If Not (mSettings Is Nothing) Then _
        g_SettingsPath = g_GetPath(mSettings.File)

End Function

Public Sub g_ProcessAck(ByVal Ack As String)

    If g_SafeLeftStr(Ack, 1) = "!" Then
        ' /* bang command */
        uDoBang g_SafeRightStr(Ack, Len(Ack) - 1)
        
    Else
        ' /* treat as URL/file/launchable */
        ShellExecute frmAbout.hWnd, vbNullString, Ack, vbNullString, vbNullString, SW_SHOW
    
    End If

End Sub

Private Sub uDoBang(ByVal Bang As String)
Dim pti As BTagList
Dim Arg() As String
Dim sz As String
Dim pa As TApp
Dim c As Long
Dim i As Long

    Arg() = Split(Bang, " ")
    c = UBound(Arg)
    Set pti = new_BTagList
    ' /* if there are any args, make them into a taglist */
    If c > 0 Then
        For i = 1 To c
            pti.Add new_BTagItem(Arg(i), "")

        Next i
    End If

    Select Case LCase$(Arg(0))

    Case "missed"
        frmAbout.bShowMissedPanel

    Case "notifications"
        ' /* show our prefs panel targeted on the app */
        If pti.CountItems = 1 Then _
            frmAbout.DoAppConfigBySignature pti.TagAt(1).Name

    Case "app_settings"
        ' /* ask the app to show its GUI */
        If pti.CountItems = 1 Then
            If g_AppRoster.PrivateFindBySignature(pti.TagAt(1).Name, pa) Then _
                pa.DoSettings 0

        End If

    Case "system"
        uProcessSystem pti

    Case "configure"
        If pti.CountItems > 0 Then
            sz = pti.TagAt(1).Name
            Select Case g_GetExtension(sz, True)
            Case "extension"
                uManageExtension Arg(0), g_RemoveExtension(sz)

            End Select
        End If

    Case Else
        g_Debug "uDoBang(): unknown command '" & Arg(0) & "'"

    End Select

End Sub

Private Sub uProcessSystem(ByRef Args As BTagList)
Dim pti As BTagItem

    Set pti = Args.TagAt(1)

    If (pti Is Nothing) Then _
        Exit Sub

    Select Case LCase$(pti.Name)

    Case "shutdown_dialog", "shutdown"
        SHShutdownDialog 0

    Case "run_dialog", "run"
        SHRunDialog 0, 0, vbNullString, vbNullString, vbNullString, SHRD_DEFAULT

    Case "lock"
        LockWorkStation

    Case "about"
        ShellAbout 0, vbNullString, vbNullString, 0

    Case "access"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL access.cpl,,0", vbNullString, SW_SHOW
    
    Case "datetime"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL timedate.cpl,,0", vbNullString, SW_SHOW
    
    Case "display"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL desk.cpl,,3", vbNullString, SW_SHOW
    
    Case "fonts"
        ShellExecute 0, "open", "control.exe", "fonts", vbNullString, SW_SHOW
    
    Case "game"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL joy.cpl,,0", vbNullString, SW_SHOW
    
    Case "software"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL appwiz.cpl,,0", vbNullString, SW_SHOW

    Case "keyboard"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL main.cpl,@1,0", vbNullString, SW_SHOW

    Case "locale"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL intl.cpl,,0", vbNullString, SW_SHOW
    
    Case "mouse"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL main.cpl,,0", vbNullString, SW_SHOW
    
    Case "network"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL ncpa.cpl,,0", vbNullString, SW_SHOW

    Case "power"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL powercfg.cpl,,0", vbNullString, SW_SHOW

    Case "printers"
        ShellExecute 0, "open", "control.exe", "printers", vbNullString, SW_SHOW

    Case "screensaver"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL desk.cpl,,1", vbNullString, SW_SHOW

    Case "sounds"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL mmsys.cpl,,1", vbNullString, SW_SHOW

    Case "admin"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL sysdm.cpl,,0", vbNullString, SW_SHOW

    Case "telephony"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL telephon.cpl,,0", vbNullString, SW_SHOW

    Case "theme"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL desk.cpl,,2", vbNullString, SW_SHOW

    Case "users"
        ShellExecute 0, "open", "control.exe", "userpasswords", vbNullString, SW_SHOW

    Case "wallpaper"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL desk.cpl,,0", vbNullString, SW_SHOW

    Case "controlpanel"
        ShellExecute 0, "open", "control.exe", vbNullString, vbNullString, SW_SHOW

    Case "trash"
        ShellExecute 0, "open", "::{645FF040-5081-101B-9F08-00AA002F954E}", vbNullString, vbNullString, SW_SHOW

    Case "mycomputer"
        ShellExecute 0, "open", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNullString, vbNullString, SW_SHOW

    Case "nethood"
        ShellExecute 0, "open", "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNullString, vbNullString, SW_SHOW

    End Select

End Sub

Public Sub g_SetLastError(ByVal Error As SNARL_STATUS_CODE)

    SetProp ghWndMain, "last_error", Error

End Sub


'        MsgBox "As this is the first time Snarl has been run, I need to test the drawing" & vbCrLf & _
'               "speed of your PC.  This test will only take a few seconds, if that.", _
'               vbOKOnly Or vbInformation, "Snarl Graphics Test"
'
'        l = GetTickCount()
'        g_Debug "main(): GFXCHK: starting graphics test (tick_count=" & CStr(l) & ")"
'
'Dim pInfo As T_NOTIFICATION_INFO
'
'        With pInfo
'            .Title = "Snarl Graphics Test"
'            .Text = "Test Message"
'            .Timeout = 1
'            .StyleToUse = ""            ' // the scheme
'
'        End With
'
'        With New CSnarlWindow
'            .Create New TAlert, pInfo, New TInternalStyle, 0, ""
'            '"Snarl Graphics Test", "Test message", 1, "", 0, 0, 0, "", New TInternalStyle, "", 0
'            .Quit
'
'        End With
'
'Dim dStep As Double
'
'        l = GetTickCount() - l
'        g_Debug "main(): GFXCHK: completed graphics test (tick_count=" & CStr(GetTickCount()) & ")"
'
'        dStep = l / 422#
'
'        If dStep < 1# Then _
'            dStep = 1#
'
'        g_Debug "main(): GFXCHK: delta=" & l & " actual=" & CStr(l / 422#) & " step=" & dStep
'
'        MsgBox "Graphics test complete.  Snarl has calculated a factor of " & Format$(dStep, "0.00") & vbCrLf & _
'               "to use when displaying messages.", vbOKOnly Or vbInformation, "Test Complete"

'Public Function g_StyleConfigGet(ByVal Name As String, Optional ByVal Default As String) As String
'
'    ' /* pre-set with default */
'
'    g_StyleConfigGet = mDefaults.ValueOf(Name)
'
'Dim sz As String
'
'    If Not (mConfig Is Nothing) Then
'        If mConfig.Find(Name, sz) Then _
'            g_ConfigGet = sz
'
'    End If
'
'End Function
'
''Public Sub g_ConfigSet(ByVal Name As String, ByVal Value As String)
''
''    If (mConfig Is Nothing) Then _
''        Exit Sub
''
''    mConfig.Update Name, Value
''    g_WriteConfig
''
''End Sub

Public Sub g_DoManualUpdateCheck()

    If mUpdateCheck.Check(True, "http://www.fullphat.net/updateinfo/snarl.updateinfo") Then
        g_Debug "g_DoManualUpdateCheck(): check initiated"

    Else
        g_Debug "g_DoManualUpdateCheck(): failed to start", LEMON_LEVEL_CRITICAL

    End If

End Sub

Public Function g_TrimLines(ByVal Text As String) As String

    Text = Replace$(Text, "\n", vbCrLf)

    ' /* pre-set default */

    g_TrimLines = Text

Dim n As Long

    n = g_SafeLong(g_ConfigGet("style.overflow_limit"))
    If (n < 4) Or (n > 12) Then _
        n = 12                  ' // must be at least 5 lines to make it meaningful to
                                ' // take up a line with the ellipsis
                                ' // i.e. line 1
                                ' //      ...
                                ' //      line 3
                                ' //      line 4
                                ' // is pointless as only line 2 is replaced

Dim sz() As String
Dim c As Long
Dim i As Long

    sz() = Split(Text, vbCrLf)
    c = UBound(sz)

    If (c + 1) > n Then
        ' /* <first line> \n <...> \n <remaining lines> */
        g_TrimLines = sz(0) & vbCrLf & "…"
        For i = c - (n - 3) To c
            g_TrimLines = g_TrimLines & vbCrLf & sz(i)

        Next i

    End If

End Function

'Public Function g_CreateBadge(ByVal Content As String) As mfxBitmap
'Const RX = 6
'Dim pr As BRect
'
'    With New mfxView
'        .SetFont "Tahoma", 7, True
'        .TextMode = MFX_TEXT_ANTIALIAS
'
'        Set pr = new_BRect(0, 0, Max(.StringWidth(Content), .StringHeight("A")), .StringHeight("A"))
'        pr.ExpandBy 8, 8
'
'        .SizeTo pr.Width, pr.Height
'        .EnableSmoothing True
'
'        .SetHighColour rgba(0, 0, 0, 190)
'        .SetLowColour rgba(0, 0, 0, 140)
'        .FillRoundRect pr, RX, RX, MFX_VERT_GRADIENT
'
'        .SetHighColour rgba(255, 255, 255)
'        .DrawString Content, pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER
'
'        .SetHighColour rgba(255, 255, 255)
'        .StrokeRoundRect pr.InsetByCopy(1, 1), RX, RX, 2
'
'        .SetHighColour rgba(0, 0, 0, 150)
'        .StrokeRoundRect pr, RX, RX, 1
'        .StrokeRoundRect pr.InsetByCopy(3, 3), RX, RX, 1
'
'        Set g_CreateBadge = .ConvertToBitmap()
'
'    End With
'
'End Function

Public Sub g_KludgeNotificationInfo(ByRef nInfo As T_NOTIFICATION_INFO, Optional ByRef pPacked As BPackedData)

    ' /* translates current T_NOTIFICATION_INFO content into a BPackedString
    '    and then stores that back into T_NOTIFICATION_INFO->OriginalContent
    '    this is so V42 styles can access the entire request */

    If (nInfo.ClassObj Is Nothing) Then
        g_Debug "g_KludgeNotificationInfo(): missing ClassObj", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

'Dim ppacked As BPackedData

    Set pPacked = New BPackedData

    With nInfo
        .Title = Replace$(.Title, "\n", vbCrLf)
        .Text = Replace$(.Text, "\n", vbCrLf)

        pPacked.Add "id", .ClassObj.Name

        If .Title <> "" Then _
            pPacked.Add "title", .Title

        If .Text <> "" Then _
            pPacked.Add "text", .Text

        If .Timeout <> 0 Then _
            pPacked.Add "timeout", CStr(.Timeout)

        If .IconPath <> "" Then _
            pPacked.Add "icon", .IconPath

        If .Priority <> 0 Then _
            pPacked.Add "priority", CStr(.Priority)

        If .DefaultAck <> "" Then _
            pPacked.Add "callback", .DefaultAck

        pPacked.Add "value", .Value

'        If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
            .Add "merge", "1"

    End With

Dim ppx As BPackedData
Dim szn As String
Dim szv As String

    ' /* add in all other custom content */

    Set ppx = New BPackedData
    With ppx
        If .SetTo(nInfo.OriginalContent) Then
            .Rewind
            Do While .GetNextItem(szn, szv)
                If Not pPacked.Exists(LCase$(szn)) Then
                    Select Case LCase$(szn)
                    Case "password"
                        Debug.Print "password removed"

                    Case Else
                        pPacked.Add szn, szv

                    End Select
                End If
            Loop
        End If
    End With

    nInfo.OriginalContent = pPacked.AsString()

'    Set nInfo.Content = pPacked










'
'
'    ' /* generates a packed string from the provided T_NOTIFICATION_INFO
'    '    struct and assigns the string to the OriginalContent element,
'    '    which is required by V42 styles */
'
'    With nInfo
'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)
'
'        .OriginalContent = "id::" & .ClassObj.Name & _
'                           "#?title::" & .Title & _
'                           "#?text::" & .Text & _
'                           "#?timeout::" & CStr(.Timeout) & _
'                           "#?icon::" & .IconPath & _
'                           "#?priority::" & CStr(.Priority) & _
'                           "#?ack::" & .DefaultAck & _
'                           "#?value::" & .Value
'
''        If (.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
''            .OriginalContent = .OriginalContent & "#?merge::1"
'
'    End With

End Sub

Public Function g_TranslateIconPath(ByVal Icon As String, ByVal StylePath As String) As String
Dim pbm As mfxBitmap
Dim pIcon As BIcon
Dim sz As String
'Dim dw As Long
Dim i As Long

    If g_SafeLeftStr(Icon, 1) = "!" Then
        ' /* convert the icon into it's corresponding file */
        Icon = g_SafeRightStr(Icon, Len(Icon) - 1) & ".png"

        ' /* default location */

        g_TranslateIconPath = g_MakePath(App.Path) & "etc\default_theme\icons\" & Icon

        ' /* here, 'StylePath' sould be either empty or contain the full path to the
        '    current style.  If it's the former we use the built-in icon set; if
        '    it's the latter we try to get the icon from the style */

        If (g_ConfigGet("use_style_icons") = "1") And (StylePath <> "") Then
            ' /* try to get it from the style... */
            sz = g_MakePath(StylePath) & "icons\" & Icon
            If g_Exists(sz) Then _
                g_TranslateIconPath = sz

        End If

    ElseIf g_SafeLeftStr(Icon, 1) = "%" Then
        ' /* whatever is after the % should be a valid HICON */
        Set pbm = uGetIcon(g_SafeLong(g_SafeRightStr(Icon, Len(Icon) - 1)))
        If NOTNULL(pbm) Then
            g_TranslateIconPath = g_GetSafeTempIconPath()
            pbm.Save g_TranslateIconPath, "image/png"

        Else
            g_Debug "TNotificationRoster.g_TranslateIconPath(): bad HICON '" & Icon & "'", LEMON_LEVEL_WARNING

        End If

    Else

        sz = LCase$(Icon)

        If g_GetExtension(sz) = "ico" Then
            ' /* icon */
            g_TranslateIconPath = uLoadICO(Icon)

        ElseIf InStr(sz, ".ico,") > 0 Then
            ' /* icon (sometimes a ",0" can appear) */
            i = InStr(Icon, ",")
            g_TranslateIconPath = uLoadICO(g_SafeLeftStr(Icon, i - 1))

        ElseIf (InStr(sz, ".dll,") > 0) Or (InStr(sz, ".exe,") > 0) Then
            ' /* icon within a resource file */
            i = InStr(Icon, ",")
            g_TranslateIconPath = uGetBestIcon(g_SafeLeftStr(Icon, i - 1), g_SafeLong(g_SafeRightStr(Icon, Len(Icon) - i)))

        Else
            g_TranslateIconPath = Icon

        End If

    End If

End Function

Private Function uLoadICO(ByVal IconPath As String) As String
Dim pbm As mfxBitmap
Dim pm As MImage
Dim pIcon As BIcon

    g_Debug "uLoadIcon()", LEMON_LEVEL_PROC_ENTER

    On Error Resume Next

    g_Debug "creating BIconContent..."
    With New BIconContent
        g_Debug "LoadFromICO('" & IconPath & "')"
        If Not .LoadFromICO(IconPath) Then _
            g_Debug "failed", LEMON_LEVEL_PROC_EXIT: _
            Exit Function

        g_Debug ".GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS)"
        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pIcon) Then
            g_Debug "success - icon is " & pIcon.Width & "x" & pIcon.Height & "x" & pIcon.ColourDepth
            
            ' /* seems to be an issue rendering 128x128 icons */
            
            If (pIcon.Width > 127) Or (pIcon.Height > 127) Then
                g_Debug "aborting due to bug"
            
            Else
                g_Debug "rendering..."
                Set pm = pIcon.Render()
                g_Debug "creating bitmap..."
                Set pbm = create_bitmap_from_image(pm)
                uLoadICO = g_GetSafeTempIconPath()
                g_Debug "saving as '" & uLoadICO & "'..."
                pbm.Save uLoadICO, "image/png"

            End If

        Else
            g_Debug "failed"

        End If

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function uGetBestIcon(ByVal IconPath As String, ByVal Index As Long) As String
Dim pbm As mfxBitmap
Dim pIcon As BIcon

    On Error Resume Next

    With New BIconContent
        If Not .LoadFromResource(IconPath, Index) Then _
            Exit Function

        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pIcon) Then
            Set pbm = create_bitmap_from_image(pIcon.Render)
            uGetBestIcon = g_GetSafeTempIconPath()
            pbm.Save uGetBestIcon, "image/png"

        End If

    End With

End Function

Private Function uGetIcon(ByVal hIcon As Long) As mfxBitmap

    On Error Resume Next

    If hIcon = 0 Then _
        Exit Function

Dim pi As BIcon

    err.Clear
    Set pi = New BIcon
    If err.Number <> 0 Then
        g_Debug "TNofiticationRoster.uGetIcon(): no icon.resource", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    If Not pi.SetFromHICON(hIcon) Then
        g_Debug "TNofiticationRoster.uGetIcon(): bad icon handle", LEMON_LEVEL_CRITICAL
        Exit Function

    End If
        
    Set uGetIcon = create_bitmap_from_image(pi.Render())

End Function

Public Function g_ShowRequest(ByVal lPid As Long, ByRef Data As BPackedData) As Long
Dim pReq As TRequester

    Set pReq = New TRequester
    g_ShowRequest = pReq.Go(lPid, Data)

    If g_ShowRequest <> 0 Then

        ' /* disable existing */
        If mReqs > 0 Then _
            mReq(mReqs).SetEnabled False

        mReqs = mReqs + 1
        ReDim Preserve mReq(mReqs)
        Set mReq(mReqs) = pReq

    End If

End Function

Public Sub g_PopRequest2()

    If mReqs < 1 Then _
        Exit Sub

    mReqs = mReqs - 1
    ReDim Preserve mReq(mReqs)
    If mReqs > 0 Then _
        mReq(mReqs).SetEnabled True

End Sub

Public Function g_CreatePacked(ByVal ClassId As String, ByVal Title As String, ByVal Text As String, Optional ByVal Timeout As Long = -1, Optional ByVal Icon As String, Optional ByVal Priority As Long = 0, Optional ByVal Ack As String, Optional ByVal Flags As SNARL41_NOTIFICATION_FLAGS, Optional ByVal Password As String, Optional ByVal AddTestAction As Boolean, Optional ByVal uID As String, Optional ByVal Percent As Long = -1) As BPackedData

    ' /* translate notification arguments into packed data
    '    currently this is only used by g_PrivateNotify()
    '    but it's flexible enough to be used elsewhere */

    Set g_CreatePacked = New BPackedData
    With g_CreatePacked

        .Add "app-sig", App.ProductName

        If ClassId <> "" Then _
            .Add "class", ClassId

        If Title <> "" Then _
            .Add "title", Title

        If Text <> "" Then _
            .Add "text", Text

        .Add "timeout", Timeout

        If Icon <> "" Then _
            .Add "icon", Icon

        .Add "priority", CStr(Priority)
        
        If Ack <> "" Then _
            .Add "ack", Ack

        .Add "flags", Hex$(Flags)               ' // flags are sent as a hex value

        If Password <> "" Then _
            .Add "password", Password

        If AddTestAction Then
            .Add "action", "Dummy Action,@1"
            .Add "action", "Dummy Action,@2"
            .Add "action", "Dummy Action,@3"

        End If

        If uID <> "" Then _
            .Add "uid", uID

        If Percent > -1 Then _
            .Add "value-percent", CStr(Percent)

    End With

End Function

Public Function g_QuickAddAction(ByVal Token As Long, ByVal Label As String, ByVal Command As String) As Long

    g_QuickAddAction = g_NotificationRoster.AddAction(Token, g_newBPackedData("label::" & Label & "#?cmd::" & Command))

End Function

Public Function g_QuickLastError() As SNARL_STATUS_CODE

    g_QuickLastError = GetProp(ghWndMain, "last_error")

End Function

'Public Function g_QuickAddClass(ByVal AppToken As Long, ByVal Id As String, ByVal Name As String, Optional ByVal Enabled As Boolean, Optional ByVal Password As String) As Long
'Dim pp As BPackedData
'
'    Set pp = New BPackedData
'    With pp
'        .Add "id", Id
'        .Add "name", Name
'        .Add "enabled", IIf(Enabled, "1", "0")
'        If Password <> "" Then _
'            .Add "password", Password
'
'    End With
'
'    g_QuickAddClass = g_DoAction("addclass", AppToken, pp)
'
'End Function

Public Function g_PrivateNotify(Optional ByVal ClassId As String = SNARL_CLASS_GENERAL, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Timeout As Long = -1, Optional ByVal Icon As String, Optional ByVal Priority As Long = 0, Optional ByVal Ack As String, Optional ByVal Flags As SNARL41_NOTIFICATION_FLAGS, Optional ByVal IntFlags As SN_NOTIFICATION_FLAGS, Optional ByVal AddTestAction As Boolean, Optional ByVal uID As String, Optional ByVal Percent As Long = -1, Optional ByVal IncludeNow As Boolean) As Long

    ' /* safe internal notification generator
    '
    '    uses g_DoNotify() to display a Snarl-generated notification
    '    without going via the Win32 messaging system
    '
    ' */

    If g_SafeLeftStr(Icon, 1) = "." Then _
        Icon = g_MakePath(App.Path) & "etc\icons\" & g_SafeRightStr(Icon, Len(Icon) - 1) & ".png"

Dim ppd As BPackedData

    Set ppd = g_CreatePacked(ClassId, Title, Text, Timeout, Icon, Priority, Ack, Flags, gSnarlPassword, AddTestAction, uID, Percent)
    If IncludeNow Then _
        ppd.Add "value-date-packed", Format$(Now(), "YYYYMMDDHHNNSS")

    ppd.Add "log", "0"

    g_PrivateNotify = g_DoNotify(0, _
                                 ppd, _
                                 Nothing, _
                                 IntFlags Or App.Major, _
                                 "", 0)

End Function

Public Function g_DoNotify(ByVal AppToken As Long, ByRef pData As BPackedData, ByRef ReplySocket As CSocket, ByVal IntFlags As SN_NOTIFICATION_FLAGS, ByVal RemoteHost As String, ByVal SenderPID As Long) As Long

    ' /* master notification generator
    '
    '    all roads should lead here - there should be no use of sn41EZNotify() or any other
    '    Win32 API function.  Similarly, there should be no by-passing of this function,
    '    except in very specific circumstances (style previews, for example) */

    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Then
        g_Debug "g_DoNotify(): app and/or notification roster missing", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

    If (pData Is Nothing) Then
        g_Debug "g_DoNotify(): arg missing", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If


    ' /* R2.5 Beta 2: if "app-sig" and "app-title" exists, we can do a register as well */

    If (pData.Exists("app-sig")) And (pData.Exists("app-title")) Then
        g_DoNotify = g_AppRoster.Add41(pData, ReplySocket, SenderPID, RemoteHost)
        If g_DoNotify < 0 Then
            g_Debug "g_DoNotify(): pre-registration failed: " & str(g_DoNotify), LEMON_LEVEL_CRITICAL
            Exit Function

        End If

        ' /* R2.5 Beta 2: if "class-id" exists, we can add the class as well */

        If pData.Exists("class-id") Then
            g_DoNotify = uAddClass(0, pData)
            If g_DoNotify <> -1 Then
                g_DoNotify = g_QuickLastError()
                g_Debug "g_DoNotify(): class registration failed: " & str(g_DoNotify), LEMON_LEVEL_CRITICAL
                Exit Function

            End If
        End If
    End If


    ' /* look for the new "replace-uid" and "update-uid" and "merge-uid" args:
    '    "replace" will remove the notification with the specified uid if it's
    '    still on-screen; "update-uid" will cause the specified notification
    '    to be updated with this content and "merge-uid" will cause the
    '    provided content to be merged with the existing notification */

Dim pn As TNotification

    If pData.Exists("replace-uid") Then
        ' /* if the specified uid (NOT token) exists, remove it */
        g_NotificationRoster.Hide 0, pData.ValueOf("replace-uid"), pData.ValueOf("app-sig"), pData.ValueOf("password")

    ElseIf pData.Exists("update-uid") Then
        ' /* if the specified uid (NOT token) exists, update with this content otherwise create a new notification */
        g_Debug "g_DoNotify(): looking for (update-)uid: " & pData.ValueOf("update-uid") & "..."

        If g_NotificationRoster.Find(0, pData.ValueOf("update-uid"), pData.ValueOf("app-sig"), pData.ValueOf("password"), pn) Then
            pn.UpdateOrMerge pData, False
            g_DoNotify = pn.Info.Token
            Exit Function

        End If

    ElseIf pData.Exists("merge-uid") Then
        ' /* if the specified uid (NOT token) exists, merge this content with that one, otherwise create a new notificaton */
        g_Debug "g_DoNotify(): looking for (merge-)uid: " & pData.ValueOf("merge-uid") & "..."

        If g_NotificationRoster.Find(0, pData.ValueOf("merge-uid"), pData.ValueOf("app-sig"), pData.ValueOf("password"), pn) Then
            pn.UpdateOrMerge pData, True
            g_DoNotify = pn.Info.Token
            Exit Function

        End If

    End If

    ' /* this still takes effect even if other options above have been used */

    If pData.Exists("uid") Then
        ' /* if the specified uid (NOT token) exists, update this content with that one, otherwise create a new notificaton */
        g_Debug "g_DoNotify(): looking for uid '" & pData.ValueOf("uid") & "' from app '" & pData.ValueOf("app-sig") & "'..."

        If g_NotificationRoster.Find(0, pData.ValueOf("uid"), pData.ValueOf("app-sig"), pData.ValueOf("password"), pn) Then
            pn.UpdateOrMerge pData, False
            g_DoNotify = pn.Info.Token
            Exit Function

        End If
    End If

    ' /* R2.4.2 DR3: now we've checked for updates and merges, we check to see if at least one of
    '    title, text or icon exists and fail with 109 if not */

'    If (Not pData.Exists("title")) And (Not pData.Exists("text")) And (Not pData.Exists("title")) Then
'        g_Debug "g_DoNotify(): must supply at least one of 'title', 'text' or 'icon'", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_ARG_MISSING
'        Exit Function
'
'    End If

Dim szClass As String
Dim pApp As TApp

    ' /* R2.4 DR7: if "app-sig" argument is specified then look for the app by signature */

    If pData.ValueOf("app-sig") <> "" Then
        If g_AppRoster.FindBySignature(pData.ValueOf("app-sig"), pApp, pData.ValueOf("password")) Then
            ' /* R2.4.1 - support for "class" keyword */
            If pData.Exists("class") Then
                szClass = pData.ValueOf("class")

            Else
                szClass = pData.ValueOf("id")

            End If

        Else
            ' /* not found / auth failure (lasterror will have been set) */
            Exit Function

        End If

    Else
        ' /* special case: if the app token is 0 we use ourself as the sending app  */
        If AppToken = 0 Then

            ' /* R2.4.2: new security setting */

            If g_ConfigGet("apps_must_register") = "1" Then
                g_Debug "g_DoNotify(): not allowed: applications must register first", LEMON_LEVEL_CRITICAL
                g_SetLastError SNARL_ERROR_NOT_REGISTERED
                Exit Function

            ElseIf g_AppRoster.FindByToken(gSnarlToken, pApp, gSnarlPassword) Then
                ' /* if we're using the Snarl app, we need the anonymous class */
                g_Debug "g_DoNotify(): using Snarl anonymous class"
                szClass = IIf((IntFlags And SN_NF_REMOTE) = 0, SNARL_CLASS_ANON, SNARL_CLASS_ANON_NET)

            Else
               ' /* Snarl's registration not found */
                g_Debug "g_DoNotify(): Snarl internal app not in roster", LEMON_LEVEL_CRITICAL
                g_SetLastError SNARL_ERROR_SYSTEM
                Exit Function

            End If

        ElseIf g_AppRoster.FindByToken(AppToken, pApp, pData.ValueOf("password")) Then
            ' /* R2.4.1 - support for "class" keyword */
            If pData.Exists("class") Then
                szClass = pData.ValueOf("class")

            Else
                szClass = pData.ValueOf("id")

            End If

        Else
            ' /* not found / auth failure (lasterror will have been set) */
            Exit Function

        End If

    End If

    ' /* include the remote sender (if there is one) - small kludge here for
    '    Growl/UDP which won't have a reply socket */

'Dim szRemoteHost As String
'
'    If Not (ReplySocket Is Nothing) Then
'        ' /* socket-based sender */
'        If ReplySocket.RemoteHost <> "" Then
'            szRemoteHost = ReplySocket.RemoteHost '& " (" & ReplySocket.RemoteHostIP & ")"
'
'        Else
'            szRemoteHost = ReplySocket.RemoteHostIP
'
'        End If
'
'    Else
'        ' /* non-socket sender (e.g. Growl via UDP) */
'        szRemoteHost = RemoteHost
'
'    End If

    ' /* R2.4 DR7 - merging is now controlled via an internal flag */

    If pData.ValueOf("merge") = "1" Then _
        IntFlags = IntFlags Or SN_NF_MERGE

    ' /* now we have the app object and we know the class, we can pass it over */

    g_DoNotify = pApp.Show41(szClass, pData, ReplySocket, IntFlags) '//, szRemoteHost)

End Function

Public Function g_DoAction(ByVal action As String, ByVal Token As Long, ByRef Args As BPackedData, Optional ByVal InternalFlags As SN_NOTIFICATION_FLAGS, Optional ByRef ReplySocket As CSocket, Optional ByVal SenderPID As Long) As Long

    ' /* this is the central hub for all incoming requests, be they from SNP, Growl/UDP
    '    or Win32.  "Token" here can be either the app token or the notification token;
    '    the action determines which one */

    ' /* Return zero on error (and set lasterror), -1 or a +ve value on success - whatever
    '    called this will figure out what to do with the return value */

    If (g_AppRoster Is Nothing) Then
        g_SetLastError SNARL_ERROR_SYSTEM
        g_Trap SOS_MISSING_ROSTER, "AppRoster"
        Exit Function

    End If

    If (g_NotificationRoster Is Nothing) Then
        g_SetLastError SNARL_ERROR_SYSTEM
        g_Trap SOS_MISSING_ROSTER, "NotificationRoster"
        Exit Function

    End If

    If (Args Is Nothing) Then
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

Dim hr As SNARL_STATUS_CODE
Dim szSource As String
Dim pApp As TApp

    ' /* assume all okay... */

    g_SetLastError 0
    If Not (ReplySocket Is Nothing) Then _
        szSource = ReplySocket.RemoteHostIP

    ' /* R2.4.2 DR3: check RemoteHostName is not actually one of our locally-assigned addresses */

'    If szSource <> "" Then
'        g_Debug "g_DoAction(): RemoteIP=" & szSource
'
'        ' /* is it actually a local IP? */
'        If InStr(get_ip_address_table(), szSource) <> 0 Then _
'            szSource = ""
'
'        g_Debug "g_DoAction(): FilteredRemoteIP=" & szSource
'
'    End If
'
'Dim pArgs As BPackedData
'Dim sSig As String
'Dim sn As String
'Dim sv As String
'
'    ' /* if this is routed from a remote computer and it's using "app-sig"
'    '    tack on the remote ip address to the provided signature */
'
'    If (szSource <> "") And (Args.Exists("app-sig")) Then
'        Set pArgs = New BPackedData
'        With Args
'            sSig = .ValueOf("app-sig")
'
'            .Rewind
'            Do While .GetNextItem(sn, sv)
'                If sn <> "app-sig" Then _
'                    pArgs.Add sn, sv
'
'            Loop
'        End With
'
'        pArgs.Add "app-sig", sSig & ":" & szSource
'        Set Args = pArgs
'
'    End If

    Select Case action

    Case "addaction"
        g_DoAction = g_NotificationRoster.AddAction(Token, Args)

    Case "addclass"
        g_DoAction = uAddClass(Token, Args)

    Case "clearclasses", "killclasses"
        g_DoAction = uRemClass(Token, Args, True)

    Case "clearactions"
        g_DoAction = g_NotificationRoster.ClearActions(Token, Args)

    Case "hello"
        ' /* reply our major version number */
        ' /* To-do: reply with an error message if Snarl isn't
        '    accepting requests, or DND mode enabled? */
        g_DoAction = App.Major

    Case "hide"
        g_DoAction = CLng(g_NotificationRoster.Hide(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password")))

    Case "isvisible"
        g_DoAction = CLng(g_NotificationRoster.IsVisible(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password")))

    Case "notify"
        g_DoAction = g_DoNotify(Token, Args, ReplySocket, InternalFlags, szSource, SenderPID)

    Case "reg", "register"
        g_DoAction = g_AppRoster.Add41(Args, ReplySocket, SenderPID, szSource)

    Case "remclass"
        g_DoAction = uRemClass(Token, Args)

    Case "test"
        ' /* only available when Snarl is running in debug mode */
        If gDebugMode Then
            g_PrivateNotify "", _
                            IIf(Args.ValueOf("alpha") = "", "Snarl", Args.ValueOf("alpha")), _
                            IIf(Args.ValueOf("beta") = "", "Test Message", Args.ValueOf("beta"))
            g_DoAction = -1

        Else
            g_SetLastError SNARL_ERROR_UNKNOWN_COMMAND
            g_DoAction = 0

        End If

    Case "unreg", "unregister"
        ' /* R2.4 DR7: can unregister using signature/password combo */
        If Args.Exists("app-sig") Then
            g_DoAction = g_AppRoster.UnregisterBySig(Args.ValueOf("app-sig"), Args.ValueOf("password"))

        Else
            g_DoAction = g_AppRoster.Unregister(Token, Args.ValueOf("password"), False)

        End If

    Case "update"
        g_DoAction = g_NotificationRoster.Update(Token, Args)

    Case "updateapp", "update_app"
        g_DoAction = g_AppRoster.Update(Token, Args)

    Case "version"
        g_DoAction = GetProp(ghWndMain, "_version")


    ' /* V42 only (no corresponding V41 command ID) */


    Case "wasmissed"
        g_DoAction = g_NotificationRoster.WasMissed(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"))

'    Case "merge"
'        ' /* specify an existing token or uid/app-sig pair that identifies the notification
'        '    to merge with.  Creates a new notification (same uid, different token) if
'        '    specified notification doesn't exist */
'        g_DoAction = g_NotificationRoster.Merge(Token, Args)

    Case "setmode"
        If Args.Exists("busy") Then _
            g_DoAction = uSetBusy(Token, Args)



    ' /* undocumented/unsupported - either private or due for public release in a future revision */

    Case "snarl"
        ' /* PRIVATE: for internal use only under V43, to be made public in V44 */
        If (g_ConfigGet("block_net_control") = "1") And (szSource <> "") Then
            g_SetLastError SNARL_ERROR_ACCESS_DENIED
            g_DoAction = 0

        Else
            uDoSystemRequest Args
            g_DoAction = (g_QuickLastError() = SNARL_SUCCESS)

        End If

    Case "request"
        ' /* PRIVATE: for internal use only under V42 */
        g_DoAction = g_ShowRequest(Token, Args)


    Case "subscribe"
        ' /* undocumented/unsupported in 2.4.2/2.5; official in V43 */
        If (ReplySocket Is Nothing) Then
            ' /* can be sent via SNP3 only as ReplySocket is required */
            g_Debug "g_DoAction(): {subscribe}: missing reply socket", LEMON_LEVEL_CRITICAL
            g_SetLastError SNARL_ERROR_BAD_SOCKET
            g_DoAction = 0

        Else
            g_Debug "g_DoAction(): {subscribe} source=" & ReplySocket.RemoteHostIP & ":" & ReplySocket.RemotePort
            hr = g_SubsRoster.AddSubscriber("snp", ReplySocket, Args)
            If hr = SNARL_SUCCESS Then
                g_DoAction = -1

            Else
                g_Debug "g_DoAction(): {subscribe}: failed (" & CStr(hr) & ")", LEMON_LEVEL_CRITICAL
                g_SetLastError hr
                g_DoAction = 0

            End If
        End If


    Case "unsubscribe"
        ' /* undocumented/unsupported in 2.4.2/2.5; official in V43 */
        If (ReplySocket Is Nothing) Then
            ' /* can be sent via SNP3 only as ReplySocket is required */
            g_Debug "g_DoAction(): {unsubscribe}: missing reply socket", LEMON_LEVEL_CRITICAL
            g_SetLastError SNARL_ERROR_BAD_SOCKET
            g_DoAction = 0

        Else
            g_Debug "g_DoAction(): {unsubscribe}: source=" & ReplySocket.RemoteHostIP & ":" & ReplySocket.RemotePort
            If g_SubsRoster.RemoveSubscriber(ReplySocket, Args) Then _
                g_DoAction = -1

        End If


    Case Else
        g_SetLastError SNARL_ERROR_UNKNOWN_COMMAND
        g_DoAction = 0

    End Select

End Function

Public Function g_newBPackedData(ByVal Content As String) As BPackedData

    Set g_newBPackedData = New BPackedData
    g_newBPackedData.SetTo Content

End Function

'Public Function g_DoMerge(ByRef Args As BPackedData) As Boolean
'
'    ' /* token may be null, in which case we must have an app-sig/uid pair */
'
'Dim i As Long
'
'    If (Args.Exists("app-sig")) And (Args.Exists("uid")) Then
'        i = g_NotificationRoster.UIDToToken(Args.ValueOf("app-sig"), Args.ValueOf("uid"), Args.ValueOf("password"))
'
'    Else
'        i = g_SafeLong(Args.ValueOf("token"))
'
'    End If
'
'
'    If i Then
'        ' /* merge with this one */
''        mItem(i).Window.MergeWith
'
'    Else
'        ' /* not found? create new... */
'
'
'
'    End If
'
'    g_SetLastError SNARL_ERROR_UNKNOWN_COMMAND
'    Exit Function
'
'
''    If i Then
''        g_Debug "TNotificationRoster.Update(): '" & g_HexStr(Token) & "' found"
''        Update = mItem(i).Window.Update(Args)
''
''    Else
''        i = uFindInMissedList(Token)
''        If i Then
''            g_Debug "TNotificationRoster.Update(): '" & g_HexStr(Token) & "' is in missed list"
''
''        Else
''            Update = False
''
''        End If
''
''    End If
'
'End Function





'Public Function g_DoNotify(ByVal Token As Long, ByRef pData As BPackedData, Optional ByRef ReplySocket As CSocket) As Long
'
'    ' /* sanity checking */
'
'    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Then
'        g_Debug "g_DoNotify(): app and/or notification roster missing", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_SYSTEM
'        Exit Function
'
'    End If
'
'    If (pData Is Nothing) Then
'        g_Debug "g_DoNotify(): arg missing", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_ARG_MISSING
'        Exit Function
'
'    End If
'
'Dim szClass As String
'Dim pApp As TApp
'
'    ' /* special case: if the app token is 0 we use ourself as the sending app  */
'
'    If Token = 0 Then
'        If g_AppRoster.FindByToken(gSnarlToken, pApp, "") Then                  ' // <--- Snarl should be password protected?
'            ' /* if we're using the Snarl app, we need the anonymous class */
'            g_Debug "g_DoNotify(): using Snarl anonymous class"
'            szClass = SNARL_CLASS_ANON
'
'        Else
'            ' /* Snarl's registration not found */
'            g_Debug "g_DoNotify(): Snarl internal app not in roster", LEMON_LEVEL_CRITICAL
'            g_SetLastError SNARL_ERROR_SYSTEM
'            Exit Function
'
'        End If
'
'    ElseIf g_AppRoster.FindByToken(Token, pApp, pData.ValueOf("password")) Then
'        szClass = pData.ValueOf("id")
'
'    Else
'        ' /* not found / auth failure (lasterror will have been set) */
'        Exit Function
'
'    End If
'
'
''Dim pInfo As T_NOTIFICATION_INFO
''Dim i As Long
''
''    With pInfo
''        .hWndReply = Val(pData.ValueOf("hwnd"))
''        .uReplyMsg = Val(pData.ValueOf("umsg"))
''        .IconPath = pData.ValueOf("icon")
''        .Text = pData.ValueOf("text")
''
''        If pData.Exists("timeout") Then
''            .Timeout = Val(pData.ValueOf("timeout"))
''
''        Else
''            .Timeout = -1
''
''        End If
''
''        .Title = pData.ValueOf("title")
''        .Priority = Val(pData.ValueOf("priority"))
''        .DefaultAck = pData.ValueOf("ack")
''        .Value = pData.ValueOf("value")
''
''        If pData.Exists("flags") Then
''            i = Val("&H" & pData.ValueOf("flags"))
''            .Flags = (i And &HFFFF&)                        ' // only keep user flags
''
''        End If
''
''        ' /* these can't be set by external applications - it's a bit klunky at
''        '    present but the notify command handling code in TMainWindow will
''        '    bounce any requests with these tags in them */
''
''        If pData.Exists("remote") Then _
''            .Flags = .Flags Or SNARL42_NOTIFICATION_REMOTE
''
''        If pData.Exists("secure") Then _
''            .Flags = .Flags Or SNARL42_NOTIFICATION_SECURE
''
''
''        .OriginalContent = pData.AsString()
''
''        Set .Socket = ReplySocket
''
''    End With
'
'    g_DoNotify = pApp.Show41(szClass, pData, ReplySocket)
'
'End Function

'Public Function g_DoUpdate41(ByVal Token As Long, ByRef pData As BPackedData) As Long
'
'    ' /* return -1 on success, 0 on failure */
'
'    g_SetLastError SNARL_ERROR_SYSTEM
'    If (g_NotificationRoster Is Nothing) Then _
'        Exit Function
'
'Dim pInfo As notification_info
'
'    If pData.Exists("title") Then
'        pInfo.Title = pData.ValueOf("title")
'
'    Else
'        pInfo.Title = Chr$(255)
'
'    End If
'
'    If pData.Exists("text") Then
'        pInfo.Text = pData.ValueOf("text")
'
'    Else
'        pInfo.Text = Chr$(255)
'
'    End If
'
'    If pData.Exists("icon") Then
'        pInfo.Icon = pData.ValueOf("icon")
'
'    Else
'        pInfo.Icon = Chr$(255)
'
'    End If
'
'Dim hr As M_RESULT
'
'    ' /* call the pre-V41 stuff here - LastError will be set */
'
'    hr = g_NotificationRoster.Update(Token, pInfo.Title, pInfo.Text, pInfo.Icon, pData.AsString)
'    If hr = M_OK Then
'        ' /* success, was timeout specified? */
'
'        If pData.Exists("timeout") Then _
'            g_NotificationRoster.SetAttribute Token, SNARL_ATTRIBUTE_TIMEOUT, pData.ValueOf("timeout")
'
'        g_DoUpdate41 = -1
'
'    Else
'        g_DoUpdate41 = 0
'
'    End If
'
'End Function

'Public Function g_GlobalMessage() As Long
'
'    g_GlobalMessage = RegisterWindowMessage(SNARL_XXX_GLOBAL_MSG)
'
'End Function

Private Function uAddClass(ByVal Token As Long, ByRef Args As BPackedData) As Long
Dim pApp As TApp

    If (Token = 0) And (Not Args.Exists("app-sig")) Then
        g_SetLastError SNARL_ERROR_ARG_MISSING

    ElseIf Token Then
        ' /* FindByToken() will set lasterror for us */
        If g_AppRoster.FindByToken(Token, pApp, Args.ValueOf("password")) Then _
            uAddClass = pApp.AddClass(Args)

    Else
        ' /* FindBySignature() will set lasterror for us */
        If g_AppRoster.FindBySignature(Args.ValueOf("app-sig"), pApp, Args.ValueOf("password")) Then _
            uAddClass = pApp.AddClass(Args)

    End If

End Function

Private Function uRemClass(ByVal Token As Long, ByRef Args As BPackedData, Optional ByVal RemoveAll As Boolean = False) As Long
Dim pApp As TApp

    If Token Then
        ' /* FindByToken() will set lasterror for us */
        If g_AppRoster.FindByToken(Token, pApp, Args.ValueOf("password")) Then _
            uRemClass = pApp.RemClass(Args, RemoveAll)

    Else
        ' /* FindBySignature() will set lasterror for us */
        If g_AppRoster.FindBySignature(Args.ValueOf("app-sig"), pApp, Args.ValueOf("password")) Then _
            uRemClass = pApp.RemClass(Args, RemoveAll)

    End If

End Function

Public Function taglist_as_string(ByRef aList As BTagList) As String

    If (aList Is Nothing) Then _
        Exit Function

Dim pt As BTagItem
Dim sz As String

    With aList
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            sz = sz & pt.Name & "::" & pt.Value & "#?"

        Loop

    End With

    taglist_as_string = g_SafeLeftStr(sz, Len(sz) - 2)

End Function

Public Sub g_SetPresence(ByVal Flags As SN_PRESENCE_FLAGS)
Dim fWasActive As Boolean
Dim fWasAway As Boolean

    fWasAway = g_IsAway()
    fWasActive = (mPresFlags = 0)

    ' /* apply */

    mPresFlags = mPresFlags Or Flags



    ' /* R2.5 Beta 2: update system flags */

Dim dw As SNARL_SYSTEM_FLAGS

    dw = g_GetSystemFlags()

    If g_IsAway() Then _
        dw = dw Or SNARL_SF_USER_AWAY

    If g_IsDND() Then _
        dw = dw Or SNARL_SF_USER_BUSY

    g_SetSystemFlags dw



    ' /* if we've gone from Active to non-Active, log the current missed count */

'    If (fWasActive) And (mPresFlags <> 0) Then _
        g_NotificationRoster.SaveCurrentMissedCount




    ' /* if we've transitioned to Away, notify registered apps */

    If (Not fWasAway) And (g_IsAway()) Then _
        g_AppRoster.SendToAll SNARL_BROADCAST_USER_AWAY


    ' /* R2.4.2 DR3: change the tray icon */
    frmAbout.bSetTrayIcon

End Sub

Public Sub g_ClearPresence(ByVal Flags As SN_PRESENCE_FLAGS)
Dim fWasAwayOrBusy As Boolean
Dim fIsAwayOrBusy As Boolean

Dim fWasActive As Boolean
Dim fWasAway As Boolean

    fWasAway = g_IsAway()
    fWasActive = (mPresFlags = 0)

    fWasAwayOrBusy = (mPresFlags <> 0)

    ' /* apply flags */
    mPresFlags = mPresFlags And (Not Flags)
    fIsAwayOrBusy = (mPresFlags <> 0)



    ' /* R2.5 Beta 2: update system flags */

Dim dw As SNARL_SYSTEM_FLAGS

    dw = g_GetSystemFlags()

    If Not g_IsAway() Then _
        dw = dw And (Not SNARL_SF_USER_AWAY)

    If Not g_IsDND() Then _
        dw = dw And (Not SNARL_SF_USER_BUSY)

    g_SetSystemFlags dw






    ' /* if we've transitioned from Away, notify registered apps */

    If (fWasAway) And (Not g_IsAway()) Then _
        g_AppRoster.SendToAll SNARL_BROADCAST_USER_BACK

    ' /* if we've transitioned to Active, check missed count */

    If (fWasAwayOrBusy) And (mPresFlags = 0) Then _
        g_NotificationRoster.CheckMissed



    ' /* R2.4.2 DR3: change the tray icon */

    If fWasAwayOrBusy <> fIsAwayOrBusy Then _
        frmAbout.bSetTrayIcon

End Sub

Public Function g_IsAway() As Boolean

    g_IsAway = ((mPresFlags And SN_PF_AWAY_MASK) <> 0)

End Function

Public Function g_IsDND() As Boolean

    g_IsDND = ((mPresFlags And SN_PF_DND_MASK) <> 0)

End Function

Public Function g_IsPresence(ByVal Flags As SN_PRESENCE_FLAGS) As Boolean

    g_IsPresence = ((mPresFlags And Flags) <> 0)

End Function

Public Function g_IsUserActive() As Boolean

    g_IsUserActive = (Not g_IsDND()) And (Not g_IsAway())

End Function

Public Function g_GetPresence() As Long

    g_GetPresence = mPresFlags

End Function

Public Function g_GetBase64Icon(ByVal Data As String) As String

    g_GetBase64Icon = uGetEncodedIcon(Replace$(Data, "%", "="))

End Function

Public Function uGetEncodedIcon(ByVal Base64 As String) As String
Dim bErr As Boolean
Dim sz As String

    On Error Resume Next

    sz = Decode64(Base64, bErr)
    If (sz = "") Or (bErr) Then
        g_Debug "uGetEncodedIcon(): failed to decode Base64", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* get a suitably unique path */

    uGetEncodedIcon = g_GetSafeTempIconPath()

Dim i As Integer

    ' /* write the data out */

    i = FreeFile()

    err.Clear
    Open uGetEncodedIcon For Binary Access Write As #i
    If err.Number = 0 Then
        Put #i, , sz
        Close #i

    End If

    g_Debug "uGetEncodedIcon(): writing icon to '" & uGetEncodedIcon & "'"

End Function

Private Function uSetBusy(ByVal Token As Long, ByRef Args As BPackedData) As Long
Dim pApp As TApp

    If Token Then
        ' /* FindByToken() will set lasterror for us */
        If Not g_AppRoster.FindByToken(Token, pApp, Args.ValueOf("password")) Then _
            Exit Function

    Else
        ' /* FindBySignature() will set lasterror for us */
        If Not g_AppRoster.FindBySignature(Args.ValueOf("app-sig"), pApp, Args.ValueOf("password")) Then _
            Exit Function

    End If

    ' /* no app? gah... */

    If (pApp Is Nothing) Then
        g_Debug "uSetBusy(): no returned app object", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

    ' /* TO-DO: allow for user to prevent the app from changing busy mode */

    Select Case g_SafeLong(Args.ValueOf("busy"))
    Case 0
        ' /* reduce count */
        uSetBusyCount False

    Case 1
        ' /* increase count */
        uSetBusyCount True

    Case Else
        ' /* error */
        g_SetLastError SNARL_ERROR_INVALID_ARG

    End Select

End Function

Private Sub uSetBusyCount(ByVal Increment As Boolean)
Static nBusy As Long

    If Increment Then
        g_Debug "uSetBusyCount(): increasing..."
        nBusy = nBusy + 1
        If nBusy = 1 Then _
            g_SetPresence SN_PF_DND_EXTERNAL

    Else
        g_Debug "uSetBusyCount(): decreasing..."
        nBusy = nBusy - 1
        If nBusy = 0 Then _
            g_ClearPresence SN_PF_DND_EXTERNAL

    End If

End Sub

Public Function g_When(ByVal Timestamp As Date) As String
Dim i As Long

    On Error GoTo fail

    ' /* default response is "<date> at <time>" */

    g_When = Format$(Timestamp, "d mmm yyyy") & " at " & Format$(Timestamp, "ttttt")

    ' /* if more than a day ago, default is enough */

    If DateDiff("d", Timestamp, Now) > 0 Then _
        Exit Function

    ' /* if an hour or more ago, use hours */

    i = DateDiff("n", Timestamp, Now)
    If i > 59 Then
        i = Fix(i / 60)
        g_When = CStr(i) & " hour" & IIf(i = 1, "", "s") & " ago"
        Exit Function

    End If

    ' /* if a minute or more ago. use minutes */

    i = DateDiff("n", Timestamp, Now)
    If i > 0 Then
        g_When = CStr(i) & " min" & IIf(i = 1, "", "s") & " ago"

    Else
        g_When = "Just now"

    End If

fail:

End Function

Public Sub g_LoadIconTheme()
Dim sz As String

    sz = g_ConfigGet("icon_theme")

    If Not g_Exists(g_MakePath(App.Path) & "etc\icons\" & sz & ".icons") Then
        g_Debug "g_LoadIconTheme(): theme '" & sz & "' not found", LEMON_LEVEL_WARNING
        sz = ""
        g_ConfigSet "icon_theme", ""

    Else
        sz = sz & ".icons\"

    End If

    sz = g_MakePath(App.Path) & "etc\icons\" & sz

    uSafeLoadImage sz, "widget-close.png", bm_CloseGadget
    uSafeLoadImage sz, "widget-actions.png", bm_ActionsGadget
    uSafeLoadImage sz, "emblem-actions.png", bm_HasActions
    uSafeLoadImage sz, "emblem-remote.png", bm_Remote
    uSafeLoadImage sz, "emblem-secure.png", bm_Secure
    uSafeLoadImage sz, "emblem-sticky.png", bm_IsSticky
    uSafeLoadImage sz, "emblem-priority.png", bm_Priority
    uSafeLoadImage sz, "emblem-forward.png", bm_Forward

    Set bm_CallbackButton = g_CreateButton(new_BPoint(66, 24))
    Set bm_Button = g_CreateButton(new_BPoint(24, 24))

'    load_image sz & "menu.png", bm_Menu                 ' // no longer used

'    If Not g_IsValidImage(bm_CloseGadget) Then _
        Set bm_CloseGadget = g_CreateBadge("X")

End Sub

Private Sub uSafeLoadImage(ByVal Path As String, ByVal Name As String, ByRef Obj As mfxBitmap)

    Set Obj = load_image_obj(Path & Name)
    If Not is_valid_image(Obj) Then _
        Set Obj = load_image_obj(g_MakePath(App.Path) & "etc\icons\" & Name)

End Sub

Public Sub g_RunFileLoadSchemes()
Dim pt As BTagItem
Dim sn As String
Dim sd As String
Dim prf As TRunFileScheme
Dim pf As CConfFile

    Set gRunFiles = new_BTagList()

    ' /* load up new format entries */

    sn = style_GetSnarlStylesPath()
    If sn <> "" Then
        sn = g_MakePath(sn) & "runfile"
        With New CFolderContent2
            If .SetTo(sn) Then
                .Rewind
                Do While .GetNextFile(sd)
                    Set prf = New TRunFileScheme
                    If prf.SetTo(sd) Then _
                        gRunFiles.Add prf

                Loop

            Else
                g_Debug "g_RunFileLoadSchemes(): %appdata%\styles\runfile\ missing", LEMON_LEVEL_CRITICAL

            End If

        End With

    Else
        g_Debug "g_RunFileLoadSchemes(): %appdata%\styles\ missing", LEMON_LEVEL_CRITICAL

    End If

    ' /* does the V1 config exist? */

    With New CConfFile
        If .SetTo(g_MakePath(App.Path) & "etc\runfile.conf", True) Then

            g_Debug "g_RunFileLoadSchemes(): got V1 runfile config"

            .Rewind
            Do While .GetEntry(sn, sd)
                If sn = "target" Then _
                    gRunFiles.Add new_BTagItem(sd, "1")

            Loop

            g_Debug "g_RunFileLoadSchemes(): converting entries..."

            ' /* convert to new format */

            sn = style_GetSnarlStylesPath()
            If sn <> "" Then
                sn = g_MakePath(sn) & "runfile"
                If g_Exists(sn) Then
                    With gRunFiles
                        .Rewind
                        Do While .GetNextTag(pt) = B_OK
                            Set pf = New CConfFile
                            pf.SetTo g_MakePath(sn) & g_CreateGUID(True) & ".runfile"
                            pf.Add "target=" & pt.Name
                            pf.Add "version=1"
                            pf.Save

                        Loop

                    End With

                    ' /* delete old V1 config */

                    g_Debug "g_RunFileLoadSchemes(): removing V1 runfile config..."
                    DeleteFile g_MakePath(App.Path) & "etc\runfile.conf"

                Else
                    g_Debug "g_RunFileLoadSchemes(): %appdata%\styles\runfile\ missing", LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "g_RunFileLoadSchemes(): %appdata%\styles\ missing", LEMON_LEVEL_CRITICAL

            End If

        End If

    End With

End Sub

Private Function uIsDebugBuild() As Boolean
Dim sz As String

    sz = LCase$(App.Comments)
    uIsDebugBuild = (InStr(sz, "debug") <> 0) 'Or (InStr(sz, "dr") <> 0) Or (InStr(sz, "alpha") <> 0)

End Function

Public Function g_DoV42Request(ByVal Request As String, ByVal SenderPID As Long, Optional ByRef ReplySocket As CSocket, Optional ByVal Flags As SN_NOTIFICATION_FLAGS) As Long

    g_Debug "g_DoV42Request(): '" & Request & "'"

    ' /* must at least have an action */

    If Request = "" Then
        g_SetLastError SNARL_ERROR_BAD_PACKET
        Exit Function

    End If

    ' /* R2.4.2 DR3: if a reply socket is provided, set SN_NF_REMOTE accordingly */

    If Not (ReplySocket Is Nothing) Then
        g_Debug "g_DoV42Request(): sender is " & ReplySocket.RemoteHost & ":" & ReplySocket.RemotePort
        If ReplySocket.RemoteHostIP <> "127.0.0.1" Then _
             Flags = Flags Or SN_NF_REMOTE

    End If

    ' /* tokenise special character pairs */

    Request = Replace$(Request, "&&", "%26")
    Request = Replace$(Request, "==", "%3d")

    ' /* convert to internal format.  Real '=' and '&' will still be url-encoded so they're safe */

    Request = Replace$(Request, "=", "::")
    Request = Replace$(Request, "&", "#?")

Dim szAction As String
Dim i As Long

    ' /* find the action/arg marker */

    i = InStr(Request, "?")
    If i Then
        ' /* action and args */
        szAction = g_SafeLeftStr(Request, i - 1)
        Request = g_SafeRightStr(Request, Len(Request) - i)

    Else
        ' /* just an action */
        szAction = Request
        Request = ""

    End If

    ' /* set the packed data and URL-decode the arguments at the same time */

Dim pData As BPackedData
Dim lToken As Long

    Set pData = New BPackedData
    pData.SetTo g_URLDecode(Request)
    lToken = g_SafeLong(pData.ValueOf("token"))

    If (Not (ReplySocket Is Nothing)) And (pData.Exists("app-sig")) Then _
        pData.Update "app-sig", pData.ValueOf("app-sig") & "#" & ReplySocket.RemoteHost

    ' /* pass to master function */

    g_DoV42Request = g_DoAction(szAction, lToken, pData, App.Major Or Flags, ReplySocket, SenderPID)  ' // R2.4.1: include major API version in this

    If g_DoV42Request = 0 Then
        g_DoV42Request = -g_QuickLastError()

    ElseIf g_DoV42Request = -1 Then
        g_DoV42Request = 0

    End If

End Function

Public Function g_GetPhat64Icon(ByVal Data As String) As String

    g_GetPhat64Icon = uGetEncodedIcon(Replace$(Replace$(Data, "#", vbCrLf), "%", "="))

End Function

'Public Function g_SetPassword(ByVal Password As String, ByVal AuthType As String) As Boolean
'
'    If (Password <> "") And (AuthType <> "") Then
'
'        ' /* create the salt */
'
'Dim szSalt As String
'Dim b As Integer
'Dim i As Integer
'
'        For i = 1 To 16
'            Randomize Timer
'            b = (Rnd * 254) + 1
'            szSalt = szSalt & g_HexStr(b, 2)
'
'        Next i
'
'        Password = Password & szSalt
'
'Dim szKey As String
'
'        Select Case AuthType
'        Case "sha1"
'            With New sha1
'                szKey = .sha1(Password)
'
'            End With
'
'        Case "sha256"
'            With New SHA256
'                szKey = .SHA256(Password)
'
'            End With
'
'        Case "md5"
'            With New MD5
'                szKey = .DigestStrToHexStr(Password)
'
'            End With
'
'        Case Else
'            g_Debug "g_SetPassword(): invalid algorithm", LEMON_LEVEL_CRITICAL
'            Exit Function
'
'        End Select
'
'        g_Debug "g_SetPassword(): " & AuthType & " " & szKey & " " & szSalt
'
'
'    ElseIf (Password = "") And (AuthType = "") Then
'        ' /* special case: clear password */
'        g_Debug "g_SetPassword(): clearing password", LEMON_LEVEL_INFO
'
'
'    Else
'        g_Debug "g_SetPassword(): missing arg", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End If
'
'    g_ConfigSet "auth_type", AuthType
'    g_ConfigSet "auth_salt", szSalt
'    g_ConfigSet "auth_key", szKey
'    g_WriteConfig
'
'    g_SetPassword = True
'
'End Function

Public Function g_SetPassword(ByVal Password As String) As Boolean
Dim sz As String

    If Password <> "" Then
        If Not EncodePlus(Password, sz) Then _
            Exit Function

    End If

    g_ConfigSet "auth_password", sz
    g_WriteConfig
    g_SetPassword = True

End Function

Public Function g_GetPassword() As String
Dim szPwd As String
Dim sz As String

    sz = g_ConfigGet("auth_password")
    If sz <> "" Then
        If DecodePlus(sz, szPwd) Then
            g_GetPassword = szPwd

        Else
            SOS_invoke New TSOSHandler

        End If

    End If

End Function

Public Sub g_SNP3SendCallback(ByRef Socket As CSocket, ByVal StatusCode As SNARL_STATUS_CODE, ByVal EventName As String, ByVal ResponseDetails As String, ByVal Content As String)

    If (Socket Is Nothing) Then _
        Exit Sub

Dim sz As String

    ' /* standard headers */

    sz = sz & "x-timestamp: " & Format$(Now(), "d mmm yyyy hh:mm:ss") & vbCrLf
    sz = sz & "x-daemon: " & "Snarl " & CStr(APP_VER) & "." & CStr(APP_SUB_VER) & IIf(APP_SUB_SUB_VER <> 0, "." & CStr(APP_SUB_SUB_VER), "") & vbCrLf
    sz = sz & "x-host: " & LCase$(g_GetComputerName()) & vbCrLf

    Socket.SendData "SNP/3.0 CALLBACK" & vbCrLf & _
                    "event-code: " & CStr(StatusCode) & vbCrLf & _
                    "event-name: " & EventName & vbCrLf & _
                    IIf(ResponseDetails <> "", ResponseDetails, "") & _
                    Content & _
                    sz & _
                    "END" & vbCrLf

End Sub

Public Sub g_Trap(ByVal Error As SOS_ERRORS, ByVal Data As String)
Dim sz As String

    sz = "**** Snarl System Error Nr $" & g_HexStr(Error) & "!" & vbCrLf

    Select Case Error
    Case SOS_BAD_COPYDATA
        sz = sz & "Bad WM_COPYDATA id $" & Data

    Case SOS_SPURIOUS_TEST
        sz = sz & "Bad WM_SNARL_TEST id $" & Data

    Case SOS_FILE_NOT_FOUND
        sz = sz & "File >" & Data & "< not found"

    Case SOS_PATH_NOT_FOUND
        sz = sz & "Path >" & Data & "< not found"

    Case SOS_MISSING_ROSTER
        sz = sz & ">" & Data & "_roster< not found"
    
    Case Else
        sz = sz & "Undefined error"

    End Select

    SOS_invoke New TSOSHandler, sz & vbCrLf

End Sub

Public Function new_BPackedData(ByVal Content As String) As BPackedData

    Set new_BPackedData = New BPackedData
    new_BPackedData.SetTo Content

End Function

Public Function g_CopyToAppData(ByVal Source As String, ByVal DestPath As String) As Boolean
Dim sz As String

    If g_GetUserFolderPath(sz) Then
        sz = g_MakePath(sz) & g_MakePath(DestPath) & g_FilenameFromPath(Source)
        If g_Exists(sz) Then
            If MsgBox("'" & g_FilenameFromPath(Source) & "' already exists.  Do you want to replace the existing file?", vbYesNo Or vbQuestion, App.Title) = vbNo Then _
                Exit Function

        End If
    
        g_CopyToAppData = (CopyFile(Source, sz, 0) <> 0)

            
'            MsgBox "Web Forwarder '" & g_FilenameFromPath(Source) & "' was installed successfully", vbInformation Or vbOKOnly, App.Title
    
    End If

End Function

Public Function g_ExtractToAppData(ByVal Source As String, ByVal DestPath As String) As Boolean
Dim sz As String

    If g_GetUserFolderPath(sz) Then
        sz = g_MakePath(sz) & g_MakePath(DestPath)
        With New CZippedContent
            If .OpenZip(Source) Then _
                g_ExtractToAppData = .Extract(sz, True, True)

        End With
    End If

End Function

Public Sub g_SetSystemFlags(ByVal SystemFlags As SNARL_SYSTEM_FLAGS)

    mFlags = SystemFlags
    SetProp ghWndMain, "_flags", SystemFlags

End Sub

Public Function g_GetSystemFlags() As SNARL_SYSTEM_FLAGS

    g_GetSystemFlags = mFlags

End Function

Public Function g_IsLocalAddress(ByVal IPAddress As String, Optional ByVal IgnoreDebugMode As Boolean = False) As Boolean

    If (Not IgnoreDebugMode) And (gDebugMode) Then _
        Exit Function

    IPAddress = LCase$(IPAddress)
    If (IPAddress = "localhost") Or (IPAddress = LCase$(g_GetComputerName())) Then
        g_IsLocalAddress = True

    Else
        g_IsLocalAddress = (InStr(get_ip_address_table(), IPAddress) <> 0)

    End If

End Function

Private Sub uCreateUserSettings()
Dim sz As String

    If Not g_GetSystemFolder(CSIDL_APPDATA, sz) Then _
        Exit Sub

    sz = g_MakePath(sz)
    g_CreateDirectory sz & "full phat"

    sz = sz & "full phat\snarl"
    g_CreateDirectory sz

    If Not g_Exists(sz & "\etc") Then _
        g_CreateDirectory sz & "\etc"

    If Not g_Exists(sz & "\etc\app-cache") Then _
        g_CreateDirectory sz & "\etc\app-cache"

    If Not g_Exists(sz & "\styles") Then _
        g_CreateDirectory sz & "\styles"

    If Not g_Exists(sz & "\styles\runfile") Then _
        g_CreateDirectory sz & "\styles\runfile"

    If Not g_Exists(sz & "\startup-sequence") Then _
        g_CreateDirectory sz & "\startup-sequence"

'        uCreateLink sz & "\extensions\banner.styleengine"
'        uCreateLink sz & "\extensions\ez.styleengine"
'        uCreateLink sz & "\extensions\meter.styleengine"
'        uCreateLink sz & "\extensions\mobile.styleengine"
'        uCreateLink sz & "\extensions\prowl.styleengine"
'        uCreateLink sz & "\extensions\runnable.styleengine"
'        uCreateLink sz & "\extensions\sapi.styleengine"
'
'    End If

    If Not g_Exists(sz & "\extensions") Then _
        g_CreateDirectory sz & "\extensions"
'        uCreateLink sz & "\extensions\AudioMon.extension"
'        uCreateLink sz & "\extensions\SnarlClock2.extension"
'        uCreateLink sz & "\extensions\SNPHTTP.extension"
'        uCreateLink sz & "\extensions\SysInfo.extension"
'        uCreateLink sz & "\extensions\TMinus.extension"
'        uCreateLink sz & "\extensions\WLANMonitor.extension"
'
'    End If

End Sub

Private Sub uCreateLink(ByVal Path As String)
Dim i As Integer

    On Error Resume Next

    i = FreeFile()
    
    Open Path For Output As #i
    Close #i

End Sub

Private Sub uRunStartupSequence(ByVal Path As String)

    If gSysAdmin.NoRunStartupSequence Then
        g_Debug "uRunStartupSequence(): blocked by administrator", LEMON_LEVEL_INFO
        Exit Sub

    End If

Dim pcf As CConfFile
Dim szTarget As String
Dim sz As String
Dim n As Long

    With New CFolderContent2
        If .SetTo(g_MakePath(Path) & "startup-sequence") Then
            .Rewind
            Do While .GetNextFile(sz, True)
                If g_GetExtension(sz) = "ssl2" Then
                    Set pcf = New CConfFile
                    pcf.SetFilename sz
                    pcf.Reload
                    
                    szTarget = pcf.ValueOf("target")
                    If szTarget <> "" Then

                        n = SW_SHOW

                        Select Case pcf.ValueOf("show")

                        Case "minimised", "minimized"
                            n = SW_MINIMIZE

                        Case "maximised", "maximized", "zoomed"
                            n = SW_MAXIMIZE

'                        Case "hidden"
'                            n = SW_HIDE

                        End Select

                        ShellExecute frmAbout.hWnd, "open", szTarget, pcf.ValueOf("args"), pcf.ValueOf("pwd"), n

                    Else
                        g_Debug "uRunStartupSequence(): '" & sz & "' has no target"

                    End If
                
                Else
                    Debug.Print "'" & sz & "' is not an ssl2"

                End If
            Loop
        End If

    End With

End Sub

Private Sub uDoSystemRequest(ByRef pArgs As BPackedData)

    ' /* syntax is "snarl?cmd=" - sets lasterror on exit */

    g_Debug "uDoSystemRequest()", LEMON_LEVEL_PROC_ENTER
    g_SetLastError SNARL_SUCCESS

    ' /* must have at least an "cmd" argument */

'    MsgBox "request: " & pArgs.AsString

Dim hr As Long

    If Not pArgs.Exists("cmd") Then
        g_Debug "command missing", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Sub

    End If

Dim szData As String
Dim szCmd As String

    ' /* get the command - defined commands thus far:
    '
    '       cmd
    '   ----------------------------------------------------
    '       load            extension|styleengine
    '       unload          extension|styleengine
    '       reload          extension|styleengine
    '       configure       extension|styleengine|application|style
    '       reload_styles
    '       reload_extensions
    '
    '
    ' */

    szCmd = LCase$(pArgs.ValueOf("cmd"))

    Select Case szCmd
    Case "reload_styles"
        melonLibClose g_StyleRoster
        melonLibOpen g_StyleRoster
        g_SetLastError SNARL_SUCCESS

    Case "reload_extensions"
        melonLibClose g_ExtnRoster
        melonLibOpen g_ExtnRoster
        g_SetLastError SNARL_SUCCESS

    Case "load", "unload", "reload"
        ' /* can be an extension or style engine */
        If pArgs.Exists("what") Then
            szData = pArgs.ValueOf("what")
            Select Case g_GetExtension(szData, True)
            Case "extension"
                g_SetLastError uManageExtension(szCmd, g_RemoveExtension(szData))
    
            Case "styleengine"
                g_SetLastError uManageStyleEngine(szCmd, szData)
    
            Case Else
                g_Debug g_Quote(szData) & ": not supported", LEMON_LEVEL_CRITICAL
                g_SetLastError SNARL_ERROR_INVALID_ARG
    
            End Select

        Else
            g_Debug "argument missing", LEMON_LEVEL_CRITICAL
            g_SetLastError SNARL_ERROR_ARG_MISSING

        End If


    Case "configure"
        ' /* can be a style, styleengine, extension or application */
        If pArgs.Exists("what") Then
            szData = pArgs.ValueOf("what")
            Select Case g_GetExtension(szData, True)
            Case "extension"
                g_SetLastError uManageExtension(szCmd, g_RemoveExtension(szData))
    
            Case "styleengine"
                g_SetLastError uManageStyleEngine(szCmd, g_RemoveExtension(szData))
    
            Case "style"
            
            
            Case Else
                ' /* assume its an application signature */
                g_SetLastError uConfigureApp(szData)
    
            End Select

        Else
            g_Debug "argument missing", LEMON_LEVEL_CRITICAL
            g_SetLastError SNARL_ERROR_ARG_MISSING

        End If

    Case "reboot"
        ' /* ends Snarl, runs DelayLoad.exe */



    Case Else
        g_Debug "invalid comand " & g_Quote(szCmd), LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_INVALID_ARG

    End Select

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Function uConfigureApp(ByVal Signature As String) As SNARL_STATUS_CODE
Dim pa As TApp

    If g_AppRoster.FindBySignature(Signature, pa, "") Then
        If pa.HasConfig Then
            g_Debug "uConfigureApp(): asking " & g_Quote(pa.Name) & " to show its config..."
            pa.DoSettings 0

        Else
            g_Debug "uConfigureApp(): " & g_Quote(pa.Name) & " is not configurable", LEMON_LEVEL_CRITICAL
            uConfigureApp = SNARL_ERROR_ACCESS_DENIED

        End If

    Else
        g_Debug "uConfigureApp(): " & g_Quote(Signature) & " not found/bad password", LEMON_LEVEL_CRITICAL
        uConfigureApp = g_QuickLastError()

    End If

End Function

Private Function uManageExtension(ByVal Command As String, ByVal Name As String) As SNARL_STATUS_CODE

    g_Debug "uManageExtension()", LEMON_LEVEL_PROC_ENTER

    If (g_ExtnRoster Is Nothing) Then
        g_Debug "fatal: extension roster not started", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_Trap SOS_MISSING_ROSTER, "extension"
        uManageExtension = SNARL_ERROR_SYSTEM
        Exit Function

    End If

Dim pe As TExtension

    Select Case Command

    Case "unload"
        uManageExtension = g_ExtnRoster.Unload(Name, True)

    Case "load"
        uManageExtension = g_ExtnRoster.Load(Name, True)

    Case "reload"
        uManageExtension = g_ExtnRoster.Reload(Name, True)

    Case "configure"
        uManageExtension = g_ExtnRoster.Configure(Name)


'        Case SN_DP_RESTART
'            If g_ExtnRoster.Find(Item, pe) Then
'                pe.SetEnabled False
'                pe.SetEnabled True
'
'            End If

    Case "install"
        ' /* 1. create link file - doesn't matter if it exists? - use "allusers=1"? */
        ' /* 2. refresh extensions */
        ' /* 3. load it */
        ' /* 4. show config? */
        MsgBox "installing extensions is not yet implemented"

    Case "uninstall"
        ' /* 1. unload
        '    2. delete link file
        ' */
        MsgBox "uninstalling extensions is not yet implemented"


    Case Else
        g_Debug g_Quote(Command) & ": unknown command", LEMON_LEVEL_CRITICAL
        uManageExtension = SNARL_ERROR_INVALID_ARG

    End Select

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function uManageStyleEngine(ByVal Command As String, ByVal Name As String) As SNARL_STATUS_CODE

    g_Debug "uManageStyleEngine()", LEMON_LEVEL_PROC_ENTER

    If (g_ExtnRoster Is Nothing) Then
        g_Debug "fatal: style roster not started", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_Trap SOS_MISSING_ROSTER, "style"
        uManageStyleEngine = SNARL_ERROR_SYSTEM
        Exit Function

    End If

Dim pe As TStyleEngine

    Select Case Command

    Case "install"
        ' /* 1. create link file - doesn't matter if it exists? - use "allusers=1"? */
        ' /* 2. refresh extensions */
        ' /* 3. load it */
        ' /* 4. show config? */
        MsgBox "installing style engines is not yet implemented"

    Case "uninstall"
        ' /* 1. unload
        '    2. delete link file
        ' */
        MsgBox "uninstalling style engines is not yet implemented"

'        Case SN_DP_RESTART
'            g_StyleRoster.Unload Item, True
'            g_StyleRoster.Load Item, True, True
'
'        Case SN_DP_UNLOAD
'            g_StyleRoster.Unload Item, True
'
'        Case SN_DP_LOAD
'            g_StyleRoster.Load Item, True, True
'

    Case "unload"
        g_Debug "unloading " & g_Quote(Name) & "..."
        If g_StyleRoster.Unload(Name, False) Then
            g_Debug "ok"

        Else
            uManageStyleEngine = SNARL_ERROR_FAILED
            g_Debug "failed", LEMON_LEVEL_CRITICAL

        End If

    Case "load"
        g_Debug "loading " & g_Quote(Name) & "..."
        If g_StyleRoster.Load(Name, True, False) Then
            g_Debug "ok"

        Else
            uManageStyleEngine = SNARL_ERROR_FAILED
            g_Debug "failed", LEMON_LEVEL_CRITICAL

        End If

    Case "reload"
        ' /* just do a stop/start
        g_Debug "unloading " & g_Quote(Name) & "..."
        g_StyleRoster.Unload Name, False
        g_Debug "loading " & g_Quote(Name) & "..."
        g_StyleRoster.Load Name, True, False


    Case "configure"
        g_Debug "finding " & g_Quote(Name) & "..."
        If g_StyleRoster.FindEngine(Name, pe) Then
            g_Debug "launching config for " & g_Quote(Name) & "..."
            If pe.IsConfigurable Then
                pe.Configure

            Else
                g_Debug "not configurable", LEMON_LEVEL_CRITICAL
                uManageStyleEngine = SNARL_ERROR_FAILED

            End If
        Else
            g_Debug "not found", LEMON_LEVEL_CRITICAL
            uManageStyleEngine = SNARL_ERROR_ADDON_NOT_FOUND
        
        End If

    Case Else
        g_Debug g_Quote(Command) & ": unknown command", LEMON_LEVEL_CRITICAL
        uManageStyleEngine = SNARL_ERROR_INVALID_ARG

    End Select

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Sub g_AddGlobalRedirect(ByVal StyleAndScheme As String, ByVal Flags As SN_REDIRECTION_FLAGS)

    gGlobalRedirectList.Add new_BTagItem(StyleAndScheme, CStr(Flags))
    g_ConfigSet "global_redirect", taglist_as_string(gGlobalRedirectList)

End Sub

Public Sub g_RemGlobalRedirect(ByVal StyleAndScheme As String)

    gGlobalRedirectList.Remove gGlobalRedirectList.IndexOf(StyleAndScheme)
    g_ConfigSet "global_redirect", taglist_as_string(gGlobalRedirectList)

End Sub

Public Sub g_UpdateRedirectList(ByRef ListControl As BControl, ByRef List As BTagList, ByVal LargeIcons As Boolean)

    If (ISNULL(ListControl)) Or (ISNULL(List)) Then _
        Exit Sub

Dim pt As BTagItem
Dim ps As TStyle
Dim sz As String
Dim szr As String
Dim f As SN_REDIRECTION_FLAGS
Dim bWhen As Boolean
Dim i As Long

    ' /* set content */

    With List
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            bWhen = True

            Select Case g_SafeLong(pt.Value)
            Case SN_RF_ALWAYS
                szr = "Always"
                bWhen = False

            Case SN_RF_WHEN_ACTIVE
                szr = "active"
                
            Case SN_RF_WHEN_AWAY
                szr = "away"
                
            Case SN_RF_WHEN_BUSY
                szr = "busy"

            Case SN_RF_WHEN_ACTIVE Or SN_RF_WHEN_AWAY
                szr = "active or away"

            Case SN_RF_WHEN_ACTIVE Or SN_RF_WHEN_BUSY
                szr = "active or busy"

            Case SN_RF_WHEN_AWAY Or SN_RF_WHEN_BUSY
                szr = "away or busy"

            Case SN_RF_NEVER
                szr = "Never"
                bWhen = False

            End Select

            szr = IIf(LargeIcons, "", "(") & IIf(bWhen, "When ", "") & szr & IIf(LargeIcons, "", ")")
            sz = sz & style_MakeFriendly(pt.Name) & IIf(LargeIcons, "", " " & szr) & "#?" & pt.Name & IIf(LargeIcons, "#?" & szr, "") & "|"
        
        Loop
    
        sz = g_SafeLeftStr(sz, Len(sz) - 1)
        ListControl.SetText sz

    End With

    ' /* set icons */

    If sz <> "" Then
        With List
            .Rewind
            Do While .GetNextTag(pt) = B_OK
                i = i + 1
                If g_StyleRoster.Find(style_GetStyleName(pt.Name), ps) Then
                    If Not g_Exists(ps.IconPath) Then
                        prefskit_SetItemObject ListControl, i, "image-object", load_image_obj(g_MakePath(App.Path) & "etc\icons\class-fwd.png")

                    Else
                        prefskit_SetItemObject ListControl, i, "image-object", ps.Icon

                    End If

                Else
                    prefskit_SetItemObject ListControl, i, "image-object", load_image_obj(g_MakePath(App.Path) & "etc\icons\no_icon.png")

                End If
            Loop
        End With
    End If

End Sub

Public Function g_ButtonLabelFromAck(ByVal Ack As String) As String

    If g_IsURL(Ack) Then
        g_ButtonLabelFromAck = "Go"

    ElseIf g_SafeLeftStr(Ack, 1) = "@" Then
        g_ButtonLabelFromAck = "Show"

    ElseIf g_SafeLeftStr(Ack, 1) = "!" Then
        g_ButtonLabelFromAck = uBangLabel(g_SafeRightStr(Ack, Len(Ack) - 1))

    Else
'    ElseIf (g_SafeLeftStr(Ack, 1) = "!") Or (g_IsFileURI(Ack)) Or (g_Exists(Ack)) Then
        g_ButtonLabelFromAck = "Open"

    End If

End Function

Private Function uBangLabel(ByVal Bang As String)
Dim i As Long

    i = InStr(Bang, " ")
    If i Then _
        Bang = g_SafeLeftStr(Bang, i - 1)

    Select Case Bang
    Case "missed"
        uBangLabel = "View"

    Case "test"
        uBangLabel = "Okay"

    Case Else
        uBangLabel = "Open"

    End Select

End Function

Public Function g_CreateButton(ByRef Size As BPoint) As mfxBitmap
Const RX = 6
Dim pr As BRect

    With New mfxView
        .SizeTo Size.x, Size.y
        .EnableSmoothing True

'        .SetHighColour rgba(0, 0, 0, 100)
'        .FillRoundRect .Bounds, RX, RX
'        .StrokeRoundRect .Bounds.InsetByCopy(2, 2), RX, RX
'
''        .SetHighColour rgba(0, 0, 0, 0)
''        .SetLowColour rgba(0, 0, 0, 48)
''        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT
'        .SetHighColour rgba(255, 255, 255, 150)
'        .StrokeRoundRect .Bounds, RX, RX, 2
'
'        .TextMode = MFX_TEXT_ANTIALIAS
'        If Label = "*" Then
'            .DrawScaledImage bm_HasActions, new_BPoint(Fix((.Width - bm_HasActions.Width) / 2), Fix((.Height - bm_HasActions.Height) / 2))
'
'        Else
'            .SetFont "Arial", 8, True
'            .SetHighColour rgba(255, 255, 255)
'            .DrawString Label, .Bounds, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER ' Or MFX_SIMPLE_OUTLINE
'
'        End If

'        .SetHighColour rgba(0, 0, 0, 110)
'        .FillRoundRect .Bounds, RX, RX
'
'        .SetLowColour rgba(255, 255, 255, 120)
'        .SetHighColour rgba(0, 0, 0, 0)
'        .StrokeFancyRoundRect .Bounds, RX, RX
'
'        .SetHighColour rgba(240, 240, 240)
'        .FillRoundRect .Bounds.InsetByCopy(1, 1), RX, RX
'
''        .SetHighColour rgba(255, 255, 255, 32)
''        .SetLowColour rgba(255, 255, 255, 0)
''        .FillRoundRect .Bounds.InsetByCopy(1, 1), RX, RX, MFX_VERT_GRADIENT
'
'        .SetHighColour rgba(0, 0, 0, 0)
'        .SetLowColour rgba(0, 0, 0, 48)
'        .FillRoundRect .Bounds.InsetByCopy(1, 1), RX, RX, MFX_VERT_GRADIENT
'        .SetHighColour rgba(0, 0, 0, 100)
'        .StrokeRoundRect .Bounds.InsetByCopy(1, 1), RX, RX




'        .SetHighColour rgba(255, 255, 255, 0)
'        .SetLowColour rgba(255, 255, 255, 110)
'        .FillRoundRect .Bounds.InsetByCopy(0, 1).OffsetByCopy(0, 1), RX, RX, MFX_VERT_GRADIENT
'
'        Set pr = .Bounds.InsetByCopy(1, 1)
'        .SetHighColour rgba(240, 240, 240)
'        .FillRoundRect pr, RX, RX
''        pr.Top = Fix(.Bounds.Height / 2)
'        .SetHighColour rgba(0, 0, 0, 0)
'        .SetLowColour rgba(0, 0, 0, 48)
'        .FillRoundRect pr, RX, RX, MFX_VERT_GRADIENT
'        .SetHighColour rgba(0, 0, 0, 90)
'        .StrokeRoundRect .Bounds.InsetByCopy(1, 1), RX, RX



        .SetHighColour rgba(230, 230, 230)
        .FillRoundRect .Bounds, RX, RX
        .SetHighColour rgba(0, 0, 0, 150)
        .StrokeRoundRect .Bounds.InsetByCopy(0, 0), RX, RX

        .SetHighColour rgba(255, 255, 255, 170)
        .SetLowColour rgba(0, 0, 0, 2)
        .StrokeFancyRoundRect .Bounds.InsetByCopy(1, 1), RX, RX

        Set g_CreateButton = .ConvertToBitmap()

    End With

End Function

Public Function g_InstallRSZ(ByVal Path As String, ByVal FromRunningInstance As Boolean) As Boolean
Dim pzc As CZippedContent
Dim pse As TStyleEngine
Dim szAuthor As String
Dim szName As String
Dim sz As String
Dim i As Long

    ' /* packed runnable style */
    Set pzc = New CZippedContent
    With pzc
        If .OpenZip(Path) Then
            If (.ContainsFile("runnable.conf")) And (.ContainsFile("style.exe")) Then
                sz = g_GetTempPath(True)
                sz = sz & g_CreateGUID(True)
                If g_CreateDirectory(sz) Then
                    .Extract sz, False, True
                    With New ConfigFile
                        .File = g_MakePath(sz) & "runnable.conf"
                        .Load
                        ' /* must have a [general] section */
                        If .SectionExists("general") Then
                            With .SectionAt(.FindSection("general"))
                                ' /* must have a name */
                                If .Find("name", szName) Then
                                    .Find "author", szAuthor

                                    If MsgBox("Do you want to install the following Runnable style?" & vbCrLf & vbCrLf & _
                                              "Name: " & szName & vbCrLf & _
                                              IIf(szAuthor <> "", "Author: " & szAuthor & vbCrLf, ""), vbQuestion Or vbYesNo, App.Title) = vbYes Then

                                        If g_ExtractToAppData(Path, "styles\runnable") Then
                                            If FromRunningInstance Then
                                                g_PrivateNotify , "Style installed", szName & " Runnable style was installed successfully", , ".good"
                                                g_StyleRoster.Unload "runnable.styleengine", False
                                                g_StyleRoster.Load "runnable.styleengine", False, False
                                                frmAbout.bNotifyDisplaysChanged

                                            Else
                                                MsgBox szName & " was installed successfully!", vbInformation Or vbOKOnly, App.Title

                                            End If
                                            g_InstallRSZ = True

                                        Else
                                            MsgBox "There was a problem installing the style.", vbExclamation Or vbOKOnly, App.Title
                                        
                                        End If
                                    End If
                                    Exit Function

                                Else
                                    g_Debug "g_InstallRSZ(): missing style name", LEMON_LEVEL_CRITICAL

                                End If
                            End With
                        Else
                            g_Debug "g_InstallRSZ(): missing [general] section", LEMON_LEVEL_CRITICAL

                        End If
                    End With
                Else
                    g_Debug "g_InstallRSZ(): failed to create temp folder", LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "g_InstallRSZ(): missing conf or exe", LEMON_LEVEL_CRITICAL

            End If
        Else
            g_Debug "g_InstallRSZ(): failed to open zip file", LEMON_LEVEL_CRITICAL
        
        End If
    End With

    MsgBox Path & " appears to be corrupt.  Try downloading it again.", vbExclamation Or vbOKOnly, App.Title

End Function

Public Function g_InstallSSZ(ByVal Path As String, ByVal FromRunningInstance As Boolean) As Boolean
Dim pzc As CZippedContent
Dim pse As TStyleEngine
Dim szAuthor As String
Dim szName As String
Dim sz As String
Dim i As Long

    ' /* packed runnable style */
    Set pzc = New CZippedContent
    With pzc
        If .OpenZip(Path) Then
            If (.ContainsFile("script.vbs")) And (.ContainsFile("script.conf")) Then
                sz = g_GetTempPath(True)
                sz = sz & g_CreateGUID(True)
                If g_CreateDirectory(sz) Then
                    .Extract sz, False, True
                    With New ConfigFile
                        .File = g_MakePath(sz) & "script.conf"
                        .Load
                        ' /* must have a [general] section */
                        If .SectionExists("general") Then
                            With .SectionAt(.FindSection("general"))
                                ' /* must have a name */
                                If .Find("name", szName) Then
                                    .Find "author", szAuthor

                                    If MsgBox("Do you want to install the following Scripted style?" & vbCrLf & vbCrLf & _
                                              "Name: " & szName & vbCrLf & _
                                              IIf(szAuthor <> "", "Author: " & szAuthor & vbCrLf, ""), vbQuestion Or vbYesNo, App.Title) = vbYes Then

                                        If g_ExtractToAppData(Path, "styles\script") Then
                                            If FromRunningInstance Then
                                                g_PrivateNotify , "Style installed", szName & " Scripted style was installed successfully", , ".good"
                                                g_StyleRoster.Unload "script.styleengine", False
                                                g_StyleRoster.Load "script.styleengine", False, False
                                                frmAbout.bNotifyDisplaysChanged

                                            Else
                                                MsgBox szName & " was installed successfully!", vbInformation Or vbOKOnly, App.Title

                                            End If
                                            g_InstallSSZ = True

                                        Else
                                            MsgBox "There was a problem installing the style.", vbExclamation Or vbOKOnly, App.Title
                                        
                                        End If
                                    End If
                                    Exit Function

                                Else
                                    g_Debug "g_InstallSSZ(): missing style name", LEMON_LEVEL_CRITICAL

                                End If
                            End With
                        Else
                            g_Debug "g_InstallSSZ(): missing [general] section", LEMON_LEVEL_CRITICAL

                        End If
                    End With
                Else
                    g_Debug "g_InstallSSZ(): failed to create temp folder", LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "g_InstallSSZ(): missing conf or exe", LEMON_LEVEL_CRITICAL

            End If
        Else
            g_Debug "g_InstallSSZ(): failed to open zip file", LEMON_LEVEL_CRITICAL
        
        End If
    End With

    MsgBox Path & " appears to be corrupt.  Try downloading it again.", vbExclamation Or vbOKOnly, App.Title

End Function

Public Function g_IsAlphaBuild() As Boolean

    g_IsAlphaBuild = (InStr(App.Comments, "Alpha") <> 0)

End Function

