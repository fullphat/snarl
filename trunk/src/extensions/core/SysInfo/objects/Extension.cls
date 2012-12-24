VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Extension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           extension.cls
    '/
    '/  Description:    Extension main object
    '/
    '/  Created:        2-Feb-2010
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /*
    '
    '   This class provides information about the Extension itself.  It _must_ implement
    '   both MWndProcSink and MVersionInfo otherwise it won't get loaded by Snarl.  It's
    '   worth noting the the MWndProcSink interface here is not used as a 'normal' window
    '   procedure; it's simply a useful interface that Snarl can send messages to.
    '
    '   See the individual methods and properties for more detailed information.
    '
    '   New for Snarl V40 you can also implement MSimpleEnum which can be used to retrieve more
    '   information about the extension.  Only MSimpleEnum.FindItem() is currently used - all
    '   other functions should return zero or False for now - with the following queries:
    '
    '       'path' - return current extension path (usually App.Path)
    '       'copyright' - return copyright information
    '       'url' - return fully formatted URL to website (e.g. 'http://wwww.fullphat.net/')
    '
    '
    '   IMPORTANT: This is a fully functional Snarl extension - if you're planning on
    '   writing your own extension you should use the template extension in the Snarl SDK.
    '
    ' */

Private Const LIB_DATE = "21-Feb-2012"
Private Const LIB_RELEASE = "3.2c"
'Private Const EXT_NAME = "SysInfo"
Private Const EXT_COMMENT = "Monitors hardware and devices on this computer"

Private Const CLASS_LOW_MEM = "lmem"
Private Const CLASS_NO_MEM = "nmem"
Private Const CLASS_HIGH_CPU = "hcpu"
Private Const CLASS_LOW_DISK = "ldsk"
Private Const CLASS_NIC_CONN = "nicc"
Private Const CLASS_NIC_DISC = "nicd"
Private Const CLASS_DUN_CONN = "dunc"
Private Const CLASS_DUN_DISC = "dund"
Private Const CLASS_VOL_MNT = "volm"
Private Const CLASS_VOL_UMNT = "volu"
Private Const CLASS_POWER_CONNECT = "pwrc"
Private Const CLASS_POWER_DISCONNECT = "pwrd"
Private Const CLASS_POWER_LOW = "pwrl"
Private Const CLASS_LOCK_KEY = "keyb"
Private Const CLASS_IP_CHANGE = "ipch"
Private Const CLASS_IE_DOWNLOAD = "iedl"
Private Const CLASS_CLIPBOARD = "clip"

Dim mLastErr As String
Dim mhWnd As Long                   ' // message handling window
Dim mPassword As String
Dim mConfig As CConfFile

Dim mhWndOwner As Long              ' // handle of the window which launched our config (typically Snarl's Prefs panel)
Dim mPanel As BPrefsPanel

    ' /* timers */
Dim WithEvents theFastTimer As BTimer       ' // 250ms - always runs
Attribute theFastTimer.VB_VarHelpID = -1
Dim WithEvents theSlowTimer As BTimer       ' // 500ms - always runs
Attribute theSlowTimer.VB_VarHelpID = -1
Dim WithEvents the3SecondTimer As BTimer    ' // always runs
Attribute the3SecondTimer.VB_VarHelpID = -1
Dim WithEvents the10SecondTimer As BTimer   ' // always runs
Attribute the10SecondTimer.VB_VarHelpID = -1
Dim WithEvents the5MinuteTimer As BTimer
Attribute the5MinuteTimer.VB_VarHelpID = -1

Dim WithEvents theNetRoster As TNetRoster2
Attribute theNetRoster.VB_VarHelpID = -1
Dim WithEvents theDUNRoster As TDUNRoster
Attribute theDUNRoster.VB_VarHelpID = -1
Dim WithEvents theVolRoster As TSimpleVolRoster
Attribute theVolRoster.VB_VarHelpID = -1
Dim mNotifiedVolsList As BTagList
Dim mAdapterList As BTagList
Dim mIEWindowList As BTagList
Dim mLocalIPList As BPackedData

    ' /* lock key states */
Dim mNumState As Boolean
Dim mCapState As Boolean
Dim mSclState As Boolean

Private Const WM_POWERBROADCAST = &H218
Private Const WM_DEVICECHANGE = &H219
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim mCurrentPower As B_POWER_STATUS
Dim mDoneLowPowerWarning As Boolean
Dim mLowPowerThreshold As Long

Dim mDoneLowRAMWarning As Boolean
Dim mLowRAMThreshold As Long
Dim mDoneNoRAMWarning As Boolean

'Dim mMonitorCPULoad As Boolean
Dim mHighCPUThreshold As Long           ' // percent
Dim mHighCPUDuration As Long            ' // seconds
Dim mCPUIsTracking As Boolean           ' // currently tracking high CPU
Dim mCPUTrackingStarted As Long         ' // tick when tracking was started
Dim mDoneHighCPUWarning As Boolean

Dim mLowDiskThreshold As Long           ' // percent
Dim mWatchRem As Boolean
Dim mWatchOpt As Boolean
Dim mDiskDoneList As BTagList

Dim mCurrentExternalIP As String
Dim WithEvents theExternalIPRequest As CHTTPRequest3
Attribute theExternalIPRequest.VB_VarHelpID = -1
'Dim mIgnoreFirstIPChange As Boolean

Dim mBlockClipboardChange As Boolean

'Implements MVersionInfo             ' // optional: allows the SDOS version command to query the object
Implements ISnarlExtension
Implements KPrefsPanel
Implements KPrefsPage
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

Dim fClipboard As Boolean

    ' /* from the clipboard? */

    If g_ClipboardHandle(hWnd, uMsg, wParam, lParam, fClipboard) Then
        If fClipboard Then
            If mBlockClipboardChange Then
                mBlockClipboardChange = False

            Else
                uOnClipboardChanged

            End If
        End If

        BWndProcSink_WndProc = True
        Exit Function

    End If

Static n As B_POWER_STATUS

    Select Case uMsg
    Case WM_DEVICECHANGE
        g_Debug "##WM_DEVICECHANGE: refreshing volume roster...", LEMON_LEVEL_INFO

        If (NOTNULL(theVolRoster)) And (mConfig.ValueOf("disk.monitor") = "1") Then _
            theVolRoster.Refresh


    Case WM_POWERBROADCAST
        g_Debug "##WM_POWERBROADCAST", LEMON_LEVEL_INFO
        n = get_power_status()
        If n <> mCurrentPower Then
            If n = B_POWER_STATE_CONNECTED Then
                uOnPowerStatusChanged True, get_battery_charge_percent()

            ElseIf n = B_POWER_STATE_DISCONNECTED Then
                uOnPowerStatusChanged False, get_battery_charge_percent()

            End If

            mCurrentPower = n

        End If


    End Select

End Function

Private Function ISnarlExtension_GetConfigWindow() As Long
Dim pPage As BPrefsPage
Dim pCtl As BControl
Dim pp As MPoint
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetWindow 1
            .SetTitle "SysInfo Preferences"
            .SetWidth 128

            ' /* CPU */

            Set pPage = new_BPrefsPage("CPU", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "cpu.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("fancytoggle2", "cpu.enabled", "Monitor system load?", , mConfig.ValueOf("cpu.enabled"))
                .Add new_BPrefsControl("label", "", "Warn when system load is at least:")

                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 100&
                pm.Add "freq", 10&
                pm.Add "label_size", 64&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "cpu.high_mark", "", "", mConfig.ValueOf("cpu.high_mark"), pm, (mConfig.ValueOf("cpu.enabled") = "1"))

                .Add new_BPrefsControl("label", "", "For at least:")

                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 60&
                pm.Add "freq", 5&
                pm.Add "label_size", 64&
                pm.Add "suffix", " seconds"
                .Add new_BPrefsControl("fancyslider", "cpu.duration", "", "", mConfig.ValueOf("cpu.duration"), pm, (mConfig.ValueOf("cpu.enabled") = "1"))
                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "cpu.test", "Test Settings", , , , (mConfig.ValueOf("cpu.enabled") = "1"))

            End With

            .AddPage pPage


            ' /* RAM */

            Set pPage = new_BPrefsPage("Memory", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "memory.png"), Me)
            With pPage
                .SetMargin 32

                .Add new_BPrefsControl("banner", "", "Low Memory")
                .Add new_BPrefsControl("label", "", "Warn when remaining physical RAM reaches:")

                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 90&
                pm.Add "freq", 10&
                pm.Add "label_size", 50&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "ram.low_mark", "", "", mConfig.ValueOf("ram.low_mark"), pm)
                .Add new_BPrefsControl("fancytoggle2", "ram.show_extra", "Include memory information?", , mConfig.ValueOf("ram.show_extra"))
'                .Add new_BPrefsControl("fancytoggle2", "ram.use_meter", "Display meter?", , mConfig.ValueOf("ram.use_meter"))

                .Add new_BPrefsControl("banner", "", "Physical Memory Exhausted")
'                .Add new_BPrefsControl("fancytoggle2", "ram.show_title", "Display title?", , mConfig.ValueOf("ram.show_title"))
                .Add new_BPrefsControl("fancytoggle2", "ram.show_text", "Display hint text?", , mConfig.ValueOf("ram.show_text"))
'                .Add new_BPrefsControl("label", "", "Only applies to the Physical RAM Exhausted notification.")

'                .Add new_BPrefsControl("label", "", "Only applies to the Low Physical RAM notification.")
'                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("label", "", "")
                .Add new_BPrefsControl("fancybutton2", "ram.test", "Test Settings")

            End With

            .AddPage pPage


            ' /* Disk Space */

            Set pPage = new_BPrefsPage("Disk", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "diskfree.png"), Me)
            With pPage
                .SetMargin 32
                
                .Add new_BPrefsControl("fancytoggle2", "disk.monitor", "Watch disk insertions and removals?", , mConfig.ValueOf("disk.monitor"))
                
                .Add new_BPrefsControl("fancytoggle2", "disk.enabled", "Monitor disk space?", , mConfig.ValueOf("disk.enabled"))
                .Add new_BPrefsControl("label", "", "Warn when disk free space drops to:")

                Set pm = New CTempMsg
                pm.Add "min", 0&
                pm.Add "max", 90&
                pm.Add "freq", 10&
                pm.Add "label_size", 50&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "disk.low_mark", "", "", mConfig.ValueOf("disk.low_mark"), pm, (mConfig.ValueOf("disk.enabled") = "1"))

                .Add new_BPrefsControl("fancytoggle2", "disk.watch_optical", "Watch optical media?", , mConfig.ValueOf("disk.watch_optical"), , (mConfig.ValueOf("disk.enabled") = "1"))
                .Add new_BPrefsControl("fancytoggle2", "disk.watch_removable", "Watch removable media (excludes floppy drives)?", , mConfig.ValueOf("disk.watch_removable"), , (mConfig.ValueOf("disk.enabled") = "1"))
'                .Add new_BPrefsControl("label", "", "Note that floppy drives are never scanned, even if the above setting is enabled.")
                .Add new_BPrefsControl("label", "", "Information to include in notification:")
                .Add new_BPrefsControl("fancycycle", "disk.verbosity", "Basic|Simple|Detailed", "", mConfig.ValueOf("disk.verbosity"), , (mConfig.ValueOf("disk.enabled") = "1"))
                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "disk.test", "Test Settings", , , , (mConfig.ValueOf("disk.enabled") = "1"))

            End With

            .AddPage pPage


            ' /* power */

            Set pPage = new_BPrefsPage("Battery", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "battery.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("label", "", "Warn when remaining battery life reaches:")

                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 90&
                pm.Add "freq", 10&
                pm.Add "label_size", 50&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "power.low_mark", "", "", mConfig.ValueOf("power.low_mark"), pm)
                .Add new_BPrefsControl("fancytoggle2", "power.warn_when_charging", "Show notification even if charging?", , mConfig.ValueOf("power.warn_when_charging"))

                .Add new_BPrefsControl("label", "", "Information to include in notification:")
                .Add new_BPrefsControl("fancycycle", "power.verbosity", "Basic|Simple|Detailed", "", mConfig.ValueOf("power.verbosity"))

'                .Add new_BPrefsControl("fancytoggle2", "power.show_text", "Display text?", , mConfig.ValueOf("power.show_text"))
'                .Add new_BPrefsControl("fancytoggle2", "power.use_meter", "Display meter?", , mConfig.ValueOf("power.use_meter"))
'                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("label", "", "")
                .Add new_BPrefsControl("fancybutton2", "power.test", "Test Settings")

            End With

            .AddPage pPage


            ' /* NIC and DUN */

            Set pPage = new_BPrefsPage("Network", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "network.png"), Me)
            With pPage
                .SetMargin 32
                
                .Add new_BPrefsControl("banner", "", "External Address")
                .Add new_BPrefsControl("fancytoggle2", "xip.enabled", "Monitor external IP address changes?", , mConfig.ValueOf("xip.enabled"))

                .Add new_BPrefsControl("banner", "", "Local Addresses")
                .Add new_BPrefsControl("fancytoggle2", "nic.enabled", "Monitor network and dial-up connections?", , mConfig.ValueOf("nic.enabled"))
                .Add new_BPrefsControl("label", "", "Information to include in notification:")
                .Add new_BPrefsControl("fancycycle", "nic.verbosity", "None (Icon Only)|Simple|Detailed", "", mConfig.ValueOf("nic.verbosity"), , (mConfig.ValueOf("nic.enabled") = "1"))
'                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("label", "", "")
                .Add new_BPrefsControl("fancybutton2", "nic.test", "Test Settings (Network)", , , , (mConfig.ValueOf("nic.enabled") = "1"))
                .Add new_BPrefsControl("fancybutton2", "dun.test", "Test Settings (Dial-up / VPN)", , , , (mConfig.ValueOf("nic.enabled") = "1"))

            End With

            .AddPage pPage


            ' /* Other */

            Set pPage = new_BPrefsPage("Other", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "keyboard.png"), Me)
            With pPage
                .SetMargin 32

                .Add new_BPrefsControl("banner", "", "Keyboard")
                .Add new_BPrefsControl("fancytoggle2", "keyb.watch_caps", "Watch Caps Lock?", , mConfig.ValueOf("keyb.watch_caps"))
                .Add new_BPrefsControl("fancytoggle2", "keyb.watch_numl", "Watch Num Lock?", , mConfig.ValueOf("keyb.watch_numl"))
                .Add new_BPrefsControl("fancytoggle2", "keyb.watch_scrl", "Watch Scroll Lock?", , mConfig.ValueOf("keyb.watch_scrl"))

                .Add new_BPrefsControl("banner", "", "Miscellaneous")
                .Add new_BPrefsControl("fancytoggle2", "other.watch_ie", "Notify when an Internet Explorer download completes?", , mConfig.ValueOf("other.watch_ie"))
                .Add new_BPrefsControl("fancytoggle2", "other.watch_clipboard", "Notify when clipboard content changes?", , mConfig.ValueOf("other.watch_clipboard"))

'                .Add new_BPrefsSeparator
'                .Add new_BPrefsControl("fancybutton2", "nic.test", "Test Settings (Network)")
'                .Add new_BPrefsControl("fancybutton2", "dun.test", "Test Settings (Dial-up / VPN)")

            End With

            .AddPage pPage


            ' /* about page */

            Set pPage = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "about.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("banner", "", "Version Information")
                .Add new_BPrefsControl("label", "", "SysInfo Extension R" & LIB_RELEASE)
                .Add new_BPrefsControl("label", "", "Version " & CStr(App.Major) & "." & CStr(App.Revision) & " (" & LIB_DATE & ")")
                .Add new_BPrefsControl("label", "", App.LegalCopyright)
'                .Add new_BPrefsControl("label", "", "")
                .Add new_BPrefsControl("fancybutton2", "go_website", "www.getsnarl.info")

                .Add new_BPrefsControl("banner", "", "Description")
                .Add new_BPrefsControl("label", "", "This extension monitors your computer's hardware and operating system and generates notifications when particular events occur.")

            End With

            .AddPage pPage

            .Go
            g_SimpleSetWindowIcon .hWnd

        End With

    End If

    ISnarlExtension_GetConfigWindow = mPanel.hWnd

End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "fullphat products"
        .Copyright = App.LegalCopyright
        .Date = LIB_DATE
        .Description = EXT_COMMENT
        .Flags = SNARL_EXTN_WANTS_PULSE Or SNARL_EXTN_IS_CONFIGURABLE
        .IconPath = g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png"
        .Name = App.Title
        .Path = App.Path
        .Release = LIB_RELEASE
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With

End Sub

Private Function ISnarlExtension_Initialize() As Long

    On Error Resume Next

    l3OpenLog "%APPDATA%\full phat\snarl\" & LCase$(App.Title) & ".log", True
    g_Debug "ISnarlExtension.Initialise()", LEMON_LEVEL_PROC_ENTER

    g_Debug "checking Snarl version..."
    If snarl_version() < 43 Then
        g_Debug "Needs Snarl V43 or better", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        mLastErr = "Requires Snarl V43 or greater"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

    err.Clear
    get_host_name
    If err.Number <> 0 Then
        g_Debug "misc2.resource missing", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        mLastErr = "Requires misc2.resource V49 or later"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

    ' /* create window */

    g_Debug "creating window..."
    EZRegisterClass App.ProductName
    mhWnd = EZ4AddWindow(App.ProductName, Me)

    If mhWnd = 0 Then
        g_Debug "couldn't create window", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        mLastErr = "Couldn't create handler window"
        ISnarlExtension_Initialize = -1
        Exit Function

    Else
        g_Debug "window is 0x" & g_HexStr(mhWnd)

    End If

    g_Debug "creating password..."
    mPassword = create_password()

    g_Debug "done", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
End Sub

Private Sub ISnarlExtension_Start()

    g_Debug "ISnarlExtension.Start()", LEMON_LEVEL_PROC_ENTER

Dim szPath As String
Dim i As Integer

    On Error Resume Next

    ' /* load config - doesn't matter if it's missing */

    Set mConfig = New CConfFile
    snGetConfigPath szPath

    g_Debug "loading config from " & g_Quote(szPath & "sysinfo.config") & "..."

    With mConfig
        .SetTo szPath & "sysinfo.config", False

        .AddIfMissing "cpu.enabled", "1"        ' // R3.04
        .AddIfMissing "cpu.high_mark", "90"
        .AddIfMissing "cpu.duration", "10"
    
        .AddIfMissing "power.low_mark", "20"
        .AddIfMissing "power.warn_when_charging", "0"
'        .AddIfMissing "power.show_text", "0"
'        .AddIfMissing "power.use_meter", "1"
        .AddIfMissing "power.verbosity", "1"
    
        .AddIfMissing "ram.low_mark", "20"
    '        .addifmissing "ram.show_title", "1"     ' // include title
        .AddIfMissing "ram.show_text", "1"      ' // include help text
'        .AddIfMissing "ram.use_meter", "1"      ' // meter instead of text (low RAM only)
        .AddIfMissing "ram.show_extra", "0"

        .AddIfMissing "nic.enabled", "1"        ' // R3.04
        .AddIfMissing "nic.verbosity", "3"
    
        .AddIfMissing "disk.enabled", "1"       ' // R3.04
        .AddIfMissing "disk.low_mark", "10"
        .AddIfMissing "disk.watch_removable", "1"
        .AddIfMissing "disk.watch_optical", "0"
        .AddIfMissing "disk.verbosity", "3"
    
        .AddIfMissing "keyb.watch_caps", "1"
        .AddIfMissing "keyb.watch_numl", "0"
        .AddIfMissing "keyb.watch_scrl", "0"
    
        .AddIfMissing "other.watch_ie", "0"
        .AddIfMissing "other.watch_clipboard", "0"

        .AddIfMissing "disk.monitor", "1"

        .AddIfMissing "xip.enabled", "0"

    End With

    ' /* register */

    g_Debug "registering with Snarl..."

    If snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png", mPassword) >= 0 Then

        g_Debug "adding classes..."

        uAddClass CLASS_LOW_MEM, "Physical memory low", False
        uAddClass CLASS_NO_MEM, "Physical memory exhausted", False
        uAddClass CLASS_HIGH_CPU, "High processor utilisation"
        uAddClass CLASS_LOW_DISK, "Low disk space", False
        uAddClass CLASS_NIC_CONN, "Network connected", False
        uAddClass CLASS_NIC_DISC, "Network disconnected", False
        uAddClass CLASS_VOL_MNT, "Volume mounted"
        uAddClass CLASS_VOL_UMNT, "Volume unmounted"
        uAddClass CLASS_POWER_CONNECT, "Power connected"
        uAddClass CLASS_POWER_DISCONNECT, "Power disconnected"
        uAddClass CLASS_POWER_LOW, "Low power warning"
        uAddClass CLASS_LOCK_KEY, "Caps, Num or Scroll lock activated"
        uAddClass CLASS_IP_CHANGE, "IP address change"
        ' /* R3.05 */
        uAddClass CLASS_IE_DOWNLOAD, "Internet Explorer download completed"
        uAddClass CLASS_CLIPBOARD, "Clipboard content changed"

    End If

    ' /* call kernel info once to get baseline */

    g_Debug "getting kernel info..."
    krnl_GetSystemInfo
    g_Debug "krnl_GetProcessorCount()=" & CStr(krnl_GetProcessorCount())
    krnl_GetCurrentTotalCPUPercent

    g_Debug "setting up lists and rosters..."

    Set mAdapterList = new_BTagList()
    Set theVolRoster = New TSimpleVolRoster
    Set mNotifiedVolsList = new_BTagList()
    Set mDiskDoneList = new_BTagList()
    Set mIEWindowList = new_BTagList()
    Set mLocalIPList = New BPackedData
    mBlockClipboardChange = True
'    mIgnoreFirstIPChange = True

    g_Debug "getting current lock key states..."

    mCapState = uGetKeyLockState(vbKeyCapital)
    mNumState = uGetKeyLockState(vbKeyNumlock)
    mSclState = uGetKeyLockState(vbKeyScrollLock)

    g_Debug "getting current power state..."

    mCurrentPower = get_power_status()

    g_Debug "getting local IP info..."
    uScanIPTable False

    uConfigChanged "cpu.high_mark"
    uConfigChanged "cpu.duration"
    uConfigChanged "cpu.enabled"                ' // set last

    ' /* NIC/DUN and volume scanning is handled by this object */

    g_Debug "enabling watchers..."
    uConfigChanged "nic.enabled"
    uConfigChanged "disk.monitor"

    uConfigChanged "power.low_mark"
    uConfigChanged "ram.low_mark"
    uConfigChanged "other.watch_clipboard"

    uConfigChanged "disk.low_mark"
    uConfigChanged "disk.watch_optical"
    uConfigChanged "disk.watch_removable"
    uConfigChanged "disk.enabled"               ' // set last

    uConfigChanged "xip.enabled"

    Set theFastTimer = new_BTimer(250)
    Set theSlowTimer = new_BTimer(500)
    Set the3SecondTimer = new_BTimer(3000)
    Set the10SecondTimer = new_BTimer(10000)
    Set the5MinuteTimer = new_BTimer(300000)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub ISnarlExtension_Stop()

    g_Debug "ISnarlExtension.Stop()", LEMON_LEVEL_PROC_ENTER

    g_Debug "stopping watchers..."

    Set theFastTimer = Nothing
    Set theSlowTimer = Nothing
    Set the3SecondTimer = Nothing
    Set the10SecondTimer = Nothing
    Set the5MinuteTimer = Nothing

    g_Debug "unregistering..."
    snarl_unregister App.ProductName, mPassword

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub ISnarlExtension_TidyUp()

    g_Debug "ISnarlExtension.TidyUp()"

    g_Debug "closing window..."
    EZ4RemoveWindow mhWnd
    EZUnregisterClass App.ProductName

    g_Debug "done", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub the10SecondTimer_Pulse()

    If mLowRAMThreshold > 0 Then _
        uCheckRamUsage

    ' /* check no physical ram */

    If mDoneNoRAMWarning Then
        ' /* see if it's back in the black... */
        If g_PercentMemInUse(True) > 0 Then
            g_Debug "(no ram check): returned above threshold", LEMON_LEVEL_INFO
            mDoneNoRAMWarning = False

        End If

    Else
        ' /* is it at or below zero */
        If g_PercentMemInUse() <= 0 Then
            g_Debug "(no ram check): below threshold", LEMON_LEVEL_INFO
            uOnPhysicalRAMGone
            mDoneNoRAMWarning = True

        End If

    End If

    ' /* disk capacity */

    If mConfig.ValueOf("disk.enabled") = "1" Then _
        uScanDiskCapacities

End Sub

Private Sub the3SecondTimer_Pulse()

    uScanIPTable True

    If mLowPowerThreshold > 0 Then _
        uCheckBatteryCharge

    If mConfig.ValueOf("cpu.enabled") = "1" Then _
        uCheckCPULoad

    If mConfig.ValueOf("nic.enabled") = "1" Then
        theNetRoster.Refresh
        theDUNRoster.Refresh

    End If

End Sub

Private Sub the5MinuteTimer_Pulse()

    If mConfig.ValueOf("xip.enabled") = "1" Then _
        uCheckExternalIP

End Sub

Private Sub theDUNRoster_Connected(ByVal Name As String, Details As melon4.BTagList)

    g_Debug "##DUNRoster.Connected(): name=" & g_Quote(Name) & " device_type=" & g_Quote(Details.ValueOf("device_type")), LEMON_LEVEL_INFO

Dim szTitle As String
Dim szText As String

    Select Case Val(mConfig.ValueOf("nic.verbosity"))
    Case 1
        ' /* icon only */

    Case 2
        ' /* simple */
        szTitle = Name
        szText = ""

    Case Else
        ' /* detailed */
        szTitle = g_Capitalize(Details.ValueOf("device_type", "Dial-up"), False) & " connected"
        szText = Name & " is now connected at " & uSpeed(Val(Details.ValueOf("speed", "0"))) & " using " & Details.ValueOf("device_name")

    End Select

'Dim i As Long

    snarl_ez_notify App.ProductName, CLASS_DUN_CONN, szTitle, szText, "!dun-connected", , , mPassword
'    i = sn41EZNotify(mToken, _
'                     CLASS_DUN_CONN, _
'                     szTitle, _
'                     szText, , "!dun-connected")

End Sub

Private Sub theDUNRoster_Disconnected(ByVal Name As String, Details As melon4.BTagList)

    g_Debug "##DUNRoster.Connected(): name=" & g_Quote(Name) & " device_type=" & g_Quote(Details.ValueOf("device_type")), LEMON_LEVEL_INFO

Dim szTitle As String
Dim szText As String

    Select Case Val(mConfig.ValueOf("nic.verbosity"))
    Case 1
        ' /* icon only */

    Case 2
        ' /* simple */
        szTitle = Name
        szText = ""

    Case Else
        ' /* detailed */
        szTitle = g_Capitalize(Details.ValueOf("device_type", "Dial-up"), False) & " disconnected"
        szText = Name & " is disconnected"

    End Select

'Dim i As Long
'
'    i = sn41EZNotify(mToken, CLASS_DUN_DISC, _
'                     szTitle, _
'                     szText, , "!dun-disconnected")

    snarl_ez_notify App.ProductName, CLASS_DUN_DISC, szTitle, szText, "!dun-disconnected", , , mPassword

End Sub

Private Sub uOnClipboardChanged()
Dim pContent As BTagList
Dim bShow As Boolean
Dim szcb As String
Dim sz As String
Dim szic As String
Dim hr As Long
Dim szText As String
Static lPrev As Long
Dim pt As BTagItem

    On Error Resume Next

    g_Debug "uOnClipboardChanged()", LEMON_LEVEL_PROC_ENTER

    Set pContent = clipboard_GetContent()

    If pContent.CountItems > 0 Then
        ' /* apply some intelligence to what's been copied */

        If pContent.IndexOf(CStr(CF_TEXT)) Then
            ' /* standard text - set notification text to it */
            sz = clipboard_GetAsString()
            If InStr(sz, "://") <> 0 Then
                ' /* url: set it as the default callback */
                szcb = sz

            ElseIf g_IsWebColour(sz, hr) Then
                ' /* icon is a block of the colour */
                szic = g_MakePath(App.Path) & "colourblock.png"
                sz = "Colour " & sz & " (" & g_ToRGBNotation(hr) & ")"

                With New mfxView
                    .SetHighColour (hr Or &HFF000000)
                    .EnableSmoothing False
                    .SizeTo 96, 96
                    .FillRect .Bounds
                    .WriteToFile szic, "image/png"

                End With

            Else
                ' /* other text */
                sz = g_Quote(Replace$(sz, vbCrLf, "¶"))

            End If

        ElseIf pContent.IndexOf(CStr(CF_BITMAP)) Then
            ' /* bitmap */
            If OpenClipboard(0) <> 0 Then
                hr = GetClipboardData(CF_BITMAP)
                uSaveBitmap hr, szic
                CloseClipboard

            End If

        Else
            ' /* we don't know about it so list all the types */
            With pContent
                .Rewind
                Do While .GetNextTag(pt) = B_OK
                    sz = sz & clipboard_GetFormatName(Val(pt.Name)) & " (" & g_HexStr(Val(pt.Name)) & ")" & vbCrLf

                Loop
            End With
        End If

        bShow = True

    Else
        g_Debug "nothing in the clipboard", LEMON_LEVEL_CRITICAL

    End If

    ' /* lPrev contains the tick count the last time we notified; if it's less than 500ms
    '    we don't bother displaying the notification in an effort to reduce spam */

    If (bShow) And ((GetTickCount - lPrev) > 500) Then
        snarl_ez_notify App.ProductName, CLASS_CLIPBOARD, "Clipboard changed", sz, IIf(szic <> "", szic, g_MakePath(App.Path) & "clipboard.png"), , , mPassword, , szcb
'        Call sn41EZNotify(mToken, CLASS_CLIPBOARD, "Clipboard changed", sz, , IIf(szic <> "", szic, g_MakePath(App.Path) & "clipboard.png"), , szcb)
        lPrev = GetTickCount()

    End If

    If szic <> "" Then _
        DeleteFile szic

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

'Private Sub uOnExternalIPAddressChanged(ByVal IPAddress As String)
'
'    snarl_ez_notify App.ProductName, CLASS_IP_CHANGE, _
'                    "External IP address changed", "Now " & IPAddress, _
'                    g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "ipaddrchange.png", , , mPassword, _
'                    "extipaddrchange"
'
''    snDoRequest "notify?app-sig=" & App.ProductName & "&password=" & mPassword & "&class=" & CLASS_IP_CHANGE & _
'                "&title=External IP address changed&text=" & IPAddress & _
'                "&icon=" & g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "ipaddrchange.png" & _
'                "&uid=extipaddrchange"
'
'End Sub

Private Sub uOnIEDownloadCompleted()

    snarl_ez_notify App.ProductName, CLASS_IE_DOWNLOAD, "Download completed", , "!system-info", , , mPassword

End Sub

Private Sub uOnLocalIPAddressChanged(ByVal Details As String)

    snDoRequest "notify?app-sig=" & App.ProductName & "&password=" & mPassword & "&class=" & CLASS_IP_CHANGE & _
                "&title=IP address change&text=" & Details & _
                "&icon=" & g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "ipaddrchange.png" & _
                "&uid=ipaddrchange&merge-uid=ipaddrchange"

End Sub

Private Sub uOnCPUMonitoringStarted()

    g_Debug "uOnCPUMonitoringStarted"

'                l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", "Started monitoring" & vbCrLf & _
'                                 "Current utilisation: " & CStr(dCurrentCPU) & "%", , "!system-info")

End Sub

Private Sub uOnCPUMonitoringStopped()

    g_Debug "uOnCPUMonitoringStopped"

'            l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", "Stopped monitoring" & vbCrLf & _
'                             "Current utilisation: " & CStr(dCurrentCPU) & "%", , "!system-info")

End Sub

Private Sub uOnCPUThresholdReached(ByVal LoadPercent As Integer, ByVal Duration As Long)

    snDoRequest "notify?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & CLASS_HIGH_CPU & _
                "&uid=high_cpu_load" & _
                "&title=High processor utilisation" & _
                "&text=The system has been " & CStr(LoadPercent) & "% busy for " & CStr(Duration) & " seconds" & _
                "&icon=!system-warning" & _
                "&action=Task Manager...,taskmgr.exe" & _
                "&action=Run...,!system run_dialog" & _
                "&action=Shutdown...,!system shutdown_dialog"

End Sub

Private Sub uOnDiskThresholdReached(ByVal FreePercent As Integer, ByVal Path As String)
Dim szText As String

    Select Case Val(mConfig.ValueOf("disk.verbosity"))
    Case 1
        ' /* basic: text is "x% remaining" */
        szText = Path & ": " & FreePercent & "% remaining"

    Case 2
        ' /* simple: text is "x of y (%) remaining" */
        With New TSimpleVolume
            If .SetTo(Path) Then
                szText = .SafeLabel & " (" & Path & ") is nearly out of space" & vbCrLf & "Capacity: " & g_FileSizeToStringEx2(.Capacity) & ", free: " & g_FileSizeToStringEx2(.Available)

            Else
                g_Debug "uOnDiskThresholdReached(): couldn't create TSimpleVolume for path " & g_Quote(Path), LEMON_LEVEL_WARNING
                szText = Path & ": " & FreePercent & "% remaining"

            End If

        End With

    Case 3
        ' /* detailed: include filesystem and serial */
        With New TSimpleVolume
            If .SetTo(Path) Then
                szText = .SafeLabel & " (" & Path & ") is nearly out of space" & vbCrLf & "Capacity: " & g_FileSizeToStringEx2(.Capacity) & ", free: " & g_FileSizeToStringEx2(.Available) & vbCrLf & _
                "Filesystem: " & .FileSystem & vbCrLf & "Serial: " & .SerialNumber

            Else
                g_Debug "uOnDiskThresholdReached(): couldn't create TSimpleVolume for path " & g_Quote(Path), LEMON_LEVEL_WARNING
                szText = Path & ": " & FreePercent & "% remaining"

            End If

        End With

    End Select

    snDoRequest "notify?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & CLASS_LOW_DISK & _
                "&title=Low disk space&text=" & szText & "&icon=!disk-low_space" & _
                "&uid=low-disk-space-" & LCase$(Path) & "&value-percent=" & CStr(100 - FreePercent) & _
                "&action=Disk Cleanup,cleanmgr.exe" & _
                "&callback=" & Path & "&callback-label=Open"

End Sub

Private Sub uOnLockKeyPressed(ByVal Key As melon4.B_MODIFIER_KEYS, ByVal Enabled As Boolean)

    Select Case Key
    Case B_CAPS_LOCK
        If (Enabled) And (mConfig.ValueOf("keyb.watch_caps") = "1") Then
            snarl_ez_notify App.ProductName, CLASS_LOCK_KEY, "Caps lock on", , "!system-warning", , , mPassword, "capslock"

        ElseIf Not Enabled Then
            snarl_hide_notification App.ProductName, "capslock", mPassword

        End If

    Case B_NUM_LOCK
        If (Enabled) And (mConfig.ValueOf("keyb.watch_numl") = "1") Then
            snarl_ez_notify App.ProductName, CLASS_LOCK_KEY, "Num lock on", , "!system-warning", , , mPassword, "numlock"

        ElseIf Not Enabled Then
            snarl_hide_notification App.ProductName, "numlock", mPassword

        End If

    Case B_SCROLL_LOCK
        If (Enabled) And (mConfig.ValueOf("keyb.watch_scrl") = "1") Then
            snarl_ez_notify App.ProductName, CLASS_LOCK_KEY, "Scroll lock on", , "!system-warning", , , mPassword, "scrllock"

        ElseIf Not Enabled Then
            snarl_hide_notification App.ProductName, "scrllock", mPassword

        End If

    End Select

End Sub

Private Sub uOnLowPowerWarning(ByVal CurrentChargePercent As Integer)

    g_Debug "uOnLowPowerWarning()", LEMON_LEVEL_PROC_ENTER

    ' /* show even if charging? */

    If (get_power_status() = B_POWER_STATE_CONNECTED) And (mConfig.ValueOf("power.warn_when_charging") = "0") Then
        g_Debug "ignoring: warn_when_charging not set", LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If



Dim szTitle As String
Dim szText As String
Dim i As Integer

    i = -1

    Select Case Val(mConfig.ValueOf("power.verbosity"))
    Case "1"
        ' /* basic: icon only */
        
    Case "2"
        ' /* simple: icon + title */
        szTitle = "Battery level low"

    Case "3"
        ' /* detailed: icon + title + current charge */
        szTitle = "Battery level low"
        szText = "Current battery charge is " & CStr(CurrentChargePercent) & "%"
        i = CurrentChargePercent

    End Select

    snarl_ez_notify App.ProductName, CLASS_POWER_LOW, _
                    szTitle, _
                    szText, "!battery-low", _
                    1, , _
                    mPassword, _
                    "lowpower"


'Dim sz As String
'
'    ' /* matrix */
'    '
'    '   show_text   use_meter
'    '       0           0           icon only
'    '       0           1           icon with meter
'    '       1           0           title and text
'    '       1           1           title and meter
'
'    sz = CStr(CurrentChargePercent)
'
'    If mConfig.ValueOf("power.show_text") = "1" Then
'        If mConfig.ValueOf("power.use_meter") = "0" Then _
'            sz = "Current battery charge is " & sz & "%"
'
'    ElseIf mConfig.ValueOf("power.use_meter") = "0" Then
'        sz = ""
'
'    End If
'
'    snarl_ez_notify App.ProductName, CLASS_POWER_LOW, _
'                    IIf(mConfig.ValueOf("power.show_text") = "1", "Battery getting low", ""), _
'                    sz, "!battery-low", _
'                    1, , _
'                    mPassword, _
'                    "lowpower"

'    mLowPowerId = sn41EZNotify(mToken, _
                               CLASS_POWER_LOW, _
                               IIf(mConfig.ValueOf("power.show_text") = "1", "Battery getting low", ""), _
                               sz, _
                               , _
                               "!battery-low", _
                               1)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub uOnPhysicalRAMGone()

    ' /* if the low RAM warning is still visible, remove it */

    snarl_hide_notification App.ProductName, "lowram", mPassword

Dim sz As String

    If mConfig.ValueOf("ram.show_text") = "1" Then _
        sz = "There is no physical memory available - Windows will begin to move memory to and from disk (paging) which will " & _
             "make the computer respond more slowly." & vbCrLf & vbCrLf & _
             "Closing unwanted applications, browser tabs or system tray applications will help restore physical memory."

    snarl_ez_notify App.ProductName, CLASS_NO_MEM, "Physical memory exhausted", sz, "!system-warning", , , mPassword, "noram"

End Sub

Private Sub uOnPhysicalRAMLow(ByVal PercentRemaining As Integer)
Dim pmx As MEMORYSTATUSEX
Dim szText As String

    szText = g_FileSizeToStringEx2(g_GetPhysMem(False), , , "0.0") & " of " & g_FileSizeToStringEx2(g_GetPhysMem(True), , , "0.0") & " (" & CStr(PercentRemaining) & "%) remaining"

    If mConfig.ValueOf("ram.show_extra") = "1" Then
        ' /* safe to assume we can use GlobalMemoryStatusEx() as Snarl only supports W2K+ */

        pmx.dwLength = Len(pmx)
        If GlobalMemoryStatusEx(pmx) Then
            szText = szText & vbCrLf & "Details (Available / Total):"
            szText = szText & vbCrLf & "Physical: " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullAvailPhys)) & " / " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullTotalPhys))
            szText = szText & vbCrLf & "Paging: " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullAvailPageFile)) & " / " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullTotalPageFile))
            szText = szText & vbCrLf & "Virtual: " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullAvailVirtual)) & " / " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullTotalVirtual))

        Else
            g_Debug "uOnPhysicalRAMLow(): GlobalMemoryStatusEx() failed: " & g_ApiError(), LEMON_LEVEL_WARNING

        End If

    End If

    snarl_ez_notify App.ProductName, CLASS_LOW_MEM, _
                    "Physical memory low", _
                    szText, _
                    "!system-warning", , , _
                    mPassword, "lowram", , PercentRemaining

'    mLowRamId = sn41EZNotify(mToken, _
                             CLASS_LOW_MEM, _
                             "Physical memory low", _
                             sz, , "!system-info")

End Sub

Private Sub uOnPowerStatusChanged(ByVal Connected As Boolean, ByVal CurrentChargePercent As Integer)

    g_Debug "theHandler.OnPowerStatusChanged()", LEMON_LEVEL_PROC_ENTER

Dim szTitle As String
Dim szText As String
Dim i As Integer

    i = -1

    ' /* remove existing notification */

    snarl_hide_notification App.ProductName, "powerchange", mPassword

    If Connected Then
        g_Debug "connected..."

        ' /* remove low power warning if it's still on screen */
        snarl_hide_notification App.ProductName, "lowpower", mPassword

        Select Case Val(mConfig.ValueOf("power.verbosity"))
        Case "1"
            ' /* basic: icon only */
            
        Case "2"
            ' /* simple: icon + title */
            szTitle = "Power connected"

        Case "3"
            ' /* detailed: icon + title + current charge */
            szTitle = "Power connected"
            szText = "Current battery charge is " & CStr(CurrentChargePercent) & "%"
            i = CurrentChargePercent

        End Select

        snarl_ez_notify App.ProductName, CLASS_POWER_CONNECT, _
                        szTitle, _
                        szText, _
                        "!power-connected", , , mPassword, "powerchange", i

'        snarl_ez_notify App.ProductName, CLASS_POWER_CONNECT, _
                        IIf(mConfig.ValueOf("power.show_text") = "1", "Power connected", ""), _
                        sz, _
                        "!power-connected", , , mPassword, "powerchange"

    Else
        g_Debug "disconnected..."

        Select Case Val(mConfig.ValueOf("power.verbosity"))
        Case "1"
            ' /* basic: icon only */
            
        Case "2"
            ' /* simple: icon + title */
            szTitle = "Power disconnected"

        Case "3"
            ' /* detailed: icon + title + current charge */
            szTitle = "Power disconnected"
            szText = "Current battery charge is " & CStr(CurrentChargePercent) & "%"
            i = CurrentChargePercent

        End Select
        
        snarl_ez_notify App.ProductName, CLASS_POWER_DISCONNECT, _
                        szTitle, _
                        szText, _
                        "!power-disconnected", , , mPassword, "powerchange", , i

'        snarl_ez_notify App.ProductName, CLASS_POWER_DISCONNECT, _
                        IIf(mConfig.ValueOf("power.show_text") = "1", "Power disconnected", ""), _
                        sz, _
                        "!power-disconnected", , , mPassword, "powerchange"

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub theFastTimer_Pulse()

    uScanKeyState

    ' /* IP address info */

    If (g_IsPressed(vbKeyEscape)) And (g_IsPressed(vbKeyI)) Then
        If Not snarl_is_notification_visible(App.ProductName, "ipaddrinfo", mPassword) Then _
            snarl_ez_notify App.ProductName, "", "Assigned IP addresses", Replace$(get_ip_address_table(), " ", "; "), "!system-info", , , mPassword, "ipaddrinfo"

    End If

    ' /* nic and dun info */

    If (g_IsPressed(vbKeyEscape)) And (g_IsPressed(vbKeyN)) Then
        If Not snarl_is_notification_visible(App.ProductName, "adapterinfo", mPassword) Then _
            snarl_ez_notify App.ProductName, "", "Network adapters", Replace$(get_ip_address_table(), " ", "; "), "!system-info", , , mPassword, "adapterinfo"

    End If
    
    ' /* memory info */

    If (g_IsPressed(vbKeyEscape)) And (g_IsPressed(vbKeyM)) Then
        If Not snarl_is_notification_visible(App.ProductName, "memoryinfo", mPassword) Then _
            snarl_ez_notify App.ProductName, "", "Memory allocation", uGetMemoryInfo(), "!system-info", , , mPassword, "memoryinfo"

    End If

    ' /* disk info */

    If (g_IsPressed(vbKeyEscape)) And (g_IsPressed(vbKeyV)) Then
        If Not snarl_is_notification_visible(App.ProductName, "diskinfo", mPassword) Then _
            snarl_ez_notify App.ProductName, "", "Disk volumes", uGetVolumeInfo(), "!system-info", , , mPassword, "diskinfo"

    End If
End Sub

Private Sub theNetRoster_Connected(ByVal Name As String, Device As melon4.BTagList)
Dim szTitle As String
Dim szText As String
    
    Select Case Val(mConfig.ValueOf("nic.verbosity"))
    Case 1
        ' /* icon only */

    Case 2
        ' /* simple */
        szTitle = Name
        szText = ""

    Case Else
        ' /* detailed */
        szTitle = "Network connected"
        szText = Name & " is now connected" & vbCrLf & "Speed: " & Device.ValueOf("formatted_speed", "0") '& _
                 IIf(sz <> "", vbCrLf & "Device: " & sz, "")

    End Select

    snarl_hide_notification App.ProductName, "net-dis-" & Name, mPassword

    snDoRequest "notify?app-sig=" & App.ProductName & "&password=" & mPassword & "&class=" & CLASS_NIC_CONN & _
                "&title=" & szTitle & "&text=" & szText & "&icon=!network-connected" & _
                "&uid=net-con-" & Name '& _
                "&style-corp-tagline=" & Device.ValueOf("formatted_speed", "0")

    Debug.Print "dev: " & Device.ValueOf("name")

'    snarl_ez_notify App.ProductName, CLASS_NIC_CONN, szTitle, szText, "!network-connected", , , mPassword

End Sub

Private Sub theNetRoster_Disconnected(ByVal Name As String, ByVal DeviceType As Long)
Dim szTitle As String
Dim szText As String

    Select Case Val(mConfig.ValueOf("nic.verbosity"))
    Case 1
        ' /* icon only */

    Case 2
        ' /* simple */
        szTitle = Name
        szText = ""

    Case Else
        ' /* detailed */
        szTitle = "Network disconnected"
        szText = Name & " is disconnected"

    End Select

    snarl_hide_notification App.ProductName, "net-con-" & Name, mPassword
    snarl_ez_notify App.ProductName, CLASS_NIC_DISC, szTitle, szText, "!network-disconnected", , , mPassword, "net-dis-" & Name

End Sub

Private Function uSpeed(ByVal Speed As Long) As String

    If Speed > 999999999 Then
        uSpeed = CStr(Speed / 1000000000) & "Gbps"

    ElseIf Speed > 999999 Then
        uSpeed = CStr(Speed / 1000000) & "Mbps"

    ElseIf Speed > 999 Then
        uSpeed = CStr(Speed / 1000) & "Kbps"

    Else
        uSpeed = CStr(Speed) & "bps"

    End If

End Function

Private Sub theSlowTimer_Pulse()

    If mConfig.ValueOf("other.watch_ie") = "1" Then _
        uCheckIEWindows

End Sub

Private Sub theVolRoster_VolumeMounted(ByVal Path As String, ByVal Label As String, ByVal VolType As Long)

    g_Debug "##VolumeMounted(): path=" & g_Quote(Path) & " label=" & g_Quote(Label) & " type=" & g_Quote(VolType), LEMON_LEVEL_INFO

    ' /* remove any corresponding "Volume unmounted" notification */
    snarl_hide_notification App.ProductName, "vol-unmounted-" & Path, mPassword

Dim nPercentUsed As Integer
Dim sz As String

    sz = uCapacityInfo(Path, nPercentUsed)

'shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}

    snarl_notify App.ProductName, CLASS_VOL_MNT, "vol-mounted", mPassword, _
                 IIf(VolType = DRIVE_REMOTE, "Share connected", "Disk inserted"), _
                 uSmartenUpLabel(Label, Path) & sz, uIcon(Path), , , Path, _
                 nPercentUsed, "action=My Computer,shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" & _
                 "&callback-label=Open"

'    snarl_ez_notify App.ProductName, CLASS_VOL_MNT, _
                 IIf(VolType = DRIVE_REMOTE, "Share connected", "Disk inserted"), _
                 uSmartenUpLabel(Label, Path) & sz, _
                 uIcon(Path), , , _
                 mPassword, _
                 "vol-mounted-" & Path, _
                 Path, _
                 nPercentUsed


'    snarl_notify App.ProductName, CLASS_VOL_MNT, "vol-mounted", mPassword, _
                 IIf(VolType = DRIVE_REMOTE, "Share connected", "Disk inserted"), _
                 uSmartenUpLabel(Label, Path) & sz, uIcon(Path), , , , _
                 nPercentUsed, "merge-uid=vol-mounted"

End Sub

Private Sub theVolRoster_VolumeUnmounted(ByVal Path As String, ByVal Label As String, ByVal VolType As Long)

    g_Debug "##VolumeUnmounted(): path=" & g_Quote(Path) & " label=" & g_Quote(Label) & " type=" & g_Quote(VolType), LEMON_LEVEL_INFO

    ' /* remove any corresponding "Volume mounted" notification */
    snarl_hide_notification App.ProductName, "vol-mounted-" & Path, mPassword

    snarl_ez_notify App.ProductName, CLASS_VOL_UMNT, _
                    IIf(VolType = DRIVE_REMOTE, "Share disconnected", "Disk ejected"), _
                    uSmartenUpLabel(Label, Path), _
                    "!vol-unmounted", , , _
                    mPassword, _
                    "vol-unmounted-" & Path

'    snarl_notify App.ProductName, CLASS_VOL_UMNT, "vol-unmounted", mPassword, _
                 IIf(VolType = DRIVE_REMOTE, "Share disconnected", "Disk removed"), _
                 uSmartenUpLabel(Label, Path), "!vol-unmounted", , , , , _
                 "merge-uid=vol-unmounted"

End Sub

Private Function uSmartenUpLabel(ByVal Label As String, Optional ByVal Path As String, Optional ByVal AddPath As Boolean = True) As String

    uSmartenUpLabel = IIf(Label = "", "<untitled>", Label)

    If AddPath Then _
        uSmartenUpLabel = uSmartenUpLabel & " (" & Path & ")"

End Function

Private Function uCapacityInfo(ByVal Path As String, ByRef PercentUsed As Integer) As String
Dim pv As TSimpleVolume

    PercentUsed = -1

    Set pv = New TSimpleVolume
    If Not pv.SetTo(Path) Then _
        Exit Function

    If (pv.IsNetworkVolume) Or (Not pv.HasCapacityInfo) Then _
        Exit Function

    PercentUsed = pv.PercentUsed
    uCapacityInfo = vbCrLf & g_FileSizeToStringEx2(pv.Capacity) & ", " & g_FileSizeToStringEx2(pv.Used) & " (" & CStr(pv.PercentUsed) & "%) used"

End Function

Private Function uIcon(ByVal Path As String) As String

    Select Case GetDriveType(Path)
    Case DRIVE_FIXED
        uIcon = "!dev-hd"

    Case DRIVE_CDROM
        uIcon = "!dev-cd"

    Case DRIVE_REMOVABLE
        uIcon = "!dev-fd"

    Case Else
        uIcon = "!vol-mounted"

    End Select

End Function

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)

    If (mConfig Is Nothing) Then _
        Exit Sub

    mConfig.Update Control.GetName, Value
    mConfig.Save

Dim pc As BControl

    Select Case Control.GetName

    Case "cpu.enabled"
        If Control.Page.Find("cpu.high_mark", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("cpu.duration", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("cpu.test", pc) Then _
            pc.SetEnabled (Value = "1")


    Case "disk.enabled"
        If Control.Page.Find("disk.low_mark", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("disk.watch_optical", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("disk.watch_removable", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("disk.verbosity", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("disk.test", pc) Then _
            pc.SetEnabled (Value = "1")


    Case "nic.enabled"
        If Control.Page.Find("nic.verbosity", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("nic.test", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("dun.test", pc) Then _
            pc.SetEnabled (Value = "1")


    End Select

    uConfigChanged Control.GetName

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Select Case Control.GetName
    Case "power.test"
        uOnPowerStatusChanged (get_power_status = B_POWER_STATE_CONNECTED), "50"
        uOnLowPowerWarning 50

    Case "ram.test"
        uOnPhysicalRAMGone
        uOnPhysicalRAMLow 100 - g_PercentMemInUse(True)

    Case "cpu.test"
        uOnCPUThresholdReached mConfig.ValueOf("cpu.high_mark"), mConfig.ValueOf("cpu.duration")

    Case "nic.test"
        theNetRoster_Connected "<test_network>", new_BTagList()
        theNetRoster_Disconnected "<test_network>", 0

    Case "dun.test"
        theDUNRoster_Connected "<test_dialup>", new_BTagList()
        theDUNRoster_Disconnected "<test_dialup>", new_BTagList()

    Case "disk.test"
        uOnDiskThresholdReached mConfig.ValueOf("disk.low_mark"), g_SafeLeftStr(App.Path, 3)

    Case "go_website"
        ShellExecute Control.Page.Panel.hWnd, "open", "http://www.fullphat.net", vbNullString, vbNullString, 5

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
'    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub



'Private Sub theNetRoster_Connected(Details As melon4.BTagList)
'Static i As Long
'
'    If Details.ValueOf("active") = "1" Then _
'        i = sn41EZNotify(mtoken, CLASS_NIC_CONN, "Network connected", _
'                         Details.ValueOf("connection") & " is now connected at " & _
'                         uSpeed(Val(Details.ValueOf("speed", "0"))) & _
'                         " using " & Details.ValueOf("name", "<unknown adapter>"), , "!network-connected")
'
'End Sub
'
'Private Sub theNetRoster_Disconnected(Details As melon4.BTagList)
'Static i As Long
'
'    ' /* was the adapter active at the time it disappeared? */
'
'    If Details.ValueOf("active") = "1" Then _
'        i = sn41EZNotify(mtoken, CLASS_NIC_DISC, "Network disconnected", _
'                         Details.ValueOf("connection") & " is disconnected", , "!network-disconnected")
'
'End Sub
'
'Private Sub theNetRoster_StatusChanged(Details As melon4.BTagList, ByVal PreviousState As Long, ByVal NewState As Long)
'Static i As Long
'
'    If (PreviousState = IF_OPER_STATUS_OPERATIONAL) And (NewState = IF_OPER_STATUS_NON_OPERATIONAL) Then
'        ' /* gone from connected to disconnected */
'        i = sn41EZNotify(mtoken, CLASS_NIC_DISC, "Network disconnected", _
'                         Details.ValueOf("connection") & " is disconnected", , "!network-disconnected")
'
'    ElseIf (PreviousState = IF_OPER_STATUS_NON_OPERATIONAL) And (NewState = IF_OPER_STATUS_OPERATIONAL) Then
'        ' /* gone from disconnected to connected */
'        i = sn41EZNotify(mtoken, CLASS_NIC_CONN, "Network connected", _
'                         Details.ValueOf("connection") & " is now connected at " & _
'                         uSpeed(Val(Details.ValueOf("speed", "0"))) & _
'                         " using " & Details.ValueOf("name", "<unknown adapter>"), , "!network-connected")
'
'    Else
'        i = sn41EZNotify(mtoken, CLASS_NIC_CONN, "Network status changed", _
'                         Details.ValueOf("connection") & " status changed from " & CStr(PreviousState) & " to " & CStr(NewState), , "")
'
'    End If
'
'End Sub

'Private Sub theCPUTimer_Pulse()
'Dim dCurrentCPU As Double
'Dim t As Long
'Dim l As Long
'
'    dCurrentCPU = krnl_GetCurrentTotalCPUPercent()
'
'    If dCurrentCPU > mSettings.CPUThreshold Then
'        ' /* over the threshold */
'
'        If mSettings.CPUHitMonitoring = 0 Then
'            ' /* first time (not tracking) */
'            If mSettings.TellMeWhenMonitoringStarts Then _
'                l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", "Started monitoring" & vbCrLf & _
'                                 "Current utilisation: " & CStr(dCurrentCPU) & "%", , "!system-info")
'
'            mSettings.CPUHitMonitoring = GetTickCount()
'
'        ElseIf Not mSettings.NotifiedHighCPU Then
'            ' /* already tracking */
'
'            t = GetTickCount() - mSettings.CPUHitMonitoring
'            If t > (mSettings.CPUHitDuration * 1000) Then
'                ' /* and past the cut-off duration */
'                l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", _
'                                 "The system has been at least " & mSettings.CPUThreshold & "% busy for " & CStr(t \ 1000) & " seconds", , "!system-info")
'
'                mSettings.NotifiedHighCPU = True
'
'            End If
'
'        End If
'
'    ElseIf mSettings.CPUHitMonitoring > 0 Then
'        ' /* we were monitoring so stop */
'        If mSettings.TellMeWhenMonitoringStops Then _
'            l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", "Stopped monitoring" & vbCrLf & _
'                             "Current utilisation: " & CStr(dCurrentCPU) & "%", , "!system-info")
'
'        mSettings.CPUHitMonitoring = 0
'        mSettings.NotifiedHighCPU = False
'
'    End If
'
'    ' /* if we have a valid notification token and we're running V42+, add actions */
'
'    If (l) And (sn41GetVersion() > 41) Then
'        sn42AddAction l, "Task Manager...", "taskmgr.exe"
'        sn42AddAction l, "Run...", "!system run_dialog"
'        sn42AddAction l, "Shutdown...", "!system shutdown_dialog"
'
'    End If
'
'End Sub
'Private Sub uCheckDiskSpace()
''Const DISK_LIMIT = 90
'Dim szPath As String
'Dim szLabel As String
'Dim lType As Long
'Dim l As Long
'
'    With theVolRoster
'        .Rewind
'
'        Do While .GetNextVolume(szPath, szLabel, lType)
'            If (lType = DRIVE_FIXED) Or (lType = DRIVE_REMOVABLE) Then
'                With New TSimpleVolume
'                    If .SetTo(szPath) Then
'                        If (mVolsDone.IndexOf(szPath) = 0) And (.PercentUsed >= mSettings.DiskThreshold) Then
'
'                            l = sn41EZNotify(mtoken, CLASS_LOW_DISK, uLabel(szLabel, szPath) & " is nearly full", _
'                                             g_FileSizeToStringEx2(.Available) & " of " & g_FileSizeToStringEx2(.Capacity) & " remaining", , "!system-info")
'
'                            mVolsDone.Add new_BTagItem(szPath, "")
'
'                        ElseIf mVolsDone.IndexOf(szPath) Then
'                            ' /* have we dropped below the limit? */
'                            If .PercentUsed < mSettings.DiskThreshold Then _
'                                mVolsDone.Remove mVolsDone.IndexOf(szPath)
'
'                        End If
'
'                    Else
'                        g_Debug "Extension.uCheckDiskSpace(): can't create CVolume for '" & szPath & "'", LEMON_LEVEL_CRITICAL
'
'                    End If
'                End With
'            End If
'        Loop
'    End With
'
'End Sub

'Private Function uIsConnected(ByVal OperStatus As Long) As String
'
'    Select Case OperStatus
'    Case IF_OPER_STATUS_OPERATIONAL
'        uIsConnected = "1"
'
'    Case Else
'        uIsConnected = "0"
'
'    End Select
'
'End Function

'Private Function uFind(ByVal dwIndex As Long) As Long
'
'    If mIFCount = 0 Then _
'        Exit Function
'
'Static pif As MIB_IFROW
'Static cb As Long
'Static i As Long
'
'    cb = Len(pif)
'    For i = 1 To mIFCount
'        ' /* retrieve a single MIB_IFROW */
'        CopyMemory pif, mIFBuffer(4 + (i - 1) * cb), cb
'        If pif.dwIndex = dwIndex Then
'            uFind = i
'            Exit Function
'
'        End If
'
'    Next i
'
'End Function

Private Sub theExternalIPRequest_Completed(ByVal WasSuccessful As Boolean)

    g_Debug "##ExternalIPRequest.Completed()", LEMON_LEVEL_PROC_ENTER

Dim bNotify As Boolean
Dim sz As String

    If WasSuccessful Then
        g_Debug "completed"
        sz = theExternalIPRequest.ResponseBody()

        If g_IsIDE() Then _
            g_Debug "address is " & sz

        If sz <> mCurrentExternalIP Then
            If mCurrentExternalIP <> "" Then
                g_Debug "external IP address has changed", LEMON_LEVEL_INFO

                snarl_ez_notify App.ProductName, CLASS_IP_CHANGE, _
                                "External IP address changed", "Was " & mCurrentExternalIP & ", now " & sz, _
                                g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "ipaddrchange.png", , , mPassword, _
                                "extipaddrchange"


            Else
                g_Debug "not notifying: first pass"

            End If

            mCurrentExternalIP = sz

        Else
            g_Debug "no change"

        End If

    Else
        g_Debug "timed out", LEMON_LEVEL_WARNING

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub


'Private Sub uEnableNICScanning(ByVal Enabled As Boolean)
'
'    g_Debug "uEnableNICScanning()", LEMON_LEVEL_PROC_ENTER
'
'    If Enabled Then
'        g_Debug "enabling..."
'
'        ' /* scan the current NIC and DUN adapters */
'
'        Set theNetRoster = New TNetRoster2
'        With theNetRoster
'            .LockRoster
'            .Refresh
'            .UnlockRoster
'
'        End With
'
'        Set theDUNRoster = New TDUNRoster
'        With theDUNRoster
'            .LockRoster
'            .Refresh
'            .UnlockRoster
'
'        End With
'
'        ' /* start watching */
'
'        Set theNICTimer = new_BTimer(4000)
'        g_Debug "enabled"
'
'    Else
'        g_Debug "disabling..."
'
'        ' /* kill the timer */
'        Set theNICTimer = Nothing
'
'        ' /* delete the rosters */
'        Set theDUNRoster = Nothing
'        Set theNetRoster = Nothing
'        g_Debug "disabled"
'
'    End If
'
'    g_Debug "", LEMON_LEVEL_PROC_EXIT
'
'End Sub

'Private Sub uEnableVolScanning(ByVal Enabled As Boolean)
'
'    g_Debug "uEnableVolScanning()", LEMON_LEVEL_PROC_ENTER
'
'    If Enabled Then
'        g_Debug "enabling..."
'
'        ' /* scan the current mounted volumes */
'
'        With theVolRoster
'            .LockRoster
'            .SetWatchedVolumes False, True
'            .Refresh
'            .UnlockRoster
'
'        End With
'
'        g_Debug "enabled"
'
'    Else
'        g_Debug "disabling..."
'
'        g_Debug "disabled"
'
'    End If
'
'    g_Debug "", LEMON_LEVEL_PROC_EXIT
'
'End Sub

Private Sub uSaveBitmap(ByVal hBitmap As Long, ByRef Path As String)

    If hBitmap = 0 Then _
        Exit Sub

Dim pbm As Long
Dim pbx As mfxBitmap
Dim cx As Long
Dim cy As Long
Dim sz As String

    If GdipCreateBitmapFromHBITMAP(hBitmap, 0, pbm) = Ok Then
        GdipGetImageWidth pbm, cx
        GdipGetImageHeight pbm, cy

        With New mfxView
            .SizeTo cx, cy
            GdipDrawImagePointRectI .Handle, pbm, 0, 0, 0, 0, cx, cy, UnitPixel
            Path = "c:\" & g_CreateGUID(True) & ".bmp"
            .WriteToFile Path, "image/bmp"

        End With

    End If

End Sub

Private Sub uAddClass(ByVal Id As String, ByVal Desc As String, Optional ByVal Enabled As Boolean = True)

    snarl_add_class App.ProductName, Id, Desc, Enabled, mPassword

End Sub

Private Function uGetKeyLockState(ByVal KeyCode As Long) As Boolean
Dim b As Integer

    b = GetKeyState(KeyCode)
    uGetKeyLockState = (b And 1)

End Function

Private Sub uCheckExternalIP()

    g_Debug "uCheckExternalIP()", LEMON_LEVEL_PROC_ENTER

    Set theExternalIPRequest = New CHTTPRequest3
    g_Debug "sending request..."
    theExternalIPRequest.GetContent "http://automation.whatismyip.com/n09230945.asp"

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub uScanIPTable(ByVal Notify As Boolean)
Dim szAdded As String
Dim szRemoved As String
Dim ppd As BPackedData
Dim szn As String
Dim i As Long

'    g_Debug "uScanIPTable(): ", LEMON_LEVEL_PROC_ENTER

    ' /* get current */

    Set ppd = New BPackedData
    ppd.SetTo get_ip_address_table(), " ", ""

'    g_Debug "tidying list..."
    If ppd.Count Then
        For i = ppd.Count To 1 Step -1
            ppd.EntryAt i, szn, ""
            If (szn = "0.0.0.0") Or (szn = "127.0.0.1") Or (szn = "") Then _
                ppd.Remove i

        Next i
    End If

    ' /* look for removals */

    If mLocalIPList.Count > 0 Then
        For i = mLocalIPList.Count To 1 Step -1
            mLocalIPList.EntryAt i, szn, ""
            If Not ppd.Exists(szn) Then
                g_Debug "uScanIPTable(): " & g_Quote(szn) & " removed", LEMON_LEVEL_INFO
                szRemoved = szRemoved & szn & "; "
                mLocalIPList.Remove i

            End If
        Next i
    End If

    ' /* look for additions */

    With ppd
        .Rewind
        Do While .GetNextItem(szn, "")
            If Not mLocalIPList.Exists(szn) Then
                g_Debug "uScanIPTable(): " & g_Quote(szn) & " added", LEMON_LEVEL_INFO
                szAdded = szAdded & szn & "; "
                mLocalIPList.Add szn, ""

            End If
        Loop
    End With

    If szAdded <> "" Then _
        szAdded = "Added: " & g_SafeLeftStr(szAdded, Len(szAdded) - 2)

    If szRemoved <> "" Then _
        szRemoved = "Removed: " & g_SafeLeftStr(szRemoved, Len(szRemoved) - 2)

Dim szText As String

    If szAdded <> "" Then
        szText = szAdded
        If szRemoved <> "" Then _
            szText = szText & vbCrLf & szRemoved

    Else
        szText = szRemoved

    End If

    ' /* only notify if something changed */

    If (szText <> "") And (Notify) Then _
        uOnLocalIPAddressChanged szText

'    g_Debug "", LEMON_LEVEL_PROC_EXIT

'"Current configured addresses are:" & vbCrLf & Replace$(sz, " ", vbCrLf)
'            snShowMessageEx Me.CLASS_1, "IP Address Change", szText, 10, g_MakePath(App.Path) & "icon.png", Me.hWnd, &H401

End Sub

Private Sub uScanKeyState()
Static b As Boolean

    b = uGetKeyLockState(vbKeyCapital)
    If b <> mCapState Then
        mCapState = b
        g_Debug "uScanKeyState(): B_CAPS_LOCK changed to " & CStr(b)
        uOnLockKeyPressed B_CAPS_LOCK, b

    End If

    b = uGetKeyLockState(vbKeyNumlock)
    If b <> mNumState Then
        mNumState = b
        g_Debug "uScanKeyState(): B_NUM_LOCK changed to " & CStr(b)
        uOnLockKeyPressed B_NUM_LOCK, b

    End If

    b = uGetKeyLockState(vbKeyScrollLock)
    If b <> mSclState Then
        mSclState = b
        g_Debug "uScanKeyState(): B_SCROLL_LOCK changed to " & CStr(b)
        uOnLockKeyPressed B_SCROLL_LOCK, b

    End If

End Sub

Private Sub uConfigChanged(ByVal Item As String)
Dim sz As String

    g_Debug "uConfigChanged()", LEMON_LEVEL_PROC_ENTER

    sz = mConfig.ValueOf(Item)
    g_Debug "arg: " & g_Quote(Item) & " current value: " & g_Quote(sz)

    Select Case Item
    Case "cpu.enabled"
        If sz = "1" Then
            g_Debug "system load monitoring started"

        Else
            g_Debug "system load monitoring stopped"

        End If

    Case "cpu.duration"
        mHighCPUDuration = Val(sz)
        g_Debug "CPU duration threshold set to " & CStr(mHighCPUDuration) & " secs"
        mDoneHighCPUWarning = False

    Case "cpu.high_mark"
        mHighCPUThreshold = MAX(MIN(Val(sz), 100), 0)
        g_Debug "CPU threshold set to " & CStr(mHighCPUThreshold) & "%"
        mDoneHighCPUWarning = False



    Case "disk.monitor"
        ' /* volume mounts/unmounts */
        If sz = "1" Then
            g_Debug "enabling disk insert/removal scanning..."
    
            ' /* scan the current mounted volumes */
    
            With theVolRoster
                .LockRoster
                .SetWatchedVolumes False, True
                .Refresh
                .UnlockRoster
    
            End With
    
            g_Debug "enabled"
    
        Else
            g_Debug "disabled disk insertion/removal scanning"
    
        End If





    Case "disk.enabled"
        If sz = "1" Then
            mDiskDoneList.MakeEmpty
            g_Debug "enabling disk capacity monitoring..."

        Else
            g_Debug "TWindow.ConfigChanged(): disabling disk capacity monitoring..."

        End If

    Case "disk.low_mark"
        mLowDiskThreshold = MAX(MIN(Val(sz), 100), 0)
        g_Debug "disk space low threshold set to " & CStr(mLowDiskThreshold) & "%"
        mDiskDoneList.MakeEmpty

    Case "disk.watch_optical"
        mWatchOpt = (sz = "1")

    Case "disk.watch_removable"
        mWatchRem = (sz = "1")



    Case "other.watch_clipboard"
        If sz = "1" Then
            ' /* enabled */
            g_Debug "enabling clipboard snooping..."
            mBlockClipboardChange = True
            g_ClipboardInit mhWnd

        Else
            ' /* disabled */
            g_Debug "disabling clipboard snooping..."
            g_ClipboardUninit

        End If

    Case "other.watch_ie"
        If sz = "1" Then
            g_Debug "IE download monitoring started"

        Else
            g_Debug "IE download monitoring stopped"

        End If

    Case "power.low_mark"
        uSetLowPowerThreshold Val(sz)

    Case "ram.low_mark"
        uSetLowRAMThreshold Val(sz)


    Case "nic.enabled"
        If sz = "1" Then
            g_Debug "enabling network and dial-up scanning..."

            ' /* scan the current NIC and DUN adapters */

            Set theNetRoster = New TNetRoster2
            With theNetRoster
                .LockRoster
                .Refresh
                .UnlockRoster
    
            End With

            Set theDUNRoster = New TDUNRoster
            With theDUNRoster
                .LockRoster
                .Refresh
                .UnlockRoster
    
            End With

            g_Debug "enabled"

        Else
            g_Debug "disabling network and dial-up scanning..."

            ' /* delete the rosters */
            Set theDUNRoster = Nothing
            Set theNetRoster = Nothing
            g_Debug "disabled"

        End If



    Case "xip.enabled"
        ' /* external IP address */
        If sz = "1" Then _
            uCheckExternalIP


    End Select

    g_Debug "", LEMON_LEVEL_PROC_EXIT

'    b = (Config.GetValueWithDefault("cpu.enabled", "") = "1")
'
'    If b <> mWatchingLoad Then
'        g_Debug "ConfigChanged(): cpu.enabled is now " & CStr(b)
'
'        If b Then
'            ' /* start watching load */
'            mHighCPUDuration = Val(Config.GetValueWithDefault("cpu.duration"))
'
'            n = Val(Config.GetValueWithDefault("cpu.high_mark"))
'            n = MAX(MIN(n, 100), 0)
'            mHighCPUThreshold = n
'
'            mDoneHighCPUWarning = False
'
'        Else
'            ' /* stop watching load */
''            mHighCPUThreshold = -1
'
'        End If
'
'        mWatchingLoad = b
'
'    End If

End Sub

Private Sub uSetLowPowerThreshold(ByVal Value As Long)

    g_Debug "uSetLowPowerThreshold()", LEMON_LEVEL_PROC_ENTER

    g_Debug "validating " & g_Quote(Value) & "..."
    If Value > 100 Then
        Value = 100

    ElseIf Value < 0 Then
        Value = 0

    End If

    mLowPowerThreshold = Value
    mDoneLowPowerWarning = False

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub uCheckBatteryCharge()
Dim m As Integer

    ' /* % battery check */
    m = get_battery_charge_percent()
    If mDoneLowPowerWarning Then
        If m > mLowPowerThreshold Then
            g_Debug "uCheckBatteryCharge(): returned above threshold", LEMON_LEVEL_INFO
            mDoneLowPowerWarning = False

        End If

    Else
        If m <= mLowPowerThreshold Then
            g_Debug "uCheckBatteryCharge(): below threshold", LEMON_LEVEL_INFO
            uOnLowPowerWarning get_battery_charge_percent()
            mDoneLowPowerWarning = True

        End If

    End If

End Sub

Private Sub uSetLowRAMThreshold(ByVal Value As Long)

    g_Debug "uSetLowRAMThreshold()", LEMON_LEVEL_PROC_ENTER

    g_Debug "validating " & g_Quote(Value) & "..."
    If Value > 100 Then
        Value = 100

    ElseIf Value < 0 Then
        Value = 0

    End If

    mLowRAMThreshold = Value
    mDoneLowRAMWarning = False

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub uCheckRamUsage()
Dim d As Double

    ' /* % physical RAM check */

    d = g_PercentMemInUse(True)

    If mDoneLowRAMWarning Then
        ' /* see if it's back above the warning limit */
        If d > mLowRAMThreshold Then
            g_Debug "uCheckRamUsage(): has returned above threshold", LEMON_LEVEL_INFO
            mDoneLowRAMWarning = False

        End If

    Else
        ' /* is it at or below the warning limit */
        If d <= mLowRAMThreshold Then
            g_Debug "uCheckRamUsage(): has dropped below threshold", LEMON_LEVEL_INFO
            uOnPhysicalRAMLow d
            mDoneLowRAMWarning = True

        End If

    End If

End Sub

Private Sub uScanDiskCapacities()
Dim szPath As String
Dim i As Integer

    For i = Asc("C") To Asc("Z")

        szPath = Chr$(i) & ":\"

        Select Case GetDriveType(szPath)
        Case DRIVE_FIXED
            uCheckDiskCapacity szPath

        Case DRIVE_REMOVABLE
            If mWatchRem Then _
                uCheckDiskCapacity szPath

        Case DRIVE_CDROM
            If mWatchOpt Then _
                uCheckDiskCapacity szPath
        
        End Select

    Next i

End Sub

Private Sub uCheckDiskCapacity(ByVal Path As String)
Dim n As Integer

    n = g_GetDiskFreePercent(Path)
    If n < 0 Then _
        Exit Sub

    If (n <= mLowDiskThreshold) And (mDiskDoneList.IndexOf(Path) = 0) Then
        ' /* on or below threshold and not already notified - notify and add to list */
        g_Debug "uCheckDiskCapacity(): " & g_Quote(Path) & " is below threshold", LEMON_LEVEL_INFO
        mDiskDoneList.Add new_BTagItem(Path, "")
        uOnDiskThresholdReached n, Path

    ElseIf (mDiskDoneList.IndexOf(Path) <> 0) And (n > mLowDiskThreshold) Then
        ' /* notified but now above threshold - remove from list */
        g_Debug "uCheckDiskCapacity(): " & g_Quote(Path) & " has returned above threshold", LEMON_LEVEL_INFO
        mDiskDoneList.Remove mDiskDoneList.IndexOf(Path)

    End If

End Sub

Private Sub uCheckCPULoad()
Dim dLoad As Double

    dLoad = krnl_GetCurrentTotalCPUPercent()

    If dLoad > mHighCPUThreshold Then
        ' /* over the threshold */

        g_Debug "uCheckCPULoad(): current load (" & CStr(dLoad) & "%) is over threshold (" & CStr(mHighCPUThreshold) & "%)"

        If Not mCPUIsTracking Then
            ' /* first time (not tracking) */
            mCPUTrackingStarted = GetTickCount()
            mCPUIsTracking = True
            g_Debug "uCheckCPULoad(): started tracking..."
            uOnCPUMonitoringStarted

        ElseIf Not mDoneHighCPUWarning Then
            ' /* already tracking... */
            If (GetTickCount() - mCPUTrackingStarted) > (mHighCPUDuration * 1000) Then
                ' /* ...and past the cut-off duration */
                g_Debug "uCheckCPULoad(): tracking duration exceeded"
                uOnCPUThresholdReached mHighCPUThreshold, CInt((GetTickCount() - mCPUTrackingStarted) / 1000)
                mDoneHighCPUWarning = True

            End If

        End If

    ElseIf mCPUIsTracking Then
        ' /* load has dropped below the threshold and we were monitoring, so stop */
        g_Debug "uCheckCPULoad(): stopped tracking..."
        uOnCPUMonitoringStopped
        mCPUIsTracking = False
        mDoneHighCPUWarning = False

    End If

End Sub

Private Sub uCheckIEWindows()
Static pt As BTagItem
Static h As Long
Static c As Long
Static i As Long
Static p As Long

    ' /* gone windows */

    With mIEWindowList
        If .CountItems Then
            For i = .CountItems To 1 Step -1
                h = Val(.TagAt(i).Name)
                If IsWindow(h) = 0 Then
                    Debug.Print "removed " & g_HexStr(h)
                    .Remove i
            
                End If
            Next i
        End If
    End With

    ' /* new windows */

    c = g_AppWindows_Refresh(True)
    If c Then
        For i = 1 To c
            h = g_AppWindows_WindowAt(i)
            If (g_ClassName(h) = "#32770") And (g_WindowText(h) = "Download complete") And (mIEWindowList.IndexOf(CStr(h)) = 0) Then
                GetWindowThreadProcessId h, p
                If g_GetProcessFriendlyName2(p) = "Internet Explorer" Then
                    Debug.Print "found window 0x" & g_HexStr(h)
                    mIEWindowList.Add new_BTagItem(CStr(h), "")
                    uOnIEDownloadCompleted

                End If
            End If
        Next i
    End If

End Sub

Private Function uGetMemoryInfo() As String
Dim pmx As MEMORYSTATUSEX

    ' /* safe to assume we can use GlobalMemoryStatusEx() as Snarl only supports W2K+ */

    pmx.dwLength = Len(pmx)
    If GlobalMemoryStatusEx(pmx) Then
        uGetMemoryInfo = uGetMemoryInfo & "Physical: " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullTotalPhys)) & ", " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullAvailPhys)) & " free"
        uGetMemoryInfo = uGetMemoryInfo & vbCrLf & "Paging: " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullTotalPageFile)) & ", " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullAvailPageFile)) & " free"
        uGetMemoryInfo = uGetMemoryInfo & vbCrLf & "Virtual: " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullTotalVirtual)) & ", " & g_FileSizeToStringEx2(g_LargeToDouble(pmx.ullAvailVirtual)) & " free"

    Else
        g_Debug "uGetMemoryInfo(): GlobalMemoryStatusEx() failed: " & g_ApiError(), LEMON_LEVEL_WARNING
        uGetMemoryInfo = "Not available"

    End If

End Function

Private Function uGetVolumeInfo() As String
Dim szp As String
Dim szl As String
Dim dwt As Long

    With New TSimpleVolRoster
'        .SetWatchedVolumes False, True
        .Refresh
        .Rewind
        Do While .GetNextVolume(szp, szl, dwt)
            uGetVolumeInfo = uGetVolumeInfo & IIf(szl = "", "<unnamed>", UCase$(szl)) & " (" & UCase$(szp) & "): " & CStr(100 - g_GetDiskFreePercent(szp)) & "% full" & vbCrLf

        Loop

    End With

    uGetVolumeInfo = g_SafeLeftStr(uGetVolumeInfo, Len(uGetVolumeInfo) - 1)

End Function
