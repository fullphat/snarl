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
    '/  Created:        8-Apr-2009
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
    '       'url' - return fully formatted URL to website (e.g. 'http://www.fullphat.net/')
    '
    '
    '   IMPORTANT: This is a fully functional Snarl extension - if you're planning on
    '   writing your own extension you should use the template extension in the Snarl SDK.
    '
    ' */

Private Const LIB_DATE = "5-Aug-2010"
Private Const LIB_RELEASE = "1.0"
Private Const EXT_NAME = "DoNotDisturb"
Private Const EXT_COMMENT = "Manages the Do Not Disturb setting in response to certain system events"

Private Const WINDOW_CLASS = "w>claire2707"

Dim mLastErr As String
Dim mhWnd As Long

'Private Type T_APP
'    Path As String
'
'End Type

'Private Type T_CONFIG
'    WatchSpecificApps As Boolean
'    App() As T_APP
'    Apps As Long
'    DebugMode As Boolean
'
'End Type

'Dim mSettings As T_CONFIG
Dim mConfig As CConfFile
Dim mPanel As BPrefsPanel           ' // our prefs panel

Dim WithEvents theSessionEvents As TSessionEvents
Attribute theSessionEvents.VB_VarHelpID = -1

Implements KPrefsPanel              ' // we're a panel handler
Implements KPrefsPage               ' // we're also a page handler

Implements ISnarlExtension
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

End Function

Private Sub Class_Initialize()

'    lemonOpenLog App.Title
    g_Debug App.Title & " running from '" & g_MakePath(App.Path) & "'"
    g_Debug ""

End Sub

Private Sub Class_Terminate()

'    lemonCloseLog App.Title

End Sub

Private Function ISnarlExtension_GetConfigWindow() As Long
Dim pPage As BPrefsPage
Dim pCtl As BControl
Dim pt As TPoint

    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetWindow 1

            Set pPage = new_BPrefsPage("", , Me)
            With pPage
                .SetMargin 24
                .SetWidth 356
'                .Add new_BPrefsControl("label", "", " ")
                .Add new_BPrefsControl("banner", "", "Enable Do Not Disturb mode:")

                .Add new_BPrefsControl("fancytoggle2", "when_idle", "After the following period of inactivity:", , mConfig.GetValueWithDefault("when_idle", "0"))
                Set pCtl = new_BPrefsControl("fancycycle", "idle_time", "5 mins|10 mins|15 mins|30 mins|1 hour", , mConfig.GetValueWithDefault("idle_time", "1"), , mConfig.GetValueWithDefault("when_idle", "0") = "1")
                
                Set pt = New TPoint
'                pt.SetTo 32, 0
                pt.SetFrom pCtl.Position
                pt.OffsetBy 32
                pCtl.SetPosition pt

                .Add pCtl

                .Add new_BPrefsControl("fancytoggle2", "when_fullscreen", "When the foreground application is full-screen?", , mConfig.GetValueWithDefault("when_fullscreen", "0"))
'                .Add new_BPrefsControl("fancybutton2", "fullscreen_settings", "Settings...")

                .Add new_BPrefsControl("fancytoggle2", "when_screensaver", "When the screensaver starts?", , mConfig.GetValueWithDefault("when_screensaver", "0"))
                .Add new_BPrefsControl("fancytoggle2", "when_locked", "When the computer is locked?", , mConfig.GetValueWithDefault("when_locked", "0"))
                .Add new_BPrefsControl("fancytoggle2", "when_standby", "When the computer enters standby mode?", , mConfig.GetValueWithDefault("when_standby", "0"))
'                .Add new_BPrefsControl("fancytoggle2", "when_lid_shut", "When the lid on my computer is shut?", , mConfig.GetValueWithDefault("when_lid_shut", "0"))

            End With
            
            .AddPage pPage

            .SetTitle "DoNotDisturb Preferences"
            .SetWidth 128
            .Go
'            g_SetWindowIconToAppResourceIcon .hWnd

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
        .Flags = SNARL_EXTN_IS_CONFIGURABLE
        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Name = EXT_NAME
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

    l3OpenLog "%APPDATA%\" & EXT_NAME & ".log"
    g_Debug "pwd: " & g_MakePath(App.Path)

    If sn41GetVersion = 0 Then
        g_Debug "Needs Snarl V41 or better", LEMON_LEVEL_CRITICAL
        mLastErr = "Requires Snarl V41 or greater"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

    Set mConfig = New CConfFile

'Dim szPath As String
'
'    mLastErr = ""
'    uInit = M_FAILED
'
'    ' /* does snarl's /etc/ path exist? */
'
'    If Not snext_getuserpath(szPath) Then
'        mLastErr = "Couldn't get Snarl user path"
'        uInit = M_NOT_FOUND
'        Exit Function
'
'    End If
'
'    ' /* create our handling window (could use a VB form here as well) */
'
    If Not EZRegisterClass(WINDOW_CLASS) Then
        ' /* no class */
        mLastErr = "Failed to create window class"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

    mhWnd = EZAddWindow(WINDOW_CLASS, Me)
    If mhWnd = 0 Then
        ' /* no window */
        mLastErr = "Failed to create window"
        EZUnregisterClass WINDOW_CLASS
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

    ' /* set up our events monitor */

    Set theSessionEvents = New TSessionEvents
    theSessionEvents.Init mhWnd

'    ' /* defaults */
'
'    With mSettings
'        .WatchSpecificApps = False
'        .Apps = 0
'        ReDim .App(0)
'        .DebugMode = False
'
'    End With
'
'    ' /* load config */
'
'Dim sz As String
'Dim i As Long
'
'    With New ConfigFile
'        .File = szPath & "etc\.dnd"
'
'        If .Load() Then
'            i = .FindSection("general")
'            If i Then
'                With .SectionAt(i)
'                    If .Find("WatchSpecificApps", sz) Then _
'                        mSettings.WatchSpecificApps = Val(sz)
'
'                    If .Find("DebugMode", sz) Then _
'                        mSettings.DebugMode = Val(sz)
'
'                End With
'            End If
'
'Dim pe As ConfigEntry
'
'            i = .FindSection("apps")
'            If i Then
'                With .SectionAt(i)
'                    .Rewind
'                    Do While .GetNextEntry(pe)
'                        uAdd pe.Value
'
'                    Loop
'
'                End With
'            End If
'
'        End If
'
'    End With
'
'    ' /* ok - no need to register any alert classes! */
'
'    uInit = M_OK
'    Exit Function
'
'ex:
'    MsgBox err.Description

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
End Sub

Private Sub ISnarlExtension_Start()
Dim szUser As String

    If snext_getuserpath(szUser) Then
        If mConfig.SetTo(g_MakePath(szUser) & ".donotdisturb") Then


        End If

    End If

End Sub

Private Sub ISnarlExtension_Stop()

End Sub

Private Sub ISnarlExtension_TidyUp()

    EZRemoveWindow mhWnd
    EZUnregisterClass WINDOW_CLASS

End Sub

'        snManageSnarl E_SET_DND_MODE, IIf(Enabled, 1, 0)
'
'End Sub

'Private Sub uWriteConfig()
'Dim szPath As String
'Dim i As Long
'
'    snext_getuserpath szPath
'
'    With New ConfigFile
'        .File = g_MakePath(szPath) & "etc\.dnd"
'
'        With .AddSectionObj("general")
'            .Add "WatchSpecificApps", IIf(mSettings.WatchSpecificApps, "1", "0")
'            .Add "DebugMode", IIf(mSettings.DebugMode, "1", "0")
'
'        End With
'
'        If mSettings.Apps Then
'            With .AddSectionObj("apps")
'                For i = 1 To mSettings.Apps
'                    .Add "app", mSettings.App(i).Path
'
'                Next i
'
'            End With
'        End If
'
'        .Save
'
'    End With
'
'End Sub

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
'Dim pc As BControl
'Dim pm As CTempMsg
'Dim sz As String
'
'    On Error Resume Next
'
'    With mPage
'        .SetMargin 0
'
'        .Add new_BPrefsControl("fancytoggle2", "ft>WatchSpecificApps", "Watch specific applications?", , IIf(mSettings.WatchSpecificApps, "1", "0"))
'
'        Set pm = New CTempMsg
'        pm.Add "item-height", 24&
'        Set pc = new_BPrefsControl("listbox", "lb>apps", sz, , , pm)
'        pc.SizeTo 0, 100
'        mPage.Add pc
'
'        .Add new_BPrefsControl("fancyplusminus", "fpm>apps")
'
''        Set pm = New CTempMsg
''        pm.Add "min", 0&
''        pm.Add "max", 60&
''        pm.Add "freq", 5&
''        pm.Add "label_size", 50&
''        .Add new_BPrefsControl("fancyslider", "fs>NotificationTimeout", "", "", CStr(0), pm)
'
'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("label", "", App.FileDescription & " V" & CStr(App.Major) & "." & CStr(App.Revision), , , , False)
'
'        If g_IsPressed(VK_LCONTROL) Then _
'            .Add new_BPrefsControl("fancytoggle2", "ft>DebugMode", "Debug Mode?", , IIf(mSettings.DebugMode, "1", "0"))
'
'    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
'Dim pc As BControl
'Dim sz As String
'Dim i As Long

'    g_Debug "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    mConfig.Update Control.GetName, Value
    mConfig.Save

    Select Case Control.GetName
    Case "when_idle"
        prefskit_SafeEnable mPanel, "idle_time", (Value = "1")

    End Select

'    Select Case Control.GetName()
'
'    Case "ft>WatchSpecificApps"
'        mSettings.WatchSpecificApps = (Value = "1")
'        uWriteConfig
'
'        If mPage.Find("lb>apps", pc) Then _
'            pc.SetEnabled mSettings.WatchSpecificApps
'
'        If mPage.Find("fpm>apps", pc) Then _
'            pc.SetEnabled mSettings.WatchSpecificApps
'
'
'    Case "ft>DebugMode"
'        mSettings.DebugMode = (Value = "1")
'        uWriteConfig
'
'
'    Case "fpm>apps"
'        If Value = "+" Then
'            With New CFileDialog
'                .DialogType = E_DIALOG_OPEN
'                .AddFilter "Applications", "exe", True, True
'                .Title = ""
'                .InitialPath = IIf(mLastPath = "", g_GetSystemFolderStr(CSIDL_PROGRAMFILES), mLastPath)
'                If .Go(True, E_FILE_DIALOG_CENTRE_OWNER, mPanel.hWnd) Then
'                    mLastPath = g_GetPath(.SelectedFile)
'                    uAdd .SelectedFile
'
'                End If
'
'            End With
'
'        Else
'            ' /* remove selected */
'
'            If Not mPage.Find("lb>apps", pc) Then _
'                Exit Sub
'
'            uRemove Val(pc.GetValue())
'
'        End If
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

'    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

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

Private Sub uAdd(ByVal Path As String)
Dim i As Long

'    With mSettings
'        ' /* check for duplicates */
'        i = uFind(Path)
'        If i > 0 Then _
'            Exit Sub
'
'        .Apps = .Apps + 1
'        ReDim Preserve .App(.Apps)
'        .App(.Apps).Path = Path
'
'    End With
'
'    uWriteConfig
'    uUpdateList "0"

End Sub

Private Sub uUpdateList(ByVal CurrentValue As String)
Dim pc As BControl
Dim sz As String
Dim szApp As String
Dim i As Long

'    If (mPage Is Nothing) Then _
'        Exit Sub
'
'    If Not mPage.Find("lb>apps", pc) Then _
'        Exit Sub
'
'    With mSettings
'        If .Apps > 0 Then
'            For i = 1 To .Apps
'                szApp = g_GetExeDescription(.App(i).Path)
'                If szApp = "" Then _
'                    szApp = LCase$(g_FilenameFromPath(.App(i).Path))
'
'                sz = sz & szApp & IIf(i < .Apps, "|", "")
'
'            Next i
'        End If
'
'    End With
'
'    pc.SetText sz
'    pc.SetValue "1"

End Sub

Private Function uFind(ByVal Path As String) As Long
'Dim i As Long
'
'    Path = LCase$(Path)
'
'    With mSettings
'        If .Apps > 0 Then
'            For i = 1 To .Apps
'                If Path = LCase$(.App(i).Path) Then
'                    uFind = i
'                    Exit Function
'
'                End If
'            Next i
'        End If
'
'    End With

End Function


Private Sub uRemove(ByVal Index As Long)

'    If (Index < 1) Or (Index > mSettings.Apps) Then _
'        Exit Sub
'
'Dim i As Long
'
'    If Index < mSettings.Apps Then
'        For i = Index To (mSettings.Apps - 1)
'            LSet mSettings.App(i) = mSettings.App(i + 1)
'
'        Next i
'
'    End If
'
'    mSettings.Apps = mSettings.Apps - 1
'    ReDim Preserve mSettings.App(mSettings.Apps)
'
'    uWriteConfig
'    uUpdateList "0"

End Sub

Private Sub theSessionEvents_ConsoleLocked(ByVal Id As Long)

    If mConfig.ValueOf("when_locked") = "1" Then _
        snManageSnarl E_SET_DND_MODE, 1

End Sub

Private Sub theSessionEvents_ConsoleUnlocked(ByVal Id As Long)

    If mConfig.ValueOf("when_locked") = "1" Then _
        snManageSnarl E_SET_DND_MODE, 0

End Sub

Private Sub theSessionEvents_FullScreenApp(ByVal IsActive As Boolean)

    If mConfig.ValueOf("when_fullscreen") = "1" Then _
        snManageSnarl E_SET_DND_MODE, IIf(IsActive, 1, 0)

End Sub

Private Sub theSessionEvents_IdleStateChanged(ByVal IsIdle As Boolean)

    If mConfig.ValueOf("when_idle") = "1" Then _
        snManageSnarl E_SET_DND_MODE, IIf(IsIdle, 1, 0)

End Sub

Private Sub theSessionEvents_ScreensaverStateChanged(ByVal IsRunning As Boolean)

    If mConfig.ValueOf("when_screensaver") = "1" Then _
        snManageSnarl E_SET_DND_MODE, IIf(IsRunning, 1, 0)

End Sub

Private Sub theSessionEvents_SystemGoingToStandby()

    If mConfig.ValueOf("when_standby") = "1" Then _
        snManageSnarl E_SET_DND_MODE, 1

End Sub

Private Sub theSessionEvents_SystemLidStateChanged(ByVal IsShut As Boolean)

    If mConfig.ValueOf("when_lid_shut") = "1" Then _
        snManageSnarl E_SET_DND_MODE, IIf(IsShut, 1, 0)

End Sub

Private Sub theSessionEvents_SystemResumed()

    If mConfig.ValueOf("when_standby") = "1" Then _
        snManageSnarl E_SET_DND_MODE, 0

End Sub
