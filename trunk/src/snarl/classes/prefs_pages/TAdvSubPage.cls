VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAdvSubPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TAdvSubPage.cls
    '/
    '/  Description:    Provides the various sub-pages within the [Advanced] page
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Name As String

Dim mPage As BPrefsPage
Dim mStyle As TStyle
Dim mThemes As BTagList

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
Dim pm As CTempMsg

    With mPage

        .SetMargin 48

        Select Case Name

        Case "gen"
'            ' /* hotkeys */
'            .Add new_BPrefsControl("banner", "", "Hotkeys")
'            .Add new_BPrefsControl("fancytoggle2", "use_hotkey", "Use a hotkey to activate Snarl's Preferences?", "", g_ConfigGet("use_hotkey"))
'            .Add new_BPrefsControl("key_picker", "hotkey_prefs", , , CStr(MOD_CONTROL) & "," & g_ConfigGet("hotkey_prefs"), , (g_ConfigGet("use_hotkey") = "1"))
'            .Add new_BPrefsControl("fancytoggle2", "", "Use a hotkey to activate Snarl's menu?", "", "0", , False)
'            .Add new_BPrefsControl("key_picker", "", , , CStr(MOD_WIN) & "," & g_ConfigGet("hotkey_prefs"), , False)
'            .Add new_BPrefsControl("label", "", "Press the key you want to use in the boxes above.  Note that the modifiers (the combination of SHIFT and CTRL keys) used are automatically set.")
'
'            ' /* legacy support */
'            .Add new_BPrefsControl("banner", "", "Legacy Support")
'            .Add new_BPrefsControl("fancytoggle2", "allow_right_clicks", "Allow notification right and middle clicks?", "", g_ConfigGet("allow_right_clicks"))
'            .Add new_BPrefsControl("label", "", "This only applies to notifications created using the V42 API; notifications created using a previous version of the API always receive these events.")
'
'            ' /* other */
'            .Add new_BPrefsControl("banner", "", "Other")
'            .Add new_BPrefsControl("fancytoolbar", "general_toolbar", "Recreate Tray Icon|Empty Icon Cache||Restore Factory Settings")


        Case "sec"
            ' /* security */
'            .SetMargin 48
'            .Add new_BPrefsControl("banner", "", "Applications")
'            .Add new_BPrefsControl("fancytoggle2", "apps_must_register", "Applications must register before creating notifications?", "", g_ConfigGet("apps_must_register"))
'            .Add new_BPrefsControl("fancytoggle2", "only_allow_secure_apps", "Only allow password-protected applications?", "", g_ConfigGet("only_allow_secure_apps"))
'            .Add new_BPrefsControl("label", "", "Applications can provide a password when registering with Snarl which makes it harder for rogue applications to generate notifications by spoofing a legitimate application.  Note that enabling this setting will prevent existing non-password protected applications from registering.")
'
''            .Add new_BPrefsControl("fancytoggle2", "require_password", "Require a password for remote applications?", "", g_ConfigGet("require_password"))
''            Set pm = New CTempMsg
''            pm.Add "password", 1&
''            .Add new_BPrefsControl("fancyedit", "network_password", g_GetPassword(), "Password:", , pm)
''            .Add new_BPrefsControl("label", "", "This setting applies to applications registering via SNP3 or GNTP from a remote source.")
'
'            .Add new_BPrefsControl("banner", "", "Notifications")
'            .Add new_BPrefsControl("fancytoggle2", "no_callback_urls", "Prevent URLs being used as the default callback?", "", g_ConfigGet("no_callback_urls"))
'            .Add new_BPrefsControl("fancytoggle2", "ignore_style_requests", "Ignore requests for specific styles?", "", g_ConfigGet("ignore_style_requests"))


        Case "dbg"
            ' /* debug */

            .Add new_BPrefsControl("banner", "", "Debugging")

            .Add new_BPrefsControl("fancytoolbar", "debug_toolbar", "View Debug Log|Invoke SOS|Garbage Collection")

'            .Add new_BPrefsControl("fancybutton2", "go_lemon", "View Debug Log")
'            .Add new_BPrefsControl("fancybutton2", "go_sos", "Invoke SOS", , , , False)
'    '        .Add new_BPrefsControl("label", "", "The log file can be useful for debugging purposes.")
'
'            .Add new_BPrefsControl("fancybutton2", "go_garbage", "Garbage Collection", , , , g_IsWinXPOrBetter())
    
    '        .Add new_BPrefsControl("separator", "")
            .Add new_BPrefsControl("banner", "", "Configuration")
            .Add new_BPrefsControl("fancybutton2", "open_config", "Open Config Folder")
            .Add new_BPrefsControl("label", "", "Opens the current config folder in Explorer so the various configuration files can be edited manually.")
    
    '        .Add new_BPrefsControl("fancybutton2", "cycle_config", "Reload Config File")
    '        .Add new_BPrefsControl("label", "", "Reloads the current configuration file.")
    
    '        .Add new_BPrefsControl("separator", "")
            .Add new_BPrefsControl("banner", "", "Diagnostics")
            .Add new_BPrefsControl("fancybutton2", "test", "Test Notification")
            .Add new_BPrefsControl("label", "", "Sends a special test message to the Snarl engine which should result in a notification appearing.  This message is sent using the same mechanism a 3rd party application would use and therefore should prove (or otherwise) that the Snarl notification engine is running correctly.")
    
    '        .Add new_BPrefsControl("separator", "")
    '        .Add new_BPrefsControl("fancybutton2", "restart_style_roster", "Restart Style Roster")


        End Select

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
'Dim sz() As String

    Select Case Control.GetName

'    Case "hotkey_prefs"
'        ' /* the key picker control should return a pair of values separated by a comma.  The first
'        '    value is the set of modifiers; the second value is the keycode of the key pressed */
'
'        sz() = Split(Value, ",")
'        If UBound(sz()) <> 1 Then _
'            Exit Sub
'
'        ' /* we're only interested in the keycode here */
'
'        If sz(1) <> g_ConfigGet("hotkey_prefs") Then
'            If frmAbout.bSetHotkeys(Val(sz(1))) Then
'                ' /* registered okay so store the new keycode */
'                g_ConfigSet "hotkey_prefs", sz(1)
'                g_Debug "TGeneralPage.ControlChanged(): [hotkey_prefs]: hotkey changed to #" & sz(1)
'
'            Else
'                g_Debug "TGeneralPage.ControlChanged(): [hotkey_prefs]: couldn't set hotkey to #" & sz(1), LEMON_LEVEL_WARNING
'                sz(1) = g_ConfigGet("hotkey_prefs")
'
'            End If
'
'        End If
'
'        Control.SetValue CStr(MOD_CONTROL) & "," & sz(1)


'    ' /* [Debug] */
'
'    Case "debug_toolbar"
'        Select Case Val(Value)
'        Case 1
'            ShellExecute 0, "open", "notepad.exe", l3LogPath(), vbNullString, SW_SHOW
'
'        Case 2
'            SOS_invoke New TSOSHandler
'
'        Case 3
'            If g_IsWinXPOrBetter() Then _
'                CoFreeUnusedLibrariesEx 0, 0
'
'        End Select

'    Case "general_toolbar"
'        Select Case Val(Value)
'        Case 1
'            ' /* tray icon */
'            frmAbout.AddTrayIcon
'
'        Case 2
'            ' /* icon cache */
'            If Not (g_NotificationRoster Is Nothing) Then _
'                g_NotificationRoster.EmptyIconCache
'
'        Case 4
'            ' /* restore factory settings */
'            If MsgBox("Are you sure you want to do this?" & vbCrLf & vbCrLf & _
'                      "Your current Snarl and application settings will be renamed with a '.backup' extension" & vbCrLf & _
'                      "in case you want to recover some or all of them at a later date.", _
'                      vbQuestion Or vbYesNo, _
'                      "Restore Factory Settings") = vbYes Then
'
'                uBackupUserConfig
'                ShellExecute 0, "open", g_MakePath(App.Path) & "delayload.exe", IIf(gDebugMode, "-debug", ""), vbNullString, SW_HIDE
'                PostQuitMessage 0
'
'            End If
'
'        End Select


    Case Else
        ' /* other controls - just write value directly */
        g_ConfigSet Control.GetName, Value

'        Select Case Control.GetName
'
'        Case "use_hotkey"
'            ' /* R2.2: we have a separate config entry now */
'            prefskit_SafeEnable Control.Page.Panel, "hotkey_prefs", (Value = "1")
'            frmAbout.bSetHotkeys
'
'        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pmi As OMMenuItem
Dim hWnd As Long

    Select Case Control.GetName()

'    Case "go_tray_icon"

    ' /* debug page */

'    Case "go_lemon"
'        ShellExecute 0, "open", "notepad.exe", l3LogPath(), vbNullString, SW_SHOW
'
'    Case "go_garbage"
'        If g_IsWinXPOrBetter() Then _
'            CoFreeUnusedLibrariesEx 0, 0
'
'    Case "open_config"
'        ShellExecute 0, "open", g_GetPath(g_SettingsPath()), vbNullString, vbNullString, SW_SHOW
'
'    Case "test"
'        If g_IsPressed(VK_SHIFT) Then
'            With New OMMenu
'                .AddItem .CreateItem("net", "Remote")
'                .AddItem .CreateItem("sec", "Secure")
'                .AddItem .CreateItem("pri", "Priority")
'                .AddItem .CreateItem("act", "Action")
'
'                Set pmi = .Track(mPage.Panel.hWnd)
'                If Not (pmi Is Nothing) Then
'                    Select Case pmi.Name
'
'                    Case "pri"
'                        g_PrivateNotify , "Power interruption", "Mains power has been lost.  UPS reports up-time of 23 mins", , "!system-warning", 1
'
'                    Case "net"
'                        g_PrivateNotify , "ukmxfs0407.gsx-pool.uk", "Server ukmxfs0407 has failed to contact time server for over 180 minutes", , "!system-warning", 0, , , SN_NF_REMOTE
'
'                    Case "sec"
'                        g_PrivateNotify , "Message received", "You have a new message from OPS/NKOR/USEMB", , "!system-info", 0, , , SN_NF_SECURE
'
'                    Case "act"
'                        g_PrivateNotify , "Actions", "123", , "!system-info", , , , , True
'
'                    End Select
'
'                End If
'
'            End With
'
'        Else
'            ' /* mimic how we would do it from an external app... */
'            hWnd = FindWindow("w>Snarl", "Snarl")
'            If IsWindow(hWnd) <> 0 Then _
'                SendMessage hWnd, WM_SNARLTEST, 0, ByVal 0&
'
'        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)

End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)

    Set mPage = Page

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

'Private Sub uBackupUserConfig()
'Dim szPath As String
'Dim sz As String
'
'    If g_GetUserFolderPath(szPath) Then
'        szPath = g_MakePath(szPath) & "etc"
'
'        ' /* rename Snarl configs */
'        MoveFile szPath & "\config41.snarl", szPath & "\config41.snarl.backup"
'        MoveFile szPath & "\styledefaults.conf", szPath & "\styledefaults.conf.backup"
'
'        ' /* rename class files */
'        With New CFolderContent2
'            If .SetTo(szPath) Then
'                .Rewind
'                Do While .GetNextFile(sz)
'                    If g_GetExtension(sz, True) = "class" Then _
'                        MoveFile sz, sz & ".backup"
'
'                Loop
'            End If
'        End With
'    End If
'
'End Sub
