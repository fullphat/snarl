VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TGeneralPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

    With mPage
'        .Add new_BPrefsControl("banner", "", "Launch Options")

        .SetMargin 48

        .Add new_BPrefsControl("banner", "", "Launch Options")

        Set pm = New CTempMsg
        pm.Add "text", "Start at login?"
        pm.Add "align", 1
        pm.Add "inset_by", 0
        .Add new_BPrefsControl("fullphat/fancytoggle", "run_on_logon", , "", g_ConfigGet("run_on_logon"), pm)
'        .Add new_BPrefsControl("fancybutton2", "start", "Stop Snarl", "")
'        .Add new_BPrefsControl("label", "lbl_running", "Snarl is running", "")
'        .Add new_BPrefsControl("separator", "")
        .Add new_BPrefsControl("fancytoggle2", "auto_update", "Check for updates on launch?", "", g_ConfigGet("auto_update"))
        .Add new_BPrefsControl("fancybutton2", "update_now", "Check now...")

        ' /* away mode */

'Dim nIdle As Long

        .Add new_BPrefsControl("banner", "", "Presence")
        .Add new_BPrefsControl("label", "", "Enable away mode after the following period of inactivity:")

'        nIdle = ROUNDUP(Val(g_ConfigGet("idle_timeout")) / 300)
'        If (nIdle < 0) Or (nIdle > 3) Then _
'            nIdle = 0

        Set pm = New CTempMsg
        pm.Add "min", 0&
        pm.Add "max", 30&
        pm.Add "freq", 5&
        pm.Add "label_size", 56&
        .Add new_BPrefsControl("fancyslider", "idle_minutes", "", "", IIf(g_ConfigGet("idle_minutes") = "2", "1", "2"), pm)

        .Add new_BPrefsControl("fancytoggle2", "away_when_locked", "Enable away mode when computer is locked?", , g_ConfigGet("away_when_locked"))
        .Add new_BPrefsControl("fancytoggle2", "away_when_fullscreen", "Enable away mode when the foreground application is fullscreen?", , g_ConfigGet("away_when_fullscreen"))

        ' /* miscellaneous */

        .Add new_BPrefsControl("banner", "", "Miscellaneous")
        .Add new_BPrefsControl("fancytoggle2", "show_msg_on_start", "Show message on startup?", "", g_ConfigGet("show_msg_on_start"), pm)
'        .Add new_BPrefsControl("fancytoggle2", "sticky_snarls", "Sticky notifications?", , g_ConfigGet("sticky_snarls"))
        .Add new_BPrefsControl("fancytoggle2", "log_only", "Log only (don't display)?", , g_ConfigGet("log_only"))


'        .Add new_BPrefsControl("fancycycle", "idle_timeout", "Never|5 Minutes|10 Minutes|15 Minutes", "", CStr(nIdle + 1))

''        .Add new_BPrefsControl("fancytoggle2", "auto_sticky_on_screensaver", "Enable sticky notifications when screensaver starts?", "", g_ConfigGet("auto_sticky_on_screensaver"))
'
''        ' /* hotkey */
''
'''        .Add new_BPrefsControl("separator", "")
''
''        .Add new_BPrefsControl("banner", "", "Hotkeys")
''
''        .Add new_BPrefsControl("fancytoggle2", "use_hotkey", "Use a hotkey to activate Snarl's Preferences?", "", IIf(gPrefs.use_hotkey, "1", "0"), pm)
''
''        .Add new_BPrefsControl("key_picker", "hotkey_prefs", , , CStr(MOD_CONTROL) & "," & CStr(gPrefs.hotkey_prefs), , gPrefs.use_hotkey)
''        .Add new_BPrefsControl("label", "", "Press the key you want to use in the box the above.  Note that while you can select the key which activates Snarl, the modifiers (the combination of SHIFT and CTRL keys) used are automatically set.                            ")



'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("fancybutton2", "go_app_manager", "Launch App Manager", "App Manager:")
'        .Add new_BPrefsControl("label", "", "The App Manager allows you to control other Snarl applications which don't have their own user interface.")
'
''        Set pc = new_BPrefsControl("label", "")
''        pc.SizeTo 0, 64
''        .Add pc
''
''        Set pm = New CTempMsg
''        pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\snarl.png"
''        pm.Add "image-height", 32
''        pm.Add "valign", "centre"
''        .Add new_BPrefsControl("labelex", "", "Snarl " & App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ") " & App.LegalCopyright, , , pm)

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pc As BControl

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()

    Case "auto_sticky_on_screensaver", "notify_on_first_register", "auto_update", "sticky_snarls", _
         "log_only", "show_msg_on_start", "run_on_logon", "away_when_locked", "away_when_fullscreen"

        g_ConfigSet Control.GetName, Value

        If Control.GetName = "run_on_logon" Then _
            g_SetAutoRun2

        Exit Sub


    Case "idle_minutes"

        Select Case Val(Value)
        Case 0
            Control.SetText "Never"

        Case 1
            Control.SetText "1 min"

        Case Else
            Control.SetText Value & " mins"

        End Select

        g_ConfigSet Control.GetName, Value
        Exit Sub


    Case Else
        Exit Sub

    End Select

    g_WriteConfig

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)
Dim pc As BControl

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()
'    Case "start"
'        ' /* toggle running state */
'        g_SetRunning Not g_IsRunning
'        Control.SetText IIf(g_IsRunning, "Stop Snarl", "Start Snarl")
''        If Control.Page.Find("lbl_running", pc) Then _
'            pc.SetText IIf(g_IsRunning, "Snarl is running", "Snarl is stopped")
'
'    Case "go_app_manager"
'        ShellExecute 0, "open", g_MakePath(App.Path) & "SNARLAPP_Manager.exe", vbNullString, vbNullString, SW_SHOW

    Case "update_now"
        g_DoManualUpdateCheck

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
'    mPage.setmargin 0

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()

End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)

End Sub

Private Function KPrefsPage_hWnd() As Long

End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)

End Sub






