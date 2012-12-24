VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAdvancedPage"
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
        .SetMargin 96

        ' /* hotkeys */

        .Add new_BPrefsControl("banner", "", "Hotkeys")
        .Add new_BPrefsControl("fancytoggle2", "use_hotkey", "Use a hotkey to activate Snarl's Preferences?", "", g_ConfigGet("use_hotkey"))
        .Add new_BPrefsControl("key_picker", "hotkey_prefs", , , CStr(MOD_CONTROL) & "," & g_ConfigGet("hotkey_prefs"), , (g_ConfigGet("use_hotkey") = "1"))

        .Add new_BPrefsControl("fancytoggle2", "", "Use a hotkey to activate Snarl's menu?", "", "0", , False)
        .Add new_BPrefsControl("key_picker", "", , , CStr(MOD_WIN) & "," & g_ConfigGet("hotkey_prefs"), , False)
        .Add new_BPrefsControl("label", "", "Press the key you want to use in the boxes above.  Note that the modifiers (the combination of SHIFT and CTRL keys) used are automatically set.")

        ' /* applications */

        .Add new_BPrefsControl("banner", "", "Applications")
        .Add new_BPrefsControl("fancytoggle2", "notify_on_first_register", "Only notify the first time an application registers?", , g_ConfigGet("notify_on_first_register"))

        ' /* presence management */

        .Add new_BPrefsControl("banner", "", "Presence Management")
        .Add new_BPrefsControl("fancycycle", "away_mode", "Log as Missed|Make Sticky|Discard|Display", "When Away:", g_ConfigGet("away_mode"))
        .Add new_BPrefsControl("fancycycle", "busy_mode", "Log as Missed|Make Sticky|Discard|Display", "When Busy:", g_ConfigGet("busy_mode"))

        ' /* other stuff */

'        .Add new_BPrefsControl("banner", "", "System Functions")
'        .Add new_BPrefsControl("fancybutton2", "go_app_manager", "Launch App Manager")
'        .Add new_BPrefsControl("label", "", "The App Manager allows you to control other Snarl applications which don't have their own user interface.")

'        .Add new_BPrefsControl("fancybutton2", "restart_style_roster", "Reload Styles")
'        .Add new_BPrefsControl("label", "", "Forces Snarl to reload all installed styles.  Under normal circumstances you shouldn't need to do this; it's provided for users who are developing their own styles and want to test them without restarting Snarl.")

'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("fancybutton2", "dnd_settings", "Applications", "Do Not Disturb Mode:")

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim sz() As String

    Select Case Control.GetName

    Case "use_hotkey"
        ' /* R2.2: we have a separate config entry now */
        g_ConfigSet Control.GetName, Value
        prefskit_SafeEnable Control.Page.Panel, "hotkey_prefs", (Value = "1")
        frmAbout.bSetHotkeys

    Case "hotkey_prefs"

        ' /* the key picker control should return a pair of values separated by a comma.  The first
        '    value is the set of modifiers; the second value is the keycode of the key pressed */

        sz() = Split(Value, ",")
        If UBound(sz()) <> 1 Then _
            Exit Sub

        ' /* we're only interested in the keycode here */

        If sz(1) <> g_ConfigGet("hotkey_prefs") Then
            If frmAbout.bSetHotkeys(Val(sz(1))) Then
                ' /* registered okay so store the new keycode */
                g_ConfigSet "hotkey_prefs", sz(1)
                g_Debug "TGeneralPage.ControlChanged(): [hotkey_prefs]: hotkey changed to #" & sz(1)

            Else
                g_Debug "TGeneralPage.ControlChanged(): [hotkey_prefs]: couldn't set hotkey to #" & sz(1), LEMON_LEVEL_WARNING
                sz(1) = g_ConfigGet("hotkey_prefs")

            End If

        End If

        Control.SetValue CStr(MOD_CONTROL) & "," & sz(1)


    Case "notify_on_first_register", "away_mode", "busy_mode"
        g_ConfigSet Control.GetName, Value


    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)
Dim pc As BControl

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()
    Case "go_app_manager"
        ShellExecute 0, "open", g_MakePath(App.Path) & "SNARLAPP_Manager.exe", vbNullString, vbNullString, SW_SHOW

    Case "restart_style_roster"
        If Not (g_StyleRoster Is Nothing) Then
            melonLibClose g_StyleRoster
            Sleep 500
'            MsgBox "Click OK when you're ready for the Style Roster to start up", vbInformation Or vbOKOnly, App.Title
            melonLibOpen g_StyleRoster

            If mPage.Panel.Find("installed_styles", pc) Then _
                pc.Notify "update_list", Nothing


        End If

'    Case "dnd_settings"
'        With New TStyleEnginePanel
'            .Go mPage.Panel.hWnd
'
'        End With

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








