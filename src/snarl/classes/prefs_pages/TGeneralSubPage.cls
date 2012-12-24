VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNetSubPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TNetSubPage.cls
    '/
    '/  Description:    Provides the sub-pages within the Network page
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim WithEvents theAddPanel As TAddForwardOrSubscriptionPanel
Attribute theAddPanel.VB_VarHelpID = -1
Dim WithEvents theNCPanel As TStyleDefaultsPanel
Attribute theNCPanel.VB_VarHelpID = -1

Dim mSelMissed As TNotification
Dim mSelHistory As TNotification
Dim mPage As BPrefsPage

Dim mSelectedForward As String              ' // guid of highlighted network forward
Dim mSelectedSubscription As String
Dim mCurrentItem As ConfigSection

Dim mDisplayStyle As TStyle
Dim mDisplayScheme As String
Dim WithEvents mRedirectStyle As TStyle
Attribute mRedirectStyle.VB_VarHelpID = -1
Dim mRedirectScheme As String

Dim mCurrentEngine As TStyleEngine
Dim mhWndOwner As Long

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
Dim pc As BControl
Dim sz As String

    g_ConfigLock

    Select Case mPage.GetName
    Case "sty-display"
        uUpdateDisplayStyleList
        prefskit_SetValue mPage.Panel, "display_styles_list", "1"

    Case "sty-redirect"
        uUpdateRedirectList

    Case "sty-engines"
        uUpdateEngineList

    Case "net-clients"
        uRefreshForwards

    Case "net-subs"
        uRefreshSubscriptions
        uRefreshSubscribers

    End Select

'    uRefreshListeners
'    uRefreshPassThruList




'    If mPage.Find("busy_style", pc) Then
'        ' /* set the icons *
'        g_StyleRoster.SetNonWindowStyleIcons2 pc
'        ' /* select the right item */
'        sz = style_GetSchemeName(g_ConfigGet("busy_style"))
'        pc.SetValue CStr(uGetListIndex(style_GetStyleName(g_ConfigGet("busy_style")), pc.GetText))
'
'        If mPage.Find("busy_scheme_list", pc) Then _
'            pc.SetValue CStr(uGetListIndex(sz, pc.GetText))
'
'    End If
'
'    If mPage.Find("away_style", pc) Then
'        ' /* set the icons *
'        g_StyleRoster.SetNonWindowStyleIcons2 pc
'        ' /* select the right item */
'        sz = style_GetSchemeName(g_ConfigGet("away_style"))
'        pc.SetValue CStr(uGetListIndex(style_GetStyleName(g_ConfigGet("away_style")), pc.GetText))
'
'        If mPage.Find("away_scheme_list", pc) Then _
'            pc.SetValue CStr(uGetListIndex(sz, pc.GetText))
'
'    End If

    g_ConfigUnlock True

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    With mPage

        Select Case mPage.GetName

        ' /* [Options->Basic] */

        Case "gen-basic"
            .SetMargin 48
            .Add new_BPrefsControl("banner", "", "Launch Options")
            .Add new_BPrefsControl("fancytoggle2", "run_on_logon", "Start at login?", "", g_ConfigGet("run_on_logon"))
            .Add new_BPrefsControl("fancytoggle2", "show_msg_on_start", "Show Welcome Message on startup?", "", g_ConfigGet("show_msg_on_start"), pm)
            .Add new_BPrefsControl("fancytoggle2", "auto_update", "Check for updates on launch?", "", g_ConfigGet("auto_update"))
            .Add new_BPrefsControl("fancybutton2", "update_now", "Check now...")
            ' /* hotkeys */
            .Add new_BPrefsControl("banner", "", "Hotkeys")
            .Add new_BPrefsControl("fancytoggle2", "use_hotkey", "Use a hotkey to activate Snarl's Preferences?", "", g_ConfigGet("use_hotkey"))
            .Add new_BPrefsControl("label", "", "Type key you want to use in the box below (the modifiers used are automatically set):")
            .Add new_BPrefsControl("key_picker", "hotkey_prefs", , , CStr(MOD_CONTROL) & "," & g_ConfigGet("hotkey_prefs"), , (g_ConfigGet("use_hotkey") = "1"))
'            .Add new_BPrefsControl("fancytoggle2", "", "Use a hotkey to activate Snarl's menu?", "", "0", , False)
'            .Add new_BPrefsControl("key_picker", "", , , CStr(MOD_WIN) & "," & g_ConfigGet("hotkey_prefs"), , False)
            .Add new_BPrefsControl("fancytoggle2", "use_notification_hotkey", "Use Windows+Esc key combination to close notifications?", , g_ConfigGet("use_notification_hotkey"))
            .Add new_BPrefsControl("label", "", "Windows+Esc closes the most recent notification; Windows+Ctrl+Esc closes all notifications.")

        ' /* [Options->Advanced] */

        Case "gen-advanced"
            .SetMargin 48
            .Add new_BPrefsControl("banner", "", "Miscellaneous")
            .Add new_BPrefsControl("fancytoggle2", "notify_on_first_register", "Notify each time an application registers?", , uSwap(g_ConfigGet("notify_on_first_register")))
            .Add new_BPrefsControl("fancytoggle2", "garbage_collection", "Periodically unregister orphaned applications?", , g_ConfigGet("garbage_collection"))
            .Add new_BPrefsControl("fancytoggle2", "log_only", "Log only (don't display)?", , g_ConfigGet("log_only"))
    '        .Add new_BPrefsControl("fancytoggle2", "sticky_snarls", "Sticky notifications?", , g_ConfigGet("sticky_snarls"))
            ' /* legacy support */
            .Add new_BPrefsControl("banner", "", "Legacy Support")
            .Add new_BPrefsControl("fancytoggle2", "allow_right_clicks", "Allow notification right and middle clicks?", "", g_ConfigGet("allow_right_clicks"))
            .Add new_BPrefsControl("label", "", "This only applies to notifications created using the V42 API; notifications created using a previous version of the API always receive these events.")
                        
            ' /* other */
            .Add new_BPrefsControl("banner", "", "Other")
            .Add new_BPrefsControl("fancytoolbar", "general_toolbar", "Recreate Tray Icon|Empty Icon Cache||Restore Factory Settings")

        ' /* [Options->Presence] */

        Case "gen-presence"
            .SetMargin 48
            ' /* away determination */
            .Add new_BPrefsControl("banner", "", "Assume I'm Away...")
            .Add new_BPrefsControl("label", "", "After the following period of inactivity:")
            Set pm = New CTempMsg
            pm.Add "min", 0&
            pm.Add "max", 30&
            pm.Add "freq", 5&
            pm.Add "label_size", 56&
            .Add new_BPrefsControl("fancyslider", "idle_minutes", "", "", IIf(g_ConfigGet("idle_minutes") = "2", "1", "2"), pm)
            .Add new_BPrefsControl("fancytoggle2", "away_when_locked", "When the computer is locked?", , g_ConfigGet("away_when_locked"))
            .Add new_BPrefsControl("fancytoggle2", "away_when_screensaver", "When the screensaver starts?", , g_ConfigGet("away_when_screensaver"))
            ' /* busy determination */
            .Add new_BPrefsControl("banner", "", "Assume I'm Busy...")
            .Add new_BPrefsControl("fancytoggle2", "away_when_fullscreen", "When the foreground application is fullscreen?", , g_ConfigGet("away_when_fullscreen"))
            ' /* on return */
            .Add new_BPrefsControl("banner", "", "When I'm Back...")
            .Add new_BPrefsControl("fancycycle", "show_missed_notifications", "Do nothing|Tell me if I missed notifications|Show Notification Centre", , g_ConfigGet("show_missed_notifications"))
            .Add new_BPrefsControl("fancybutton2", "do_nc_settings", "Notification Centre Settings...")
'            .Add new_BPrefsControl("fancytoggle2", "show_missed_notification", "Tell me if I missed any notifications?", "", g_ConfigGet("show_missed_notification"))
'            .Add new_BPrefsSeparator()
'            .Add new_BPrefsControl("label", "", "The above settings affect normal priority notifications only; high priority notifications are always displayed and low priority notifications are always discarded while Do Not Disturb mode is enabled.")


        ' /* [Options->Security] */

        Case "gen-security"
            .SetMargin 24
            .Add new_BPrefsControl("banner", "", "Applications")
            .Add new_BPrefsControl("fancytoggle2", "apps_must_register", "Applications must register before creating notifications?", "", g_ConfigGet("apps_must_register"))
            .Add new_BPrefsControl("fancytoggle2", "only_allow_secure_apps", "Only allow password-protected applications?", "", g_ConfigGet("only_allow_secure_apps"))
            .Add new_BPrefsControl("label", "", "Applications can provide a password when registering with Snarl which makes it harder for rogue applications to generate notifications by spoofing a legitimate application.  Note that enabling this setting will prevent existing non-password protected applications from registering.")
            .Add new_BPrefsControl("banner", "", "Notifications")
            .Add new_BPrefsControl("fancytoggle2", "no_callback_urls", "Prevent URLs being used as the default callback?", "", g_ConfigGet("no_callback_urls"))
            .Add new_BPrefsControl("fancytoggle2", "ignore_style_requests", "Ignore requests for specific styles?", "", g_ConfigGet("ignore_style_requests"))
'            .Add new_BPrefsControl("banner", "", "Network")


        ' /* [Options->Debug] */

        Case "gen-debug"
            .SetMargin 48
            .Add new_BPrefsControl("banner", "", "Debugging")
            .Add new_BPrefsControl("fancytoolbar", "debug_toolbar", "View Debug Log|Invoke SOS|Garbage Collection")
            .Add new_BPrefsControl("banner", "", "Configuration")
            .Add new_BPrefsControl("fancybutton2", "open_config", "Open Config Folder")
            .Add new_BPrefsControl("label", "", "Opens the current config folder in Explorer so the various configuration files can be edited manually.")
            .Add new_BPrefsControl("banner", "", "Diagnostics")
            .Add new_BPrefsControl("fancybutton2", "test", "Test Notification")
            .Add new_BPrefsControl("label", "", "Sends a special test message to the Snarl engine which should result in a notification appearing.  This message is sent using the same mechanism a 3rd party application would use and therefore should prove (or otherwise) that the Snarl notification engine is running correctly.")



        Case "sty-display"
            .SetMargin 0
            .Add new_BPrefsControl("label", "", "These are the currently installed styles:")

            Set pm = New CTempMsg
            pm.Add "item-height", 36&
            Set pc = new_BPrefsControl("listbox", "display_styles_list", , , , pm)
            pc.SizeTo 0, (4 * 36) + 2
            .Add pc
        
            .Add new_BPrefsControl("label", "display_schemes_count")
            .Add new_BPrefsControl("fancytoolbar", "display_style_configure_toolbar", "Configure...|")
        
            Set pm = New CTempMsg
            pm.Add "item-height", 24&
            Set pc = new_BPrefsControl("listbox", "display_schemes_list", , , , pm)
            pc.SizeTo 0, (4 * 24) + 2
            .Add pc

            .Add new_BPrefsControl("fancytoolbar", "display_style_toolbar", "Preview|Preview Priority|Make Default")




'        Case "sty-display"
'            .SetMargin 0
'            .Add new_BPrefsControl("label", "", "Display styles generate content which Snarl displays on-screen:")
'
'            Set pm = New CTempMsg
'            pm.Add "item-height", 36&
'            .Add new_BPrefsControl("combo", "display_styles", , , , pm)
'
'            ' /* style description */
'            Set pc = new_BPrefsControl("label", "display_desc", " ")
'            pc.SizeTo 0, 26
'            .Add pc
'
'            ' /* 'the <style> has <n> schemes' label */
'            .Add new_BPrefsControl("label", "display_scheme_label", "   ")
'
'            Set pm = New CTempMsg
'            pm.Add "item-height", 24&
'            .Add new_BPrefsControl("listbox", "display_schemes", , , , pm)
'
'            pm.Add "scale_to_fit", 1&
'            Set pc = new_BPrefsControl("image", "image_preview", "", , , pm)
'            pc.SizeTo 0, 100
'            .Add pc
'
'            .Add new_BPrefsControl("fancytoolbar", "display_toolbar", "Configure...|Preview|Preview Priority|Set as Default||Advanced...")
        
        
        
        Case "sty-redirect"

            .SetMargin 0
            .Add new_BPrefsControl("label", "", "Redirects typically send the notification content to an external source (such as a web forwarding service, a mobile phone, etc.) or a different medium (such as speaking the notification).")

            Set pm = New CTempMsg
            pm.Add "item-height", 36&
            .Add new_BPrefsControl("combo", "redirect_styles", , , , pm)

            ' /* style description */
            Set pc = new_BPrefsControl("label", "redirect_desc", " ")
            pc.SizeTo 0, 26
            .Add pc

            .Add new_BPrefsControl("fancytoolbar", "redirect_configure_toolbar", "Configure...|")

            ' /* 'the <style> has <n> schemes' label */
            .Add new_BPrefsControl("label", "redirect_scheme_label", "   ")

            Set pm = New CTempMsg
            pm.Add "item-height", 24&
            Set pc = new_BPrefsControl("listbox", "redirect_schemes", , , , pm)
            pc.SizeTo 0, (6 * 24) + 2
            .Add pc

            .Add new_BPrefsControl("fancytoolbar", "redirect_toolbar", "Test|Test Priority")
        

        Case "sty-engines"
            .SetMargin 0
            .Add new_BPrefsControl("label", "", "Style Engines provide the Displays and Redirects which Snarl uses to generate the on-screen (or off-screen) notification content.")

            Set pm = New CTempMsg
            pm.Add "item-height", 36&
            Set pc = new_BPrefsControl("listbox", "engine_list", , , , pm)
            pc.SizeTo 0, (4 * 36) + 2 + 12
            .Add pc
        
'            Set pc = new_BPrefsControl("label", "engine_desc")
'            pc.SizeTo 0, 26
            .Add new_BPrefsControl("label", "engine_desc")
        
            Set pm = New CTempMsg
            pm.Add "item-height", 26&
            Set pc = new_BPrefsControl("listbox", "engine_style_list", , , , pm)
            pc.SizeTo 0, (4 * 26) + 2 + 6
            .Add pc
        
        
            .Add new_BPrefsControl("fancytoolbar", "style_engine_toolbar", "Start|Stop|Details|Configure...||Show in Explorer")
'            .Add new_BPrefsControl("fancybutton2", "fb2>install", "Install new Style Engine...")




        Case "net-general"
            .SetMargin 78
            .Add new_BPrefsControl("banner", "", "Receiving")
            .Add new_BPrefsControl("fancytoggle2", "listen_for_snarl", "Listen for incoming Growl or Snarl notifications?", , g_ConfigGet("listen_for_snarl"), pm)
            .Add new_BPrefsControl("fancytoggle2", "listen_for_json", "Listen for incoming JSON notifications?", , g_ConfigGet("listen_for_json"), pm)
            .Add new_BPrefsControl("label", "", "Snarl will listen on ports 9887 and 23053 for incoming notifications, and port 9889 for incoming JSON notifications.  You may need to adjust your firewall settings in order to receive these.")
            .Add new_BPrefsControl("banner", "", "Sending")
'            .Add new_BPrefsControl("fancytoggle2", "include_host_name_when_forwarding", "Include computer name when forwarding notifications?", , g_ConfigGet("include_host_name_when_forwarding"))
            .Add new_BPrefsControl("fancytoggle2", "include_icon_when_forwarding", "Include icon when forwarding notifications?", , g_ConfigGet("include_icon_when_forwarding"))
            .Add new_BPrefsControl("banner", "", "Security")
            .Add new_BPrefsControl("fancytoggle2", "require_password", "Require a password for remote applications?", "", g_ConfigGet("require_password"))
            Set pm = New CTempMsg
            pm.Add "password", 1&
            .Add new_BPrefsControl("fancyedit", "network_password", g_GetPassword(), "Password:", , pm)
            .Add new_BPrefsControl("label", "", "This setting applies to applications registering via SNP3 or GNTP from a remote source.  Note that enabling a password will prevent SNP 1.0 and SNP 2.0 applications from registering.")
            .Add new_BPrefsControl("fancytoggle2", "block_net_control", "Prevent access to system functions via the network?", "", g_ConfigGet("block_net_control"))

    
        Case "net-clients"
            .SetMargin 0
            .Add new_BPrefsControl("label", "", "Snarl can forward notification content to remote computers running Snarl or an SNP or GNTP compatible product.  Forwarders are defined here and then assigned on a per-class basis:")
            Set pm = New CTempMsg
            pm.Add "item-height", 36&
            pm.Add "plus_minus", 1&
            Set pc = new_BPrefsControl("listbox", "net_forward_list", , , , pm)
            pc.SizeTo 0, (7 * 36) + 2
            .Add pc
            .Add new_BPrefsControl("fancyplusminus", "add_remove_forward")
            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "test_forward", "Test")

'            Set pm = New CTempMsg
'            pm.Add "plus_minus", 1&
'            pm.Add "item-height", 36&
'            pm.Add "checkboxes", 1&
'            Set pc = new_BPrefsControl("listbox", "net_forward_list", , , , pm)
'            pc.SizeTo 0, (7 * 36) + 2
'            .Add pc
'            .Add new_BPrefsControl("fancyplusminus", "fpm>forward")


        Case "net-subs"
            .SetMargin 0
'            .Add new_BPrefsControl("banner", "", "Subscriptions")
            .Add new_BPrefsControl("label", "", "Snarl will subscribe to notifications from the following computers:")
            Set pm = New CTempMsg
            pm.Add "plus_minus", 1&
            pm.Add "item-height", 36&
'            pm.Add "checkboxes", 1&
            Set pc = new_BPrefsControl("listbox", "net_subscription_list", , , , pm)
            pc.SizeTo 0, (4 * 36) + 2
            .Add pc
            .Add new_BPrefsControl("fancyplusminus", "add_rem_subscription")

'            .Add new_BPrefsControl("banner", "", "Subscribers")
            .Add new_BPrefsControl("label", "", "The following computers are subscribed to notifications from this computer:")
            Set pm = New CTempMsg
            pm.Add "item-height", 26&
            Set pc = new_BPrefsControl("listbox", "net_subscriber_list", , , , pm)
            pc.SizeTo 0, (3 * 26) + 2
            .Add pc

            .Add new_BPrefsControl("fancytoggle2", "allow_subs", "Allow remote computers to subscribe to notifications?", "", g_ConfigGet("allow_subs"))
            .Add new_BPrefsControl("fancytoggle2", "notify_when_subscriber_added", "Notify me when remote computers subscribe to this computer?", "", g_ConfigGet("notify_when_subscriber_added"))

    
'        Case "net-subscribers"
'            .SetMargin 0
'            .Add new_BPrefsControl("label", "", "The following computers are subscribed to notifications from this computer:")
'
'            Set pm = New CTempMsg
'            pm.Add "item-height", 26&
'            Set pc = new_BPrefsControl("listbox", "net_subscriber_list", , , , pm)
'            pc.SizeTo 0, (5 * 26) + 2
'            .Add pc
'
''            .Add new_BPrefsControl("fancyplusminus", "fpm>forward")
    
    
    
        Case "net-listeners"
            .SetMargin 0
            .Add new_BPrefsControl("label", "", "Snarl will listen on the following URLs:")

            Set pm = New CTempMsg
            pm.Add "plus_minus", 1&
            pm.Add "item-height", 36&
            pm.Add "checkboxes", 1&
            Set pc = new_BPrefsControl("listbox", "listener_list", , , , pm)
            pc.SizeTo 0, 240
            .Add pc
    
            .Add new_BPrefsControl("fancyplusminus", "add_remove_listener")


        Case "his-all"
            .SetMargin 0
            Set pm = New CTempMsg
            pm.Add "checkboxes", 1&
            pm.Add "item-height", 36&
            Set pc = new_BPrefsControl("listbox", "history_list", "", "", "", pm)
            pc.SizeTo 0, (7 * 36) + 2
            .Add pc

            .Add new_BPrefsControl("fancytoolbar", "history_toolbar", "Select All|Deselect All|Invert Selection||Remove Selected")
            .Add new_BPrefsControl("fancybutton2", "history_more", "Actions...")
            .Add new_BPrefsControl("label", "", "Double-click to replay the notification")


        Case "his-missed"
            .SetMargin 0
            Set pm = New CTempMsg
            pm.Add "checkboxes", 1&
            pm.Add "item-height", 36&
            Set pc = new_BPrefsControl("listbox", "missed_list", , , , pm)
            pc.SizeTo 0, (7 * 36) + 2
            .Add pc

            .Add new_BPrefsControl("fancytoolbar", "missed_toolbar", "Select All|Deselect All|Invert Selection||Remove Selected")
            .Add new_BPrefsControl("fancybutton2", "missed_more", "Actions...")
            .Add new_BPrefsControl("label", "", "Double-click to replay the notification")


        ' /* [ Default Style panel ] */

        Case "defs-general"
            .SetMargin 90
            .Add new_BPrefsControl("banner", "", "Size")
            .Add prefskit_NewSlider("style:width", "Width:", 80, 560, 60, , Val(gStyleDefaults.ValueOf("width")))
            Set pc = prefskit_NewSlider("style:height", "Height:", 20, 300, 40, , Val(gStyleDefaults.ValueOf("height")))
            pc.SetEnabled False
            .Add pc
            .Add new_BPrefsControl("banner", "", "Background")
            .Add new_BPrefsControl("fancycolourpicker", "style:background-colour", , "Normal:", gStyleDefaults.ValueOf("background-colour"))
            .Add new_BPrefsControl("fancycolourpicker", "style:background-colour-priority", , "Priority:", gStyleDefaults.ValueOf("background-colour-priority"))

        Case "defs-title"
            .SetMargin 90
            .Add new_BPrefsControl("fancyfontpicker", "style:title-font", gStyleDefaults.ValueOf("title-font"), "Font:", gStyleDefaults.ValueOf("title-font"))
            .Add new_BPrefsControl("fancycolourpicker", "style:title-colour", , "Normal:", gStyleDefaults.ValueOf("title-colour"))
            .Add new_BPrefsControl("fancycolourpicker", "style:title-colour-priority", , "Priority:", gStyleDefaults.ValueOf("title-colour-priority"))
'            .Add new_BPrefsControl("fancycycle", "style:title-weight", "Normal|Bold", "Weight:", gStyleDefaults.ValueOf("title-weight"))
            .Add prefskit_NewSlider("style:title-opacity", "Opacity:", 0, 100, 10, , Val(gStyleDefaults.ValueOf("title-opacity")), "%")
        
        Case "defs-text"
            .SetMargin 90
            .Add new_BPrefsControl("fancyfontpicker", "style:text-font", gStyleDefaults.ValueOf("text-font"), "Font:", gStyleDefaults.ValueOf("text-font"))
            .Add new_BPrefsControl("fancycolourpicker", "style:text-colour", , "Normal:", gStyleDefaults.ValueOf("text-colour"))
            .Add new_BPrefsControl("fancycolourpicker", "style:text-colour-priority", , "Priority:", gStyleDefaults.ValueOf("text-colour-priority"))
'            .Add new_BPrefsControl("fancycycle", "style:text-weight", "Normal|Bold", "Weight:", gStyleDefaults.ValueOf("text-weight"))
            .Add prefskit_NewSlider("style:text-opacity", "Opacity:", 0, 100, 10, , Val(gStyleDefaults.ValueOf("text-opacity")), "%")

        Case "defs-adv"
            .SetMargin 24
            .Add new_BPrefsControl("banner", "", "Visual Aids")
            .Add new_BPrefsControl("fancytoggle2", "style:eyesight", "Enable support for people with visual impairments?", "", gStyleDefaults.ValueOf("eyesight"))
'            .Add new_BPrefsControl("fancytoggle2", "style:animate", "Enable animations?", "", gStyleDefaults.ValueOf("animate"))
        
        
        ' /* [Notification Centre settings panel] */
        
        Case "nc-general"
            .SetMargin 90
            .Add new_BPrefsControl("banner", "", "Colours")
            .Add new_BPrefsControl("fancycolourpicker", "nc-col-background", , "Background:", g_ConfigGet("nc-col-background"))
            .Add new_BPrefsControl("fancycolourpicker", "nc-col-text", , "Text:", g_ConfigGet("nc-col-text"))
            .Add prefskit_NewSlider("nc-opacity-percent", "Opacity:", 5, 100, 5, , g_SafeLong(g_ConfigGet("nc-opacity-percent")), "%")
            .Add new_BPrefsControl("fancybutton2", "nc-restore_defaults", "Restore defaults")





        
        
'            Set pm = New CTempMsg
'            pm.Add "min", 80&
'            pm.Add "max", 560&
'            pm.Add "freq", 60&
'            .Add new_BPrefsControl("fancyslider", "style:width", "", "Width:", gStyleDefaults.ValueOf("width"), pm)
''            Set pm = New CTempMsg
''            pm.Add "min", 20&
''            pm.Add "max", 300&
''            pm.Add "freq", 40&
''            .Add new_BPrefsControl("fancyslider", "style:height", "", "Height:", gStyleDefaults.ValueOf("height"), pm, False)
        
        ' /* [Presence] */

'        Case "pre-active"
'            ' /* Pass-Through settings */
'            .SetMargin 24
'            .Add new_BPrefsControl("banner", "", "When I'm At My Computer")
'            .Add new_BPrefsControl("label", "", "In addition to displaying the notification using the selected style, Snarl will also redirect it to the following:")
'
'            Set pm = New CTempMsg
'            pm.Add "checkboxes", 1
'            pm.Add "item-height", CLng(30)
'            Set pc = new_BPrefsControl("listbox", "global_shadow_list", g_StyleRoster.GetNonWindowStyleList(), "", "", pm)
'            pc.SizeTo 0, (9 * 30) + 2
'            .Add pc
'
'        Case "pre-away"
'            .SetMargin 24
'            .Add new_BPrefsControl("banner", "", "When Away")
'            .Add new_BPrefsControl("fancycycle", "away_mode", "Log notification as missed|Display sticky|Discard notification|Display normally|Display as high priority|Redirect to...", "", g_ConfigGet("away_mode"))
'            .Add new_BPrefsCombo("away_style", g_StyleRoster.GetNonWindowStyleList2(), "", , 24, , , , (g_ConfigGet("away_mode") = "6"))
'            .Add new_BPrefsCombo("away_scheme_list", "", "", , 24, , , , (g_ConfigGet("away_mode") = "6"))
'
''            Set pc = new_BPrefsControl("listbox", "away_scheme_list", , , , , (g_ConfigGet("away_mode") = "6"))
''            .Add pc
'
''            .Add new_BPrefsControl("banner", "", "Assume I'm Away...")
''            .Add new_BPrefsControl("label", "", "After the following period of inactivity:")
''            Set pm = New CTempMsg
''            pm.Add "min", 0&
''            pm.Add "max", 30&
''            pm.Add "freq", 5&
''            pm.Add "label_size", 56&
''            .Add new_BPrefsControl("fancyslider", "idle_minutes", "", "", IIf(g_ConfigGet("idle_minutes") = "2", "1", "2"), pm)
''            .Add new_BPrefsControl("fancytoggle2", "away_when_locked", "When the computer is locked?", , g_ConfigGet("away_when_locked"))
''            .Add new_BPrefsControl("fancytoggle2", "away_when_screensaver", "When the screensaver starts?", , g_ConfigGet("away_when_screensaver"))
''            .Add new_BPrefsControl("label", "", "Note that the above settings only apply to normal priority notifications.  More granular settings are available in the class configuration preferences panel.")
'
'        Case "pre-busy"
'            .SetMargin 24
'            .Add new_BPrefsControl("banner", "", "When Busy")
'            .Add new_BPrefsControl("fancycycle", "busy_mode", "Log notification as missed|Display sticky|Discard notification|Display normally|Display as high priority|Redirect to...", "", g_ConfigGet("busy_mode"))
'            .Add new_BPrefsCombo("busy_style", g_StyleRoster.GetNonWindowStyleList2(), "", , 24, , , , (g_ConfigGet("busy_mode") = "6"))
'            .Add new_BPrefsCombo("busy_scheme_list", "", "", , 24, , , , (g_ConfigGet("busy_mode") = "6"))
'
'''            Set pc = new_BPrefsControl("listbox", "busy_scheme_list", , , , , (g_ConfigGet("away_mode") = "6"))
'''            .Add pc
''
''            .Add new_BPrefsControl("banner", "", "Assume I'm Busy...")
''            .Add new_BPrefsControl("fancytoggle2", "away_when_fullscreen", "When the foreground application is fullscreen?", , g_ConfigGet("away_when_fullscreen"))
''            .Add new_BPrefsControl("label", "", "Note that the above settings only apply to normal priority notifications.  More granular settings are available in the class configuration preferences panel.")
        
        End Select

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

Dim sz As String

    ' /* handle controls from the Default Style Settings panel */

    If g_SafeLeftStr(Control.Page.GetName, 5) = "defs-" Then
        sz = g_SafeRightStr(Control.GetName, Len(Control.GetName) - 6)
        gStyleDefaults.Update sz, Value
        gStyleDefaults.Save
        Exit Sub

    End If

Dim pmi As OMMenuItem
Dim pm As CTempMsg
Dim pc As BControl
Dim i As Long

    Select Case Control.GetName()

    Case "notify_on_first_register"
        ' /* meaning reversed in R2.6 */
        g_ConfigSet Control.GetName, uSwap(Value)
        Exit Sub

    Case "net_forward_list"
        mSelectedForward = prefskit_GetItem(Control, "data")
'        Set mCurrent = g_SubsRoster.Forwarders.SectionAt(g_SubsRoster.Config.FindSection(mSelectedForward))
        Exit Sub

    Case "net_subscription_list"
        mSelectedSubscription = prefskit_GetItem(Control, "data")
        Exit Sub

    Case "display_styles_list"
        sz = prefskit_GetItem(Control, "data")
        Set mDisplayStyle = Nothing
        g_StyleRoster.Find sz, mDisplayStyle
        prefskit_SetItemEx mPage.Panel, "display_style_configure_toolbar", "enabled", 1, IIf(mDisplayStyle.IsConfigurable, 1&, 0&)
        uUpdateDisplaySchemes
        Exit Sub

    Case "display_schemes_list"
        mDisplayScheme = prefskit_GetItem(Control, "data")
        Exit Sub

    Case "display_style_configure_toolbar"
        If Val(Value) = 1 Then _
            mDisplayStyle.ShowPrefs mPage.Panel.hWnd

    Case "display_style_toolbar"
        If NOTNULL(mDisplayStyle) Then
            Select Case Val(Value)
            Case 1
                mDisplayStyle.DoSchemePreview mDisplayScheme, False, 50

            Case 2
                mDisplayStyle.DoSchemePreview mDisplayScheme, True, 50

            Case 3
                g_ConfigSet "default_style", LCase$(mDisplayStyle.Name & "/" & mDisplayScheme)

            End Select
        End If
        Exit Sub

    Case "redirect_styles"
        uRedirectStyleChanged Control
        Exit Sub

    Case "redirect_schemes"
        mRedirectScheme = prefskit_GetItem(Control, "data")
        Exit Sub

    Case "redirect_configure_toolbar"
        uRedirectToolbarClicked Val(Value)
        Exit Sub

    Case "redirect_toolbar"
        uRedirectToolbarClicked Val(Value) + 1
        Exit Sub

    Case "add_remove_forward"
        If Value = "+" Then
            Set theAddPanel = New TAddForwardOrSubscriptionPanel
            theAddPanel.Go mPage.Panel.hWnd, Nothing, True

        Else
            ' /* remove selected entry */
            If g_SubsRoster.Remove(mSelectedForward) Then _
                uRefreshForwards

        End If

    Case "add_rem_subscription"
        If Value = "+" Then
            Set theAddPanel = New TAddForwardOrSubscriptionPanel
            theAddPanel.Go mPage.Panel.hWnd, Nothing, False

        Else
            ' /* remove selected entry */
            If g_SubsRoster.Remove(mSelectedSubscription) Then _
                uRefreshSubscriptions

        End If


'    Case "add_remove_listener"
'        uAddOrRemoveListener (Value = "+")


    ' /* [Options->General]

    '   /* (none defined) */

    ' /* [Options->Advanced]

    Case "hotkey_prefs"
        uSnarlHotkeyChanged Control, Value
        Exit Sub

    Case "general_toolbar"
        uGeneralToolbarClicked Val(Value)
        Exit Sub


    ' /* [Options->Security]

    '   /* (none defined) */

    ' /* [Options->Debug]

    ' /* [Debug] */

    Case "debug_toolbar"
        Select Case Val(Value)
        Case 1
            ShellExecute 0, "open", "notepad.exe", l3LogPath(), vbNullString, SW_SHOW

        Case 2
            SOS_invoke New TSOSHandler

        Case 3
            If g_IsWinXPOrBetter() Then _
                CoFreeUnusedLibrariesEx 0, 0

        End Select
        Exit Sub


    ' /* [History->All] */

    Case "history_list"
        i = g_NotificationRoster.History.IndexOf(prefskit_GetItem(Control, "data"))
        If i Then _
            Set mSelHistory = g_NotificationRoster.History.TagAt(i)

    Case "history_toolbar"
        Select Case Val(Value)
        Case 1
            ' /* select all */
            If Control.Page.Find("history_list", pc) Then
                With pc
                    For i = 1 To Val(.DoExCmd(B_COUNT_ITEMS))
                        prefskit_SetItem pc, i, "checked", 1&

                    Next i

                End With

            End If

        Case 2
            ' /* deselect all */
            If Control.Page.Find("history_list", pc) Then
                With pc
                    For i = 1 To Val(.DoExCmd(B_COUNT_ITEMS))
                        prefskit_SetItem pc, i, "checked", 0&

                    Next i

                End With

            End If

        Case 3
            ' /* invert */
            If Control.Page.Find("history_list", pc) Then
                With pc
                    For i = 1 To Val(.DoExCmd(B_COUNT_ITEMS))
                        prefskit_SetItem pc, i, "checked", IIf(prefskit_GetItem(pc, "checked", i) = "1", 0&, 1&)

                    Next i

                End With
            End If

        Case 5
            ' /* remove */
            If mPage.Panel.Find("history_list", pc) Then
                With pc
                    For i = prefskit_GetItemCount(pc) To 1 Step -1
                        If prefskit_GetItem(pc, "checked", i) = "1" Then _
                            g_NotificationRoster.History.Remove g_NotificationRoster.History.IndexOf(prefskit_GetItem(pc, "data", i))

                    Next i

                    frmAbout.bUpdateHistoryList

                End With
            End If

        End Select



    ' /* [History->Missed] */


    Case "missed_list"
        i = g_NotificationRoster.MissedList.IndexOf(prefskit_GetItem(Control, "data"))
        If i Then _
            Set mSelMissed = g_NotificationRoster.MissedList.TagAt(i)

    Case "missed_toolbar"
        Select Case Val(Value)
        Case 1
            ' /* select all */
            If Control.Page.Find("missed_list", pc) Then
                With pc
                    For i = 1 To Val(.DoExCmd(B_COUNT_ITEMS))
                        prefskit_SetItem pc, i, "checked", 1&

                    Next i

                End With

            End If

        Case 2
            ' /* deselect all */
            If Control.Page.Find("missed_list", pc) Then
                With pc
                    For i = 1 To Val(.DoExCmd(B_COUNT_ITEMS))
                        prefskit_SetItem pc, i, "checked", 0&

                    Next i

                End With

            End If

        Case 3
            ' /* invert */
            If Control.Page.Find("missed_list", pc) Then
                With pc
                    For i = 1 To Val(.DoExCmd(B_COUNT_ITEMS))
                        prefskit_SetItem pc, i, "checked", IIf(prefskit_GetItem(pc, "checked", i) = "1", 0&, 1&)

                    Next i

                End With

            End If

        Case 5
            ' /* remove */
            If mPage.Panel.Find("missed_list", pc) Then
                With pc
                    For i = prefskit_GetItemCount(pc) To 1 Step -1
                        If prefskit_GetItem(pc, "checked", i) = "1" Then _
                            g_NotificationRoster.MissedList.Remove g_NotificationRoster.MissedList.IndexOf(prefskit_GetItem(pc, "data", i))
                    Next i
                    
                    frmAbout.bUpdateMissedList

                End With
            End If

        End Select


    ' /* [Presence->Away] */


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


'    Case "away_style"
'        ' /* away style combo */
'        If mPage.Find("away_scheme_list", pc) Then
'            pc.SetText g_StyleRoster.GetSchemeList(prefskit_GetItem(Control, "data"))
'            pc.SetValue "1"
'
'        End If
'
'    Case "away_scheme_list"
'        If mPage.Find("away_style", pc) Then _
'            g_ConfigSet pc.GetName, prefskit_GetItem(pc, "data") & "/" & prefskit_GetItem(Control, "data")
'
'    Case "busy_style"
'        ' /* busy style combo */
'        If mPage.Find("busy_scheme_list", pc) Then
'            pc.SetText g_StyleRoster.GetSchemeList(prefskit_GetItem(Control, "data"))
'            pc.SetValue "1"
'
'        End If
'
'    Case "busy_scheme_list"
'        If mPage.Find("busy_style", pc) Then _
'            g_ConfigSet pc.GetName, prefskit_GetItem(pc, "data") & "/" & prefskit_GetItem(Control, "data")



    ' /* [AddOns] -> [Style Engines] */

    Case "engine_list"
        uStyleEngineChanged Value
        Exit Sub

    Case "style_engine_toolbar"
        If (mCurrentEngine Is Nothing) Then _
            Exit Sub

        Select Case Val(Value)
        Case 1
            ' /* start */
            If g_StyleRoster.Load(mCurrentEngine.Name, True, True) Then _
                uUpdateEngineList

            uUpdateDisplayStyleList

        Case 2
            ' /* stop */
            sz = mCurrentEngine.Name
            Set mCurrentEngine = Nothing
            If g_StyleRoster.Unload(sz, True) Then _
                uUpdateEngineList

            uUpdateDisplayStyleList

        Case 3
            ' /* details */
            mCurrentEngine.DoDetailsNotification

        Case 4
            ' /* configure... */
            mCurrentEngine.Configure

        Case 6
            ' /* show in Explorer */
            ShellExecute 0, "open", mCurrentEngine.Engine.Path, vbNullString, vbNullString, SW_SHOWNORMAL

        End Select

        Exit Sub












    Case Else
        ' /* other controls - just write value directly */
        g_ConfigSet Control.GetName, Value

    End Select


    ' /* post-processing */

    Select Case Control.GetName()

    Case "listen_for_json"
        frmAbout.EnableJSON (Value = "1")

    Case "listen_for_snarl"
        frmAbout.EnableSNP (Value = "1")

    Case "away_mode"
'        prefskit_SafeEnable mPage.Panel, "away_style", (Value = "6")
'        prefskit_SafeEnable mPage.Panel, "away_scheme_list", (Value = "6")

    Case "busy_mode"
'        prefskit_SafeEnable mPage.Panel, "busy_style", (Value = "6")
'        prefskit_SafeEnable mPage.Panel, "busy_scheme_list", (Value = "6")

    ' /* [NC settings panel] */

    Case "nc-col-background", "nc-col-text", "nc-opacity-percent"
        g_NotificationRoster.NC.Repaint

    ' /* [Options] */

    Case "use_notification_hotkey"
        frmAbout.bSetNotificationHotkey (Value = "1")

    Case "run_on_logon"
        Select Case Value
        Case "1"
            If Not add_registry_startup_item("Snarl", g_MakePath(App.Path) & LCase$(App.EXEName) & ".exe") Then _
                g_Debug "*failed to write to registry startup key", LEMON_LEVEL_CRITICAL

        Case Else
            rem_registry_startup_item "Snarl", g_MakePath(App.Path) & LCase$(App.EXEName) & ".exe"

        End Select

    Case "use_hotkey"
        ' /* R2.2: we have a separate config entry now */
        prefskit_SafeEnable Control.Page.Panel, "hotkey_prefs", (Value = "1")
        frmAbout.bSetHotkeys

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName()

'    Case "display_schemes"
'        uDisplayToolbarClicked 2

    Case "redirect_schemes"
        uRedirectToolbarClicked 2

    Case "do_nc_settings"
        Set theNCPanel = New TStyleDefaultsPanel
        theNCPanel.Go mPage.Panel.hWnd, True

    Case "nc-restore_defaults"
        prefskit_SetValue mPage.Panel, "nc-col-background", CStr(rgba(31, 33, 33))
        prefskit_SetValue mPage.Panel, "nc-col-text", CStr(rgba(255, 255, 255))

    Case "history_list"
        ' /* replay the selected item */
        If NOTNULL(mSelHistory) Then _
            g_NotificationRoster.Add mSelHistory.Info, Nothing, False, True

    Case "missed_list"
        ' /* replay the selected item */
        If NOTNULL(mSelMissed) Then
            mSelMissed.Replay
            frmAbout.bUpdateMissedList

        End If

    Case "history_more", "missed_more"
        uDoHistoryMoreButton Control, (Control.GetName = "history_more")

    Case "open_config"
        ShellExecute 0, "open", g_GetPath(g_SettingsPath()), vbNullString, vbNullString, SW_SHOW

    Case "test"
        uTestButtonClicked


    Case "net_forward_list"
        ' /* edit selected forward */
        Set theAddPanel = New TAddForwardOrSubscriptionPanel
        Set mCurrentItem = g_SubsRoster.Config.SectionAt(g_SubsRoster.Config.FindSection(mSelectedForward))
        theAddPanel.Go mPage.Panel.hWnd, mCurrentItem, True

    Case "net_subscription_list"
        ' /* edit selected subscription */
        Set theAddPanel = New TAddForwardOrSubscriptionPanel
        Set mCurrentItem = g_SubsRoster.Config.SectionAt(g_SubsRoster.Config.FindSection(mSelectedSubscription))
        theAddPanel.Go mPage.Panel.hWnd, mCurrentItem, False


    Case "test_forward"
        
Dim pInfo As T_NOTIFICATION_INFO
Dim ps As ConfigSection

        If g_SubsRoster.Find(mSelectedForward, ps) Then
            ' /* mock-up a basic notification */
            With pInfo
                .Title = "Network Forwarding Test"
                .Text = "Hello from Snarl on " & g_GetComputerName() & "!"
                Set .ClassObj = g_AppRoster.AppAt(g_AppRoster.IndexOfSig(App.ProductName)).AlertAt(1)
                .IconPath = g_MakePath(App.Path) & "etc\icons\good.png"
                .Timeout = -1

            End With
            g_KludgeNotificationInfo pInfo

            If Not g_SubsRoster.ForwardNotification(ps, pInfo) Then _
                MsgBox "Notification forwarding failed.  Check the protocol and host names are valid.", vbExclamation Or vbOKOnly, "Network Forwarder"

        Else

        End If

'    Case "engine_style_list"
'        If NOTNULL(mCurrentEngine) Then
'            MsgBox prefskit_GetItem(Control)
'
'        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
Dim ppd As BPackedData
Dim ps As TSubscriber
Dim i As Long

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

    Select Case Control.GetName()
    Case "net_forward_list", "net_subscription_list"

        If Notification = "update_list" Then
            MsgBox "!! subscriber list !!"

        ElseIf Notification = "refresh" Then
            ' /* refresh forwarder and subscription lists */
            uRefreshForwards

        ElseIf Notification = "checked" Then
'            Debug.Print prefskit_GetItem(Control, "data") & " > " & prefskit_GetItem(Control, "checked")
'            Set ps = g_SubsRoster.List.TagAt(g_SubsRoster.List.IndexOf(prefskit_GetItem(Control, "data")))
'            If NOTNULL(ps) Then
'                ps.SetEnabled (prefskit_GetItem(Control, "checked", i) = "1")
'                uRefreshForwards
'
'            End If

        End If

    Case "net_subscriber_list"
        If Notification = "refresh" Then _
            uRefreshSubscribers

    Case "display_styles_list"
        If Notification = "refresh" Then _
            uUpdateDisplayStyleList

    Case "engine_list"
        If Notification = "refresh" Then _
            uUpdateEngineList


    Case "global_shadow_list"
        If Notification = "checked" Then
            ' /* rebuild selection */
            Set ppd = New BPackedData
            For i = 1 To prefskit_GetItemCount(Control)
                If prefskit_GetItem(Control, "checked", i) = "1" Then _
                    ppd.Add prefskit_GetItem(Control, "data"), "1"

            Next i

            g_ConfigSet "global_shadow_list", ppd.AsString

        End If

    End Select

End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)

    Set mPage = Page

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

Private Sub uRefreshForwards()
Dim pc As BControl

    If mPage.Find("net_forward_list", pc) Then _
        g_SubsRoster.BuildForwarderList pc

End Sub

Private Sub uRefreshSubscriptions()
Dim pc As BControl

    If mPage.Find("net_subscription_list", pc) Then _
        g_SubsRoster.BuildSubscriptionList pc

End Sub

'Private Sub uDoSubscriberMenu(ByRef Control As BControl)
'Dim pi As OMMenuItem
'Dim rc As RECT
'
'    GetWindowRect Control.hWnd, rc
'
'    With New OMMenu
'        .AddItem .CreateItem(CStr(SN_ST_SNP3_FORWARD), "Forward to remote computer...")
'        .AddItem .CreateItem(CStr(SN_ST_SNP3_SUBSCRIPTION), "Subscribe to remote server...")
''        .AddItem .CreateItem("subg", "Subscribe to remote GNTP server...")
'
'        Set pi = .Track(mPage.Panel.hWnd, new_BPoint(rc.Left, rc.Bottom))
'
'    End With
'
'    If (pi Is Nothing) Then _
'        Exit Sub
'
'    Select Case pi.Name
'    Case "subs"
'
'    Case "subg"
'
'    Case "frwd"
'
'    End Select
'
'    Set theAddPanel = New TAddRemoteComputerPanel
'    theAddPanel.Go Control.Page.Panel.hWnd, Val(pi.Name)
'
'End Sub

Private Sub mRedirectStyle_StylePrefsWindowClosed()
Dim pc As BControl

    ' /* the style's configuration window has gone, so let's
    '    refresh the list of schemes it provides */

    If mPage.Find("redirect_schemes", pc) Then _
        KPrefsPage_ControlChanged pc, pc.GetValue

End Sub

Private Sub theAddPanel_Closed()

    Set theAddPanel = Nothing
    Set mCurrentItem = Nothing

End Sub

'Private Sub uAddOrRemoveListener(ByVal Add As Boolean)
'Dim pc As BControl
'Dim sz As String
'Dim i As Long
'
''    If Add Then
''        Set theAddPanel = New TAddRemoteComputerPanel
''        theAddPanel.Go mPage.Panel.hWnd, SN_ST_COMET
''
''    Else
''        ' /* remove selected entry */
''        If mPage.Find("listener_list", pc) Then
''            i = Val(pc.GetValue())
''            If i > 0 Then
''                sz = prefskit_GetItem(pc, "data", i)
''                If g_SubsRoster.RemoveListenerByGuid(sz) Then _
''                    uRefreshListeners
''
''            End If
''        End If
''    End If
'
'End Sub

Private Sub uRefreshListeners()
'Dim pc As BControl
'
'    If Not mPage.Find("listener_list", pc) Then _
'        Exit Sub
'
'Dim pl As TWebListener
'Dim sz As String
'
'    With g_SubsRoster.Listeners
'        .Rewind
'
'        Do While .GetNextTag(pl) = B_OK
'            sz = sz & pl.Name & "#?" & pl.Guid & "#?" & pl.URL & "|"
'
'        Loop
'
'    End With
'
'    sz = g_SafeLeftStr(sz, Len(sz) - 1)
'    pc.SetText sz
'
'    If sz = "" Then _
'        Exit Sub
'
'    ' /* set the icons */
'
'Dim pm As CTempMsg
'Dim i As Long
'
'    Set pm = New CTempMsg
'
'    With g_SubsRoster.Listeners
'        .Rewind
'        Do While .GetNextTag(pl) = B_OK
'            i = i + 1
'            sz = g_MakePath(App.Path) & "etc\icons\no_icon.png"
'
'            pm.Replace "index", i
'            pm.Replace "image-file", sz
''            pm.Replace "checked", IIf(ps.IsEnabled, 1&, 0&)
'            pc.DoExCmd B_SET_ITEM, pm
'
'        Loop
'
'    End With

End Sub

Private Function uGetListIndex(ByVal Name As String, ByVal List As String) As Long

    On Error Resume Next

Dim s() As String

'    MsgBox "looking for " & Name & " in " & List

    s = Split(LCase$(List), "|")
    If UBound(s) < 0 Then _
        Exit Function

Dim r() As String
Dim i As Long

    For i = 0 To UBound(s)
        r = Split(s(i), "#?")
        If r(1) = Name Then
            uGetListIndex = i + 1
            Exit Function

        End If
    Next i

End Function

'Private Sub uRefreshPassThruList()
'
'    ' /* 2.4.2: global shadowing */
'
'Dim ppd As BPackedData
'
'    Set ppd = New BPackedData
'    ppd.SetTo g_ConfigGet("global_shadow_list")
'
'Dim pc As BControl
'
'    If Not mPage.Find("global_shadow_list", pc) Then _
'        Exit Sub
'
'Dim j As Long
'
'    j = prefskit_GetItemCount(pc)
'    If j = 0 Then _
'        Exit Sub
'
'Dim sz As String
'Dim i As Long
'
'    For i = 1 To j
'        If ppd.ValueOf(prefskit_GetItem(pc, "data", i)) = "1" Then _
'            prefskit_SetItem pc, i, "checked", 1&
'
'    Next i
'
'    g_StyleRoster.SetNonWindowStyleIcons pc
'
'End Sub

Private Sub uRefreshSubscribers()
Dim pc As BControl

    If Not mPage.Find("net_subscriber_list", pc) Then _
        Exit Sub

Dim ps As TSubscriber
Dim sz As String

    With g_SubsRoster.Subscribers
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            sz = sz & ps.FormattedName & " using " & UCase$(ps.Protocol) & "#?" & ps.AddressAndPort & "|"

        Loop

    End With

    sz = g_SafeLeftStr(sz, Len(sz) - 1)
    pc.SetText sz

    If sz = "" Then _
        Exit Sub

    ' /* set the icons */

Dim i As Long

    With g_SubsRoster.Subscribers
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            i = i + 1
            prefskit_SetItem pc, i, "image-file", g_MakePath(App.Path) & "etc\icons\protocol-" & ps.Protocol & ".png"

        Loop
    End With

End Sub

Private Sub uUpdateRedirectList()
Dim pc As BControl

    If (Not (mPage.Find("redirect_styles", pc))) Or (g_StyleRoster Is Nothing) Then _
        Exit Sub

Dim px As TStyle
Dim sz As String

    ' /* build the list content */

    With g_StyleRoster
        .Rewind
        Do While .GetNextStyle(px)
            With px
                If px.IsRedirect() Then _
                    sz = sz & .Name & " " & .Release & "#?" & LCase$(.Name) & "#?" & _
                              IIf(.Copyright <> "", .Copyright & " ", "") & _
                              IIf(.URL <> "", "(" & .URL & ")", "") & _
                              "|"

            End With
        Loop

    End With

    ' /* set the combo's content (this also clears the previous content) */

    sz = g_SafeLeftStr(sz, Len(sz) - 1)
    If sz = "" Then
        pc.SetText "<No styles>"
        pc.SetEnabled False

        ' /* also clear the schemes list and the label */

        If mPage.Find("redirect_schemes", pc) Then _
            pc.SetText ""

        If mPage.Find("redirect_scheme_label", pc) Then _
            pc.SetText ""

        Exit Sub

    Else
        pc.SetEnabled True
        pc.SetText sz

    End If

    ' /* set the icons */

Dim j As Long

    With g_StyleRoster
        .Rewind
        Do While .GetNextStyle(px)
            If px.IsRedirect() Then
                j = j + 1
                sz = px.IconPath

                ' /* sort of backwards-compatability/kludge: if a style doesn't appear to have
                '    a specific icon path, look for a default icon */

                If Not g_Exists(sz) Then _
                    sz = g_MakePath(px.Path) & px.Name & ".png"

                ' /* use default icon if custom one not found */

                If Not g_Exists(sz) Then _
                    sz = g_MakePath(App.Path) & "etc\icons\class-fwd.png"

                prefskit_SetItem pc, j, "image-file", sz

            End If
        Loop
    End With

    pc.SetValue "1"

End Sub

Private Sub uRedirectStyleChanged(ByRef pCombo As BControl)

    On Error Resume Next

    Set mRedirectStyle = g_StyleRoster.StyleAt(g_StyleRoster.IndexOf(prefskit_GetItem(pCombo, "data")))

    If ISNULL(mRedirectStyle) Then _
        Exit Sub

Dim pc As BControl

'    If mPage.Find("redirect_scheme_label", pc) Then _
        pc.SetText CStr(mRedirectStyle.CountSchemes) & " scheme" & IIf(mRedirectStyle.CountSchemes = 1, "", "s") & ":"

    If mPage.Find("redirect_configure_toolbar", pc) Then _
        prefskit_SetItem pc, 1, "enabled", IIf((mRedirectStyle.Flags And S_STYLE_IS_CONFIGURABLE), 1&, 0&)

    uUpdateRedirectSchemeList

    If mPage.Find("redirect_desc", pc) Then _
        pc.SetText mRedirectStyle.Description

End Sub

Private Sub uUpdateRedirectSchemeList()
Dim pc As BControl

    If (Not (mPage.Find("redirect_schemes", pc))) Or (mRedirectStyle Is Nothing) Then _
        Exit Sub

    ' /* set the listbox content */

    pc.SetText mRedirectStyle.GetSchemeList(True)

    ' /* set the icons */

Dim i As Long

    With mRedirectStyle
        prefskit_SafeSetText mPage.Panel, "redirect_scheme_label", CStr(.CountSchemes) & " scheme" & IIf(.CountSchemes = 1, "", "s") & ":"
        If .CountSchemes Then
            For i = 1 To .CountSchemes
                prefskit_SetItemObject pc, i, "image-object", .SchemeIconIcon(i)
'                prefskit_SetItem pc, i, "image-file", .SchemeIcon(i)

            Next i
        End If
    End With

    pc.SetValue "1"

End Sub

Private Sub uRedirectToolbarClicked(ByVal Index As Long)

    If ISNULL(mRedirectStyle) Then _
        Exit Sub

    Select Case Index
    Case 1
        ' /* configure style */
        mRedirectStyle.ShowPrefs mPage.Panel.hWnd

    Case 2, 3
        ' /* test or test priority */
        mRedirectStyle.DoSchemePreview mRedirectScheme, (Index = 3), 50

        ' /* check private (for now) flag */
        If (mRedirectStyle.Flags And S_STYLE_REDIRECT_TO_SCREEN) = 0 Then _
            g_PrivateNotify , "Redirect test", "A " & IIf(Index = 3, "priority ", "") & "notification has been created using " & _
                            mRedirectStyle.Name & " (" & mRedirectScheme & ")."

    End Select

End Sub

Private Sub uDoHistoryMoreButton(ByRef Control As BControl, ByVal IsHistoryTab As Boolean)

    g_Debug "TNetSubPage.uDoHistoryMoreButton()", LEMON_LEVEL_PROC_ENTER

Dim pc As BControl

    If Not Control.Page.Find(IIf(IsHistoryTab, "history", "missed") & "_list", pc) Then
        g_Debug "couldn't find list control", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If

Dim pl As BTagList

    If IsHistoryTab Then
        g_Debug "getting history list..."
        Set pl = g_NotificationRoster.History

    Else
        g_Debug "getting missed list..."
        Set pl = g_NotificationRoster.MissedList

    End If

Dim pr As BRect

    Set pr = BW_Frame(Control.hWnd)

Dim pmi As OMMenuItem
Dim pn As TNotification
Dim sz As String
Dim i As Long
Dim l As Long
Dim c As Long

    g_Debug "doing menu..."

    ' /* count number of selected entries */
    For i = 1 To Val(pc.DoExCmd(B_COUNT_ITEMS))
        If prefskit_GetItem(pc, "checked", i) = "1" Then _
            c = c + 1

    Next i

    With New OMMenu
        .AddItem .CreateItem("play", "Replay", , (c > 0))
        .AddItem .CreateItem("copy", "Copy to Clipboard", , (c > 0))

        Set pmi = .Track(mPage.Panel.hWnd, new_BPoint(pr.Left, pr.Bottom))
        If NOTNULL(pmi) Then
            Select Case pmi.Name
            Case "copy"
                g_Debug "copy to clipboard..."
                With pc
                    For i = 1 To Val(.DoExCmd(B_COUNT_ITEMS))
                        If prefskit_GetItem(pc, "checked", i) = "1" Then
                            l = pl.IndexOf(prefskit_GetItem(pc, "data", i))
                            If l Then
                                Set pn = pl.TagAt(l)
                                g_Debug "adding notification #" & prefskit_GetItem(pc, "data", i) & "..."
                                sz = sz & CStr(pn.Info.DateStamp) & " from " & pn.AppName & " using " & pn.Info.ClassObj.Description & vbCrLf & _
                                          pn.Info.Title & vbCrLf & pn.Info.Text & vbCrLf & _
                                          pn.Info.IconPath & vbCrLf & _
                                          pn.Info.DefaultAck & vbCrLf & _
                                          "--" & vbCrLf

                            Else
                                g_Debug "notification #" & prefskit_GetItem(pc, "data", i) & " not found", LEMON_LEVEL_WARNING

                            End If
                        End If
                    Next i
                End With

                g_Debug "pasting to clipboard..."
                Clipboard.Clear
                Clipboard.SetText sz, vbCFText
                g_Debug "done"

            Case "play"
                g_Debug "replay..."
                With pc
                    For i = 1 To Val(.DoExCmd(B_COUNT_ITEMS))
                        If prefskit_GetItem(pc, "checked", i) = "1" Then
                            l = pl.IndexOf(prefskit_GetItem(pc, "data", i))
                            If l Then
                                Set pn = pl.TagAt(l)
                                g_NotificationRoster.Add pn.Info, Nothing, False, True

                            End If
                        End If
                    Next i
                End With
            End Select
        End If

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub uUpdateEngineList()
Dim pc As BControl

    If Not (mPage.Find("engine_list", pc)) Then _
        Exit Sub

Dim pse As TStyleEngine
Dim sz As String
Dim n As Long
Dim i As Long

    n = Max(Val(pc.GetValue), 1)

    If Not (g_StyleRoster Is Nothing) Then
        If g_StyleRoster.CountEngines > 0 Then

            With g_StyleRoster.EngineList
                .Rewind
                Do While .GetNextTag(pse) = B_OK
                    sz = sz & g_Capitalise(g_RemoveExtension(pse.Name)) & "#?0#?"
                    Select Case pse.State
                    Case SN_SE_LOADED
                        sz = sz & IIf(pse.Description <> "", pse.Description, " ")

                    Case SN_SE_NOT_LOADED
                        sz = sz & "Not loaded"

                    Case SN_SE_FAILED
                        sz = sz & "Failed: " & pse.Error

                    End Select
                    
                    sz = sz & "|"

                Loop

            End With

            pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

            ' /* set the icons */
            With g_StyleRoster.EngineList
                .Rewind
                Do While .GetNextTag(pse) = B_OK
                    i = i + 1
                    prefskit_SetItemObject pc, i, "image-object", load_image_obj(g_MakePath(App.Path) & "etc\icons\" & IIf(pse.State = SN_SE_LOADED, "good", "bad") & ".png")

                Loop

            End With

        Else
            pc.SetText "<None>"
            n = 1

        End If

    Else
        pc.SetText "<Style Roster not running>"
        n = 1

    End If

    pc.SetValue CStr(n)

End Sub

Private Sub uUpdateEngineStyleList()
Dim pc As BControl

    If (ISNULL(mCurrentEngine)) Or (Not mPage.Find("engine_style_list", pc)) Then _
        Exit Sub

Dim pl As BTagList
Dim pi As BTagItem
Dim sz As String

    Set pl = mCurrentEngine.StyleList
    With pl
        .Rewind
        Do While .GetNextTag(pi) = B_OK
            sz = sz & pi.Name & "|"

        Loop

    End With

    pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

Dim i As Long

    With pl
        .Rewind
        Do While .GetNextTag(pi) = B_OK
            i = i + 1
            If g_Exists(pi.Value) Then
                sz = pi.Value

            Else
                sz = g_MakePath(App.Path) & "etc\icons\style.png"

            End If
            prefskit_SetItemObject pc, i, "image-object", load_image_obj(sz)

        Loop

    End With

End Sub

Private Sub uSnarlHotkeyChanged(ByRef Control As BControl, ByVal Value As String)
Dim sz() As String

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
            g_Debug "TNetSubPage.uSnarlHotkeyChanged(): [hotkey_prefs]: hotkey changed to #" & sz(1)

        Else
            g_Debug "TNetSubPage.uSnarlHotkeyChanged: [hotkey_prefs]: couldn't set hotkey to #" & sz(1), LEMON_LEVEL_WARNING
            sz(1) = g_ConfigGet("hotkey_prefs")

        End If

    End If

    ' /* update control with the CTRL modifier added */

    Control.SetValue CStr(MOD_CONTROL) & "," & sz(1)

End Sub

Private Sub uStyleEngineChanged(ByVal Value As String)

    If ISNULL(g_StyleRoster) Then _
        Exit Sub

Dim pc As BControl
Dim sz As String

    Set mCurrentEngine = g_StyleRoster.EngineAt(Val(Value))
    If Not (mCurrentEngine Is Nothing) Then
        ' /* configure the toolbar */
        If mPage.Find("style_engine_toolbar", pc) Then
            prefskit_SetItem pc, 1, "enabled", IIf(mCurrentEngine.State = SN_SE_LOADED, 0&, 1&)     ' // "Start"
            prefskit_SetItem pc, 2, "enabled", IIf((mCurrentEngine.State = SN_SE_LOADED) And (Not mCurrentEngine.IsSystem), 1&, 0&)     ' // "Stop"
            prefskit_SetItem pc, 3, "enabled", IIf(mCurrentEngine.State = SN_SE_LOADED, 1&, 0&)     ' // "Details"
            prefskit_SetItem pc, 4, "enabled", mCurrentEngine.IsConfigurable                        ' // "Configure..."
            prefskit_SetItem pc, 6, "enabled", IIf((mCurrentEngine.State = SN_SE_LOADED) And (Not mCurrentEngine.IsSystem), 1&, 0&)     ' // "Show in Explorer"

        End If

        ' /* set the description label */
        
        Select Case mCurrentEngine.State
        Case SN_SE_LOADED
            sz = "Loaded (" & mCurrentEngine.StyleCountString & ")"

        Case SN_SE_NOT_LOADED
            sz = "Not loaded"

        Case SN_SE_FAILED
            sz = "Failed: " & mCurrentEngine.Error

        End Select
        
        prefskit_SafeSetText mPage.Panel, "engine_desc", sz
        uUpdateEngineStyleList

    End If

End Sub

Private Sub uGeneralToolbarClicked(ByVal Index As Long)

    Select Case Index
    Case 1
        ' /* tray icon */
        frmAbout.AddTrayIcon

    Case 2
        ' /* icon cache */
        If Not (g_NotificationRoster Is Nothing) Then _
            g_NotificationRoster.EmptyIconCache

    Case 4
        ' /* restore factory settings */
        If MsgBox("Are you sure you want to do this?" & vbCrLf & vbCrLf & _
                  "Your current Snarl and application settings will be renamed with a '.backup' extension" & vbCrLf & _
                  "in case you want to recover some or all of them at a later date.", _
                  vbQuestion Or vbYesNo, _
                  "Restore Factory Settings") = vbYes Then

            uBackupUserConfig
            ShellExecute 0, "open", g_MakePath(App.Path) & "delayload.exe", IIf(gDebugMode, "-debug", ""), vbNullString, SW_HIDE
            PostQuitMessage 0

        End If

    End Select

End Sub

Private Sub uBackupUserConfig()
Dim szPath As String
Dim sz As String

    If g_GetUserFolderPath(szPath) Then
        szPath = g_MakePath(szPath) & "etc"

        ' /* rename Snarl configs */
        MoveFile szPath & "\config41.snarl", szPath & "\config41.snarl.backup"
        MoveFile szPath & "\styledefaults.conf", szPath & "\styledefaults.conf.backup"

        ' /* rename class files */
        With New CFolderContent2
            If .SetTo(szPath) Then
                .Rewind
                Do While .GetNextFile(sz)
                    If g_GetExtension(sz, True) = "class" Then _
                        MoveFile sz, sz & ".backup"

                Loop
            End If
        End With
    End If

End Sub

Private Sub uTestButtonClicked()
Dim pmi As OMMenuItem
Dim hWnd As Long

    If g_IsPressed(VK_SHIFT) Then
        With New OMMenu
            .AddItem .CreateItem("net", "Remote")
            .AddItem .CreateItem("sec", "Secure")
            .AddItem .CreateItem("pri", "Priority")
            .AddItem .CreateItem("act", "Action")

            Set pmi = .Track(mPage.Panel.hWnd)
            If Not (pmi Is Nothing) Then
                Select Case pmi.Name

                Case "pri"
                    g_PrivateNotify , "Power interruption", "Mains power has been lost.  UPS reports up-time of 23 mins", , "!system-warning", 1
                
                Case "net"
                    g_PrivateNotify , "ukmxfs0407.gsx-pool.uk", "Server ukmxfs0407 has failed to contact time server for over 180 minutes", , "!system-warning", 0, , , SN_NF_REMOTE
                
                Case "sec"
                    g_PrivateNotify , "Message received", "You have a new message from OPS/NKOR/USEMB", , "!system-info", 0, , , SN_NF_SECURE

                Case "act"
                    g_PrivateNotify , "Actions", "123", , "!system-info", , , , , True

                End Select
            End If
        End With

    Else
        ' /* mimic how we would do it from an external app... */
        hWnd = FindWindow("w>Snarl", "Snarl")
        If IsWindow(hWnd) <> 0 Then _
            SendMessage hWnd, WM_SNARLTEST, 0, ByVal 0&

    End If

End Sub

Private Function uSwap(ByVal Value As String) As String

    uSwap = IIf(Value = "1", "0", "1")

End Function

Private Function uUpdateDisplayStyleList() As String
Dim pc As BControl

    If Not (mPage.Panel.Find("display_styles_list", pc)) Then _
        Exit Function

Dim ps As TStyle
Dim sz As String
Dim n As Long
Dim i As Long

    n = Max(Val(pc.GetValue), 1)
    If Not (g_StyleRoster Is Nothing) Then
        With g_StyleRoster
            If .CountStyles > 0 Then
                .Rewind
                Do While .GetNextStyle(ps)
                    If Not ps.IsRedirect() Then _
                        sz = sz & ps.NameEx & "#?" & ps.Name & "#?" & IIf(ps.Description = "", IIf(ps.Copyright <> "", ps.Copyright, " "), ps.Description) & "|"

                Loop

                pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

                ' /* set the icons */
                .Rewind
                Do While .GetNextStyle(ps)
                    If Not ps.IsRedirect Then
                        i = i + 1
                        prefskit_SetItemObject pc, i, "image-object", ps.Icon

'                        prefskit_SetItem pc, i, "image-file", ps.SafeIconPath

                    End If
                Loop

            Else
                pc.SetText "<None>"
                n = 1

            End If
        End With
    Else
        pc.SetText "<Style Roster not available>"
        n = 1

    End If

    pc.SetValue CStr(n)

End Function

Private Sub uUpdateDisplaySchemes()
Dim pc As BControl

    If Not (mPage.Panel.Find("display_schemes_list", pc)) Then _
        Exit Sub

Dim sz As String
Dim i As Long

    If NOTNULL(mDisplayStyle) Then
        With mDisplayStyle
            If .CountSchemes Then
                For i = 1 To .CountSchemes
                    sz = sz & .SchemeAt(i) & "#?" & LCase$(.SafeSchemeAt(i)) & "|"

                Next i

                pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

                For i = 1 To .CountSchemes
                    prefskit_SetItemObject pc, i, "image-object", .SchemeIconIcon(i)
'                    prefskit_SetItem pc, i, "image-file", .SchemeIcon(i)

                Next i

                pc.SetValue "1"

            Else
                pc.SetText "<None>"

            End If

            prefskit_SafeSetText mPage.Panel, "display_schemes_count", CStr(.CountSchemes) & " scheme" & IIf(.CountSchemes = 1, "", "s") & ".  Provider: " & .Engine.Name

        End With

    Else
        pc.SetText "<None>"

    End If

End Sub

Private Sub theAddPanel_Saved(ByVal Name As String, ByVal Hosts As String, ByVal Protocol As String, ByVal ItemType As String, ByVal Password As String)
Dim ps As ConfigSection

    ' /* sort out the config file */

    If ISNULL(mCurrentItem) Then
        ' /* create a new section */
        With g_SubsRoster.Config
            Set ps = .AddSectionObj(g_CreateGUID(True))
            ps.Add "type", ItemType
            ps.Add "name", Name
            ps.Add "hosts", Hosts
            ps.Add "protocol", Protocol
            ps.Add "password", Password

        End With
    Else
        ' /* update existing section */
        With mCurrentItem
            .Update "name", Name
            .Update "hosts", Hosts
            .Update "protocol", Protocol
            .Update "password", Password

        End With

    End If

    g_SubsRoster.Config.Save

    ' /* sort out rosters */

    If ItemType = "forwarder" Then
        uRefreshForwards

    Else
        ' /* subscription */
        If ISNULL(mCurrentItem) Then _
            g_SubsRoster.AddSubscription Name, Hosts, Protocol, ps.Name, Password

        uRefreshSubscriptions
    
    End If

End Sub
