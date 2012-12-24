VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TConfigureClassPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TConfigureClassPanel.cls
    '/
    '/  Description:    Provides the floating class configuration panel
    '/
    '/  Created:        9-Sep-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mPanel As BPrefsPanel
Dim mhWndPanel As Long

Dim mSelectedApp As TApp
Dim mSelectedClass As TAlert

Dim mIgnoreChanges As Boolean           ' // when set KPrefsPage_ControlChanged() doesn't update config
Dim WithEvents thePicker As TRedirectPicker
Attribute thePicker.VB_VarHelpID = -1
Dim mSelectedRedirect As String
Dim mSelectedForward As String

Dim WithEvents theAdvPanel As TAdvEventPanel
Attribute theAdvPanel.VB_VarHelpID = -1

Public Event Done()

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    If (mIgnoreChanges) Then _
        Exit Sub

Dim pc As BControl
Dim sz As String

    ' /* update values */

    Select Case Control.GetName

    Case "", "redirection_list"
        Exit Sub

    Case "content_toolbar"
        Set theAdvPanel = New TAdvEventPanel
        theAdvPanel.Go mPanel.hWnd, mSelectedClass
        Exit Sub

    Case "use_same_toggle"
        mSelectedApp.SetUseSameSettingsForAllClasses (Value = "1")
        uUpdateClassList
        Exit Sub

    Case "class_list"
        ' /* get the selected class from the listbox */
        Set mSelectedClass = mSelectedApp.AlertAt(Val(Value))
        uClassChanged


'        If (mselectedclass Is Nothing) Then _
            g_Debug "TAppPage.ControlChanged(): class #" & Value & " for '" & mCurrentApp.Name & "' doesn't exist", LEMON_LEVEL_WARNING

        ' /* if the class configuration panel is open, tell it to update */

'        If Not (theClassPanel Is Nothing) Then _
            theClassPanel.ClassChanged













    Case "net_forward_list"
        mSelectedForward = prefskit_GetItem(Control, "data")
        Exit Sub

    Case "add_remove_redirect"
        If Value = "+" Then
            Set thePicker = New TRedirectPicker
            thePicker.Go mPanel.hWnd, ""
            Exit Sub

        Else
            mSelectedClass.RemRedirector prefskit_GetItemEx(mPanel, "redirection_list", "data", 0)
            If mPanel.Find("redirection_list", pc) Then
                g_UpdateRedirectList pc, mSelectedClass.RedirectList, False
                pc.SetValue CStr(prefskit_GetItemCount(pc))
                
            End If
            Exit Sub

        End If

    Case "duration", "position", "duplicates"
        mSelectedClass.SetValue Control.GetName, CStr(Val(Value) - 1)

    Case "priority"
        Select Case Val(Value)
        Case 1
            ' /* app decides */
            mSelectedClass.SetValue "custom_priority", "0"

        Case 2
            mSelectedClass.SetValue "custom_priority", "1"
            mSelectedClass.SetValue "priority", "-1"

        Case 3
            mSelectedClass.SetValue "custom_priority", "1"
            mSelectedClass.SetValue "priority", "0"

        Case 4
            mSelectedClass.SetValue "custom_priority", "1"
            mSelectedClass.SetValue "priority", "1"

        End Select

    Case "style_name"
        ' /* 'style' combo */
        uUpdateSchemeList

    Case "scheme_name"
        ' /* 'scheme' combo */
        If mPanel.Find("style_name", pc) Then
            sz = prefskit_GetItem(pc)               ' // get the selected style from the style combo
            If sz <> "" Then _
                mSelectedClass.SetStyleAndScheme sz, prefskit_GetItem(Control)

        End If


'    Case "lb>notify"
'        mCurrentForward = Val(Value)
'        mCurrentForwardName = LCase$(Replace$(prefskit_GetItem(Control), ": ", "/"))
'        uSetConditions





    Case Else
'    Case "show_on_screen", "use_style", "custom_timeout", "use_custom_icon", "custom_icon", _
         "use_custom_sound", "custom_sound", "ignore_input", "use_custom_ack", "custom_ack", "log_during_dnd", "block_spam", _
         "override-normal", "override-away", "override-busy", "block_if_foreground", "use_custom_actions", "custom_actions", _
         "log_in_history"

        Debug.Print "TConfigureClassPanel.ControlChanged(): setting " & g_Quote(Control.GetName) & " to " & g_Quote(Value)
        mSelectedClass.SetValue Control.GetName, Value


    End Select

    ' /* post-processing
    
    Select Case Control.GetName()
    Case "custom_timeout"
        Control.SetText IIf(Value = "0", "Sticky", Value & " second" & IIf(Value = "1", "", "s"))

    Case "use_style"
        If (Value = "1") Then _
            mSelectedClass.SetStyleAndScheme uSelectedItem("style_name"), uSelectedItem("scheme_name")

    End Select

    ' /* update control status */

    uSetControls

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim sz As String

    Select Case Control.GetName

    Case "test"
        If Not (mSelectedClass Is Nothing) Then _
            mSelectedClass.App.TestNotification mSelectedClass

    Case "redirection_list"
        mSelectedRedirect = Control.GetValue()
        Set thePicker = New TRedirectPicker
        sz = prefskit_GetItem(Control, "data")
        thePicker.Go mPanel.hWnd, sz, Val(mSelectedClass.RedirectList.ValueOf(sz, , False))

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print Control.GetName & " - " & Notification

    Select Case Control.GetName
    Case "net_forward_list"
        If (Notification = "checked") And (mSelectedForward <> "") Then
            ' /* get item state */
            If prefskit_GetItem(Control, "checked") = "1" Then
                ' /* add the forward */
                mSelectedClass.AddNetForward mSelectedForward
                
            Else
                ' /* remove it */
                mSelectedClass.RemNetForward mSelectedForward

            End If
        End If

    Case "class_list"
        If (Notification = "checked") And (NOTNULL(mSelectedClass)) Then _
            mSelectedClass.SetEnabled (prefskit_GetItem(Control, "checked") = "1")

    End Select

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    EnableWindow mhWndPanel, -1
    g_WindowToFront mhWndPanel, True
    RaiseEvent Done

End Sub

Private Sub KPrefsPanel_Ready()

    Debug.Print "TConfigureClassPanel.Ready"

    uUpdateStyleList
    uUpdateClassList
    uClassChanged

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndOwner As Long, ByRef SelectedApp As TApp)

    If ISNULL(SelectedApp) Then _
        Exit Sub

    Set mSelectedApp = SelectedApp
    mhWndPanel = g_GetTopLevel(hWndOwner)
    Set mPanel = New BPrefsPanel

Dim ppSub As BPrefsPage
Dim pcTabs As BControl
Dim pp As BPrefsPage
Dim pm As CTempMsg
Dim pc As BControl

    With mPanel
        .SetHandler Me
        .SetWindow 1                                            ' // so it doesn't appear during Go()
        .SetWidth 482
        .SetTitle SelectedApp.Name & " Events"

        Set pp = new_BPrefsPage("", , Me)
        With pp
            .SetMargin 0

            Set pm = New CTempMsg
            pm.Add "checkboxes", 1
            pm.Add "item-height", 24&
            Set pc = new_BPrefsControl("listbox", "class_list", "", "", 1, pm)
            pc.SizeTo 0, (24 * 4) + 2
            .Add pc

            .Add new_BPrefsControl("fancytoggle2", "use_same_toggle", "Use the same settings for all events?", , IIf(SelectedApp.UseSameSettingsForAllClasses, "1", "0"))

            Set pm = New CTempMsg
            pm.Add "height", 260&
            Set pcTabs = new_BPrefsControl("tabstrip", "", , , , pm)

            ' /* Appearance */
            Set ppSub = new_BPrefsPage("", , Me)
            With ppSub
                .SetMargin 102
                ' /* display */
                .Add new_BPrefsControl("fancytoggle2", "show_on_screen", "Show on screen?", "Visibility:")
                ' /* style */
                .Add new_BPrefsControl("fancytoggle2", "use_style", "Use a specific style?", "Style:")
                .Add new_BPrefsCombo("style_name", "", "", , 22)
                .Add new_BPrefsCombo("scheme_name", "", "", , 22)
                ' /* position */
                .Add new_BPrefsControl("fancycycle", "position", "Default|Top Left|Top Right|Bottom Left|Bottom Right", "Position:")
                ' /* R2.6: background tint */
                .Add new_BPrefsControl("fancycolourpicker", "colour-tint", , "Tint:")
                .Add new_BPrefsControl("label", "", "Select black for no tint.")
            End With
            BTabStrip_AddPage pcTabs, "Appearance", ppSub

            ' /* Behaviour */
            Set ppSub = new_BPrefsPage("", , Me)
            With ppSub
                .SetMargin 102
                ' /* priority */
                .Add new_BPrefsControl("fancycycle", "priority", "Application Decides|Low|Normal|High", "Priority:")
                ' /* duration */
                .Add new_BPrefsControl("fancycycle", "duration", "Default|Application Decides|Custom", "Duration:")
                .Add prefskit_NewSlider("custom_timeout", "", 0, 60, 10, 64)
                ' /* logging */
                .Add new_BPrefsControl("fancycycle", "log_in_history", "Always|Never|Application Decides", "Logging:")


'                Set pm = New CTempMsg
'                pm.Add "min", 0&
'                pm.Add "max", 60&
'                pm.Add "freq", 10&
'                pm.Add "label_size", 64&
'                .Add new_BPrefsControl("fancyslider", "custom_timeout", , , , pm)

'                ' /* presence overrides */
'                .Add new_BPrefsControl("banner", "", "Presence Overrides")
'                .Add new_BPrefsControl("fancycycle", "override-normal", "Do Default|Log as Missed|Display Sticky|Discard|Display Normally|Display High Priority", "While Active:")
'                .Add new_BPrefsControl("fancycycle", "override-away", "Do Default|Log as Missed|Display Sticky|Discard|Display Normally|Display High Priority", "While Away:")
'                .Add new_BPrefsControl("fancycycle", "override-busy", "Do Default|Log as Missed|Display Sticky|Discard|Display Normally|Display High Priority", "While Busy:")

            End With
            BTabStrip_AddPage pcTabs, "Behaviour", ppSub

'            ' /* Forwarding */
'            Set ppSub = new_BPrefsPage("", , Me)
'            With ppSub
'                .SetMargin 0
'                ' /* network */
'                .Add new_BPrefsControl("label", "", "Forward to the following computers:")
'                Set pm = New CTempMsg
'                pm.Add "item-height", 24&
'                pm.Add "checkboxes", 1&
'                Set pc = new_BPrefsControl("listbox", "net_forward_list", , "", , pm)
'                pc.SizeTo 0, (3 * 24) + 2
'                .Add pc
'                .Add new_BPrefsControl("fancytoggle2", "send_to_subscribers", "Send to subscribers?", "")
'
'                ' /* redirection */
'                .Add new_BPrefsControl("label", "", "Redirect notifications using the following:")
'                Set pm = New CTempMsg
'                pm.Add "plus_minus", 1&
'                pm.Add "item-height", 24&
'                Set pc = new_BPrefsControl("listbox", "redirection_list", "", "", "1", pm)
'                pc.SizeTo 0, (3 * 24) + 2
'                .Add pc
'                .Add new_BPrefsControl("fancyplusminus", "add_remove_redirect")
'
'            End With
'            BTabStrip_AddPage pcTabs, "Forwarding", ppSub

            ' /* Network */
            Set ppSub = new_BPrefsPage("", , Me)
            With ppSub
                .SetMargin 0
                .SetWidth 300
                .Add new_BPrefsControl("label", "", "Forward to the following destinations:")
                Set pm = New CTempMsg
                pm.Add "item-height", 26&
                pm.Add "checkboxes", 1&
                Set pc = new_BPrefsControl("listbox", "net_forward_list", , "", , pm)
                pc.SizeTo 0, (4 * 26) + 2
                .Add pc
'                .Add new_BPrefsControl("label", "", "Forwards are assigned in [Network]->[Forwarding].")
                .Add new_BPrefsControl("fancytoggle2", "send_to_subscribers", "Send to subscribed computers?", "")
                .Add new_BPrefsControl("fancytoggle2", "send_to_redirects", "Send content to assigned Redirects?", "")
                .Add new_BPrefsControl("label", "", "Forwards and Redirects can be created and managed in the [Gateway] page.")

            End With
            BTabStrip_AddPage pcTabs, "Forwarding", ppSub

            ' /* Redirection */
            Set ppSub = new_BPrefsPage("", , Me)
            With ppSub
                .SetMargin 0
                .Add new_BPrefsControl("label", "", "Redirect using the following:")
                Set pm = New CTempMsg
                pm.Add "plus_minus", 1&
                pm.Add "item-height", 26&
                Set pc = new_BPrefsControl("listbox", "redirection_list", "", "", "1", pm)
                pc.SizeTo 0, (6 * 26) + 2
                .Add pc
                .Add new_BPrefsControl("fancyplusminus", "add_remove_redirect")

            End With
            BTabStrip_AddPage pcTabs, "Redirection", ppSub






'                .Add new_BPrefsControl("label", "", "Redirectors can be configured in the AddOns page.")


            ' /* Advanced */
            Set ppSub = new_BPrefsPage("", , Me)
            With ppSub
                .SetMargin 0
'                .Add new_BPrefsControl("banner", "", "Miscellaneous")
                ' /* misc toggles */
                .Add new_BPrefsControl("fancytoggle2", "block_if_foreground", "Suppress notifications if application is foreground?")
                .Add new_BPrefsControl("fancytoggle2", "ignore_input", "Ignore mouse and other input events?", "")
                .Add new_BPrefsControl("label", "", "Enabling this setting makes notifications created using this class invisible to the mouse pointer, allowing access to content beneath them but at the same time preventing any interaction with the notifications themselves.")
'                .Add new_BPrefsControl("fancytoggle2", "block_spam", "Prevent notification flooding?", "")
                .Add new_BPrefsControl("fancytoggle2", "prevent_close", "Prevent the sending application from closing notifications?", "")
                .Add new_BPrefsControl("label", "", "If enabled, the sending application will not be able to programmatically remove notifications.  Note that this may adversely affect the way some applications interact with Snarl.")

                ' /* content */
                .Add new_BPrefsControl("fancytoolbar", "content_toolbar", "Advanced Settings...|", "")

'                ' /* presence */
'                .Add new_BPrefsControl("banner", "", "Presence")
'                .Add new_BPrefsControl("fancycycle", "override-normal", "Do default|Log as missed|Display Sticky|Discard|Display normally|Display High Priority", "When active:")
'                .Add new_BPrefsControl("fancycycle", "override-away", "Do default|Log as missed|Display Sticky|Discard|Display normally|Display High Priority", "While away:")
'                .Add new_BPrefsControl("fancycycle", "override-busy", "Do default|Log as missed|Display Sticky|Discard|Display normally|Display High Priority", "While busy:")

            End With
            BTabStrip_AddPage pcTabs, "Advanced", ppSub

            ' /* add the tabs themselves */
            .Add pcTabs
            ' /* test button */
            .Add new_BPrefsControl("fancybutton2", "test", "Test")

        End With
        .AddPage pp

'        ' /* advanced */
'
'        Set pp = new_BPrefsPage("Advanced", load_image_obj(g_MakePath(App.Path) & "etc\icons\class-ext.png"), Me)
'        With pp
'            .SetMargin 102
'            .SetWidth 300
'            .Add new_BPrefsControl("banner", "", "Content Customisation")
''            .Add new_BPrefsControl("label", "", "The following can be used to replace content created by the application.")
'
'            ' /* R2.3: custom icon */
''            .Add new_BPrefsControl("fancytoggle2", "use_custom_icon", "Use custom icon?")
'            .Add new_BPrefsControl("fancyfilepicker", "custom_icon", , "Icon:", "")
'            ' /* R2.3: custom sound */
''            .Add new_BPrefsControl("fancytoggle2", "use_custom_sound", "Play custom sound?", "")
'            .Add new_BPrefsControl("fancyfilepicker", "custom_sound", , "Sound:")
'            ' /* R2.31: custom ack - renamed in R2.4 DR7 to "Callback" */
''            .Add new_BPrefsControl("fancytoggle2", "use_custom_ack", "Custom callback?", "")
'            .Add new_BPrefsControl("fancyedit", "custom_ack", , "Callback:")
'            ' /* R2.4.2 DR3: custom actions */
''            .Add new_BPrefsControl("fancytoggle2", "use_custom_actions", "Custom actions?", "")
'            .Add new_BPrefsControl("fancyedit", "custom_actions", , "Actions:")
''            .Add new_BPrefsControl("label", "", "Separate multiple actions with semicolons, for example " & g_Quote("open=!open;close=!close"))
'
'            .Add new_BPrefsControl("banner", "", "Redaction")
'            .Add new_BPrefsControl("fancycycle", "redact-when", "Never|Personal or higher|Private or higher|Confidential or higher", "When:")
'            .Add new_BPrefsControl("fancyedit", "redact-with", , "With:")
'
'            ' /* R2.31: test button here too */
'            .Add new_BPrefsSeparator
'            .Add new_BPrefsControl("fancybutton2", "test", "Test")
'
'        End With
'
'        .AddPage pp
'
'            ' /* duplicates */
''            .Add new_BPrefsControl("fancycycle", "duplicates", "Allow|Block|Delay", "Flooding:")
''            ' /* R2.4.2 DR3: custom actions */
''            .Add new_BPrefsControl("fancytoggle2", "use_custom_actions", "Custom actions?", "")
''            .Add new_BPrefsControl("fancyedit", "custom_actions", , "Actions:")
''            .Add new_BPrefsControl("label", "", "<Action>=<command>.  Separate multiple actions with semi-colons.")
''                .Add new_BPrefsControl("fancycycle", "fc>fwd.when", "Never|Always|Away|Do Not Disturb Set", "When:")
''            .Add new_BPrefsControl("label", "", "View-only mode makes the notification invisible to the mouse pointer, allowing you to access content beneath it.  View-only notifications cannot be interacted with.")
'
''            ' /* R2.31: log during DnD */
''            .Add new_BPrefsControl("fancytoggle2", "log_during_dnd", "Always log when I'm busy", "")
'
''            If Not (mselectedclass Is Nothing) Then _
''                sz = mselectedclass.GetValue("custom_timeout")
'
''            ' /* duration */
''            .Add new_BPrefsControl("fancycycle", "duration", "Default|Application Decides|Custom", "Duration:")
''            Set pm = New CTempMsg
''            pm.Add "min", 0&
''            pm.Add "max", 60&
''            pm.Add "freq", 10&
''            pm.Add "label_size", 64&
''            .Add new_BPrefsControl("fancyslider", "custom_timeout", , , , pm)
''
''            ' /* priority */
''            .Add new_BPrefsControl("fancycycle", "priority", "Application Decides|Low|Normal|High", "Priority:")
'
''            ' /* R2.31: duplicates renamed to spamming */
''            .Add new_BPrefsControl("fancytoggle2", "block_spam", "Prevent notification flooding?", "Other:")


        .Go

        ' /* configure window look */

        g_SetWindowIconFromWindowIcon .hWnd, mhWndPanel

        If hWndOwner <> 0 Then
            SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(mPanel.hWnd, GWL_STYLE) And (Not WS_MINIMIZEBOX)
            g_MoveWindow .hWnd, , , , , True, mhWndPanel
            SetWindowLong .hWnd, GWL_HWNDPARENT, mhWndPanel
            EnableWindow mhWndPanel, 0

        Else
            g_MoveWindow .hWnd, , , , True

        End If

        g_ShowWindow .hWnd, True, True

    End With

End Sub

Public Sub Quit()

    RemoveSubClass mhWndPanel

    If Not (mPanel Is Nothing) Then
        RemoveSubClass mPanel.hWnd
        mPanel.Quit

    End If

End Sub

Private Sub uClassChanged()

    If ISNULL(mSelectedClass) Then
        g_Debug "TConfigureClassPanel.uClassChanged(): no class selected", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    mIgnoreChanges = True

    uSetValues
    uSetControls

Dim sz As String

    sz = mSelectedClass.GetValue("custom_timeout")
    prefskit_SafeSetText mPanel, "custom_timeout", IIf(sz = "0", "Sticky", sz & " second" & IIf(sz = "1", "", "s"))

    mIgnoreChanges = False

End Sub

Private Sub uUpdateStyleList()
Dim pc As BControl

    If Not (mPanel.Find("style_name", pc)) Or (g_StyleRoster Is Nothing) Then _
        Exit Sub

    pc.SetText g_StyleRoster.GetStyleList()

'Dim pm As CTempMsg
'Dim sz As String
Dim i As Long

'    Set pm = New CTempMsg
    For i = 1 To g_StyleRoster.CountStyles()
        prefskit_SetItemObject pc, i, "image-object", g_StyleRoster.StyleAt(i).Icon

'        sz = g_StyleRoster.StyleAt(i).IconPath
'        If Not g_Exists(sz) Then _
'            sz = g_MakePath(App.Path) & "etc\icons\style.png"
'
'        pm.Replace "image-file", sz
'        pm.Replace "index", i
'        pc.DoExCmd B_SET_ITEM, pm

    Next i

End Sub

Private Sub uUpdateSchemeList()
Dim pc As BControl

    If Not mPanel.Find("scheme_name", pc) Then _
        Exit Sub

    If (Not (mPanel.Find("style_name", pc))) Or (g_StyleRoster Is Nothing) Then _
        Exit Sub

Dim px As TStyle

    ' /* get the style object */
    Set px = g_StyleRoster.StyleAt(Val(pc.GetValue))

    ' /* get the scheme combo control */
    mPanel.Find "scheme_name", pc

    If (px Is Nothing) Then
        ' /* style not found so assume <Default> */
        g_Debug "TConfigureClassPanel.uUpdateSchemeList(): style not found, assuming <default>..."
        pc.SetText ""
        pc.SetEnabled False
        Exit Sub

    End If

Dim pm As CTempMsg
Dim i As Long

    If px.CountSchemes = 0 Then
        g_Debug "TConfigureClassPanel.uUpdateSchemeList(): style '" & px.Name & "' has no schemes", LEMON_LEVEL_WARNING
        pc.SetText ""

    Else
        pc.SetText px.GetSchemeList()

        ' /* set the icons */
        Set pm = New CTempMsg
        For i = 1 To px.CountSchemes                '// pc.DoExCmd(B_COUNT_ITEMS)
            prefskit_SetItemObject pc, i, "image-object", px.SchemeIconIcon(i)
'            pm.Replace "index", i
'            pm.Replace "image-file", g_MakePath(px.Path) & px.SchemeAt(i) & ".icon"
'            pc.DoExCmd B_SET_ITEM, pm

        Next i

        ' /* select the first item */
        pc.SetValue "1"

    End If

    ' /* enable/disable the combo based on the scheme count */

    If pc.IsEnabled Then _
        pc.SetEnabled (px.CountSchemes > 0)

End Sub

Private Sub uSafeEnable(ByVal Name As String, ByVal Enabled As Boolean)
Dim pc As BControl

    If mPanel.Find(Name, pc) Then
        If pc.IsEnabled <> Enabled Then _
            pc.SetEnabled Enabled

    End If

End Sub

Private Sub uSetValues()
Dim dw As Long

    With mSelectedClass

        ' /* Appearance */
        prefskit_SetValue mPanel, "use_style", .GetValue("use_style")
        prefskit_SetValue mPanel, "show_on_screen", .GetValue("show_on_screen")
        prefskit_SetValue mPanel, "colour-tint", .GetValue("colour-tint")

        ' /* Behaviour */
        dw = Val(.GetValue("duration"))
        prefskit_SetValue mPanel, "duration", CStr(dw + 1)
        prefskit_SetValue mPanel, "custom_timeout", .GetValue("custom_timeout")
        dw = Val(.GetValue("position"))
        prefskit_SetValue mPanel, "position", CStr(dw + 1)
        prefskit_SetValue mPanel, "log_in_history", .GetValue("log_in_history")
'        prefskit_SetValue mPanel, "override-normal", .GetValue("override-normal")
'        prefskit_SetValue mPanel, "override-away", .GetValue("override-away")
'        prefskit_SetValue mPanel, "override-busy", .GetValue("override-busy")

        ' /* Content */
'        prefskit_SetValue mPanel, "redact-when", uRedactionIndex(.GetValue("redact-when"))
'        prefskit_SafeSetText mPanel, "redact-with", .GetValue("redact-with")
'        prefskit_SetValue mPanel, "redact-always", .GetValue("redact-always")

        If .GetValue("custom_priority") = "0" Then
            ' /* app decides */
            prefskit_SetValue mPanel, "priority", "1"

        Else
            Select Case Val(.GetValue("priority"))
            Case Is < 0
                prefskit_SetValue mPanel, "priority", "2"

            Case 0
                prefskit_SetValue mPanel, "priority", "3"

            Case Is > 0
                prefskit_SetValue mPanel, "priority", "4"

            End Select

        End If


    
    
Dim b As Boolean
Dim sz As String
Dim i As Long
Dim j As Long

        b = (.GetValue("use_style") = "1")
        If b Then
            ' /* look it up */
            sz = .GetValue("style")
            i = g_StyleRoster.IndexOf(style_GetStyleName(sz))
            If i > 0 Then
                ' /* look up the scheme */
                j = g_StyleRoster.StyleAt(i).SchemeIndex(style_GetSchemeName(sz))
                Debug.Print "index of " & style_GetSchemeName(sz) & " : " & j
    
            Else
                g_Debug "TAppsPage.uUpdateSubPages(): style '" & sz & "' not found", LEMON_LEVEL_WARNING
    
            End If
    
        End If
    
        If (i = 0) Or (j = 0) Then
            ' /* style and/or scheme not found OR we're using the default */
            sz = g_ConfigGet("default_style")
            i = g_StyleRoster.IndexOf(style_GetStyleName(sz))
            If i > 0 Then _
                j = g_StyleRoster.StyleAt(i).SchemeIndex(style_GetSchemeName(sz))
    
        End If
    
        prefskit_SetValue mPanel, "style_name", CStr(i)
        uUpdateSchemeList
        prefskit_SetValue mPanel, "scheme_name", CStr(j)


        ' /* Content */
''        prefskit_SetValue mPanel, "use_custom_icon", .GetValue("use_custom_icon")
'        prefskit_SafeSetText mPanel, "custom_icon", .GetValue("custom_icon")
''        prefskit_SetValue mPanel, "use_custom_sound", .GetValue("use_custom_sound")
'        prefskit_SafeSetText mPanel, "custom_sound", .GetValue("custom_sound")
'        ' /* R2.31d3 */
''        prefskit_SetValue mPanel, "use_custom_ack", .GetValue("use_custom_ack")
'        prefskit_SafeSetText mPanel, "custom_ack", .GetValue("custom_ack")
'        ' /* R2.4.2 DR3 */
''        prefskit_SetValue mPanel, "use_custom_actions", .GetValue("use_custom_actions")
'        prefskit_SafeSetText mPanel, "custom_actions", .GetValue("custom_actions")

        ' /* Advanced */
        prefskit_SetValue mPanel, "block_if_foreground", .GetValue("block_if_foreground")
        prefskit_SetValue mPanel, "ignore_input", .GetValue("ignore_input")
        prefskit_SetValue mPanel, "block_spam", .GetValue("block_spam")
        prefskit_SetValue mPanel, "prevent_close", .GetValue("prevent_close")


Dim pc As BControl

        ' /* forwarding */

        If mPanel.Find("net_forward_list", pc) Then
            g_SubsRoster.BuildForwarderList pc, .NetForwardList
            mSelectedForward = prefskit_GetItem(pc, "data")         ' // covers bug in prefs kit which doesn't trigger a change

        End If

        prefskit_SetValue mPanel, "send_to_subscribers", .GetValue("send_to_subscribers")
        prefskit_SetValue mPanel, "send_to_redirects", .GetValue("send_to_redirects")

        ' /* redirection */

        If mPanel.Find("redirection_list", pc) Then _
            g_UpdateRedirectList pc, .RedirectList, False


    End With

End Sub

Private Sub uSetControls()
Dim pc As BControl
Dim B2 As Boolean
Dim b As Boolean

    ' /* enable/disable controls */

    If mPanel.Find("show_on_screen", pc) Then
        b = (pc.GetValue = "1")

        uSafeEnable "position", b
        uSafeEnable "priority", b
        uSafeEnable "use_style", b

        uSafeEnable "ignore_input", (g_ConfigGet("ignore_input") = "0")

        ' /* only enable the combos if "show_on_screen" and "use custom style" are enabled */

        B2 = b
        If B2 Then
            If mPanel.Find("use_style", pc) Then _
                B2 = (pc.GetValue = "1")

        End If

        uSafeEnable "style_name", B2
        uSafeEnable "scheme_name", B2

        uSafeEnable "duration", b

        ' /* only enable the timeout slider if SN_ND_CUSTOM is set */

        B2 = b
        If B2 Then
            If mPanel.Find("duration", pc) Then _
                B2 = (pc.GetValue = CStr(SN_ND_CUSTOM + 1))

        End If
        
        uSafeEnable "custom_timeout", B2

    End If

End Sub

Private Function uSelectedItem(ByVal Name As String) As String
Dim pc As BControl

    If mPanel.Find(Name, pc) Then _
        uSelectedItem = prefskit_GetItem(pc)

End Function

Private Sub theAdvPanel_Closed()

    Set theAdvPanel = Nothing

End Sub

Private Sub thePicker_Closed()

    Set thePicker = Nothing

End Sub

Private Sub thePicker_Saved(ByVal Name As String, ByVal Flags As SN_REDIRECTION_FLAGS)

    mSelectedClass.AddRedirector Name, Flags

Dim pc As BControl

    If mPanel.Find("redirection_list", pc) Then
        ' /* refresh the list */
        g_UpdateRedirectList pc, mSelectedClass.RedirectList, False

        ' /* if an entry was previously selected, select it again now otherwise selected the last one (i.e. new one) */
        If mSelectedRedirect = "" Then _
            mSelectedRedirect = CStr(prefskit_GetItemCount(pc))

        pc.SetValue mSelectedRedirect

    End If

    mSelectedRedirect = ""

End Sub

'Public Function uRedactionIndex(ByVal Redaction As Long) As String
'
'    uRedactionIndex = CStr((Redaction \ 16) + 1)
'
'End Function
'
'Public Function uRedactionFromIndex(ByVal Index As Long) As String
'
'    Index = Index - 1
'    uRedactionFromIndex = CStr(Index * 16)
'
'End Function

Private Sub uUpdateClassList()
Dim pc As BControl

    If Not mPanel.Find("class_list", pc) Then _
        Exit Sub

Dim sz As String
Dim i As Long

    With mSelectedApp
        If .CountAlerts() = 0 Then
            ' /* corrupt! */
            g_Debug "TAppsPage.uUpdateClassList(): '" & .Name & "' has no alerts", LEMON_LEVEL_CRITICAL
            pc.SetText ""

        ElseIf (.CountAlerts() = 1) Or (.UseSameSettingsForAllClasses) Then
            ' /* only the '_all' setting */
            pc.SetText "[All notifications]#?0"
            prefskit_SetItem pc, 1, "checked", IIf(.AlertAt(1).IsEnabled, 1&, 0&)
            pc.SetValue "1"

        Else
            ' /* multiple alerts */
            sz = "[Other notifications]#?0|"
            For i = 2 To .CountAlerts()
                sz = sz & .AlertAt(i).Description() & "#?" & CStr(i) & IIf(i < .CountAlerts(), "|", "")

            Next i
            pc.SetText sz

            ' /* set icons and checkboxes */
            For i = 1 To .CountAlerts()
                prefskit_SetItem pc, i, "checked", IIf(.AlertAt(i).IsEnabled, 1&, 0&)
                prefskit_SetItemObject pc, i, "image-object", load_image_obj(.AlertAt(i).RealIconPath)

            Next i
            pc.SetValue "1"

        End If
    End With

End Sub

Public Sub SelectClass(ByVal Class As String)
Dim i As Long

    ' /* select the specified class */

    i = 1
    If Class <> "" Then
        ' /* find the class */
        i = mSelectedApp.IndexOf(Class)
        If i = 0 Then _
            i = 1

    End If

    ' /* select it */
    prefskit_SetValue mPanel, "class_list", CStr(i)

End Sub

Public Sub Refresh()
Dim sz As String

    sz = prefskit_GetValue(mPanel, "class_list")

    uUpdateStyleList
    uUpdateClassList
    uClassChanged

    prefskit_SetValue mPanel, "class_list", sz

End Sub
