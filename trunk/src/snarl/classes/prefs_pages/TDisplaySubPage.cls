VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDisplaySubPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TDisplaySubPage.cls
    '/
    '/  Description:    Provides the two sub-pages within the Display page
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Name As String

Dim mPage As BPrefsPage
Dim mStyle As TStyle
Dim mThemes As BTagList

Dim mDisplayStyle As TStyle
Dim mDisplayScheme As String

Dim WithEvents thePicker As TRedirectPicker
Attribute thePicker.VB_VarHelpID = -1
Dim mSelected As String

Implements KPrefsPage

'Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
'Dim pc As BControl
'Dim ps As TStyle
'
'    On Error Resume Next
'
'    If uMsg = WM_DESTROY Then
'
'        ' /* config window has gone */
'
'        window_subclass hWnd, Nothing
'        EnableWindow mPage.Panel.hWnd, -1
'
''        ' /* it's likely that the schemes will have changes so refresh them now.  This
''        '    is _really_ ugly code and needs to be sorted out once a better style
''        '    type library has been developed */
''
'''        If Not (g_StyleRoster Is Nothing) Then
'''            If g_StyleRoster.RefreshStyle(mStyle.Info.Name, ps) Then _
'''                LSet mStyle = ps
'''
'''        End If
''
''        uUpdateDisplaySchemeList
'''        uUpdateRedirectSchemeList
'
'        g_ShowWindow mPage.Panel.hWnd, True, True
'
'
'        If mPage.Find("display_schemes", pc) Then _
'            KPrefsPage_ControlChanged pc, pc.GetValue
'
''        If mPage.Find("redirect_schemes", pc) Then _
''            KPrefsPage_ControlChanged pc, pc.GetValue
'
'    End If
'
'End Function

Private Sub KPrefsPage_AllAttached()

    g_Debug "[Notifications Page].AllAttached(): ", LEMON_LEVEL_PROC_ENTER

    g_Debug "locking config..."
    g_ConfigLock

Dim pc As BControl

    If mPage.Find("default_duration", pc) Then _
        pc.SetValue g_ConfigGet("default_duration")


    ' /* refresh style list */
    g_Debug "loading displayable styles list..."
    uUpdateDisplayList True


    ' /* global redirect list */
    If mPage.Panel.Find("redirection_list", pc) Then _
        g_UpdateRedirectList pc, gGlobalRedirectList, True

    g_ConfigUnlock True
    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg
Dim sz As String
Dim b As Boolean

    With mPage
        Select Case .GetName()
        Case "appearance1"
            ' /* general */
            .Add new_BPrefsCombo("default_style_list", "", "Style:", , 36)
            .Add new_BPrefsCombo("default_scheme_list", "", "Scheme:", , 26)
            .Add new_BPrefsControl("fancytoolbar", "display_toolbar", "Configure...|Make Default")
            .Add prefskit_NewImage("image_preview", "Preview:", True, , 80)
            .Add prefskit_NewSlider("global_opacity", "Opacity:", 1, 100, 10, , Val(g_ConfigGet("global_opacity")), "%")
            uGetIconThemes
            .Add new_BPrefsControl("fancycycle", "icon_theme", uGetIconThemeList(), "Emblems:", CStr(uIconThemeIndex()), , (mThemes.CountItems > 1))
            .Add new_BPrefsControl("fancytoggle2", "show_timestamp", "Show timestamp and application name while pointer is over notification?", "", g_ConfigGet("show_timestamp"))
            .Add new_BPrefsControl("fancytoggle2", "", "Enable support for visual aids?", "", g_ConfigGet("visual_aids"))
        
        Case "appearance2"
            ' /* layout */
            .Add new_BPrefsControl("banner", "", "Placement")
            b = uGetScreens(sz)
            If Not b Then _
                sz = "<Primary Screen>|"

            .Add new_BPrefsControl("fancycycle", "default_screen", sz, "Screen:", g_ConfigGet("default_screen"), , b)
            .Add new_BPrefsControl("fancycycle", "default_position", "Top Left#?0|Top Right#?1|Bottom Left#?2|Bottom Right#?3", "Position:", g_ConfigGet("default_position"))
            .Add prefskit_NewSlider("margin_spacing", "Spacing:", -20, 20, 5, , Val(g_ConfigGet("margin_spacing")))
            .Add new_BPrefsControl("fancycycle", "notification_scaling", "50%|100%|150%|  200%  ", "Scaling:", uScalingToIndex())
            ' /* dropshadow */
            .Add new_BPrefsControl("banner", "", "Drop Shadow")
            .Add prefskit_NewSlider("dropshadow_strength", "Strength:", 0, 100, 10, , Val(g_ConfigGet("dropshadow_strength")), "%")
            .Add prefskit_NewSlider("dropshadow_size", "Size:", 0, 20, 2)

        Case "behaviour"
'            .Add new_BPrefsControl("banner", "", "General")
            .Add prefskit_NewSlider("default_duration", "Duration:", 0, 60, 5, 50, -1)
            .Add new_BPrefsControl("fancycycle", "style.overflow_limit", "5 lines|7 lines|15 lines", "Limit text to:", uOverflowLimit())
            .Add new_BPrefsControl("fancytoggle2", "use_style_icons", "Use style-provided icons (where possible) over Snarl icons?", "", g_ConfigGet("use_style_icons"))
'            .Add new_BPrefsControl("banner", "", "Interaction")
            .Add new_BPrefsControl("fancytoggle2", "ignore_input", "Ignore mouse and other input events?", "Interaction:", g_ConfigGet("ignore_input"))
            .Add new_BPrefsControl("label", "", "Notifications cannot be interacted with and are dimmed while the mouse pointer is over them.")
            .Add new_BPrefsControl("fancytoggle2", "callback_as_button", "Display default callback as a button?", "", g_ConfigGet("callback_as_button"))
            .Add new_BPrefsControl("label", "", "Disable this setting to allow invoking a notification by clicking on its content.")
            .Add new_BPrefsControl("banner", "", "Presence")
            .Add new_BPrefsControl("fancycycle", "away_mode", "Log as missed|Display sticky|Discard notification|Display normally|Make high priority", "When Away:", g_ConfigGet("away_mode"))
            .Add new_BPrefsControl("fancycycle", "busy_mode", "Log as missed|Display sticky|Discard notification|Display normally|Make high priority", "When Busy:", g_ConfigGet("busy_mode"))

'            .Add new_BPrefsControl("banner", "", "Notification Overlay")

        Case "sounds"
            ' /* sounds */
            .Add new_BPrefsControl("banner", "", "General")
            .Add new_BPrefsControl("fancytoggle2", "enable_sounds", "Enable sounds?", "", g_ConfigGet("enable_sounds"))
            .Add new_BPrefsControl("fancytoggle2", "use_style_sounds", "Use style-provided sounds (if available)?", , g_ConfigGet("use_style_sounds"))
            .Add new_BPrefsControl("fancytoggle2", "prefer_style_sounds", "Style-provided sounds take priority over app-provided sounds?", , g_ConfigGet("prefer_style_sounds"))
'        .Add new_BPrefsControl("label", "", "If enabled Snarl will use the Style sounds even if an application specifies its own.")
            .Add new_BPrefsControl("banner", "", "Default Sounds")
            .Add new_BPrefsControl("fancyfilepicker", "default_notification_sound", g_ConfigGet("default_notification_sound"), "Normal:")
            .Add new_BPrefsControl("fancyfilepicker", "default_priority_sound", g_ConfigGet("default_priority_sound"), "Priority:")
            .Add new_BPrefsControl("label", "", "Leave blank for no default sound or specify a system sound by prefixing with '!'.  System sounds are stored HKEY_CURRENT_USER\AppEvents\EventLabels.")


        Case "redirection"
            ' /* redirection */
            .SetMargin 0
'            .Add new_BPrefsControl("banner", "", "Redirection")
            .Add new_BPrefsControl("label", "", "In addition to the on-screen display, you can also redirect notification content based on your current presence.  The Redirects themselves can be configured in the AddOns page.")
            Set pm = New CTempMsg
            pm.Add "plus_minus", 1&
            pm.Add "item-height", 36&
            Set pc = new_BPrefsControl("listbox", "redirection_list", "", "", "1", pm)
            pc.SizeTo 0, (7 * 36) + 2
            .Add pc
            .Add new_BPrefsControl("fancyplusminus", "add_remove_redirect")


        Case "advanced"
            ' /* advanced */
            .SetMargin 24
'            .Add new_BPrefsControl("banner", "", "Presence")
'            .Add new_BPrefsControl("fancycycle", "away_mode", "Log as missed|Display sticky|Discard notification|Display normally|Make high priority", "When Away:", g_ConfigGet("away_mode"))
'            .Add new_BPrefsControl("fancycycle", "busy_mode", "Log as missed|Display sticky|Discard notification|Display normally|Make high priority", "When Busy:", g_ConfigGet("busy_mode"))

            .Add new_BPrefsControl("banner", "", "Content Detection")
            .Add new_BPrefsControl("fancytoggle2", "auto_detect_url", "Automatically detect URLs in a notification?", "", g_ConfigGet("auto_detect_url"))
            .Add new_BPrefsControl("label", "", "If the notification's text is a URL and the notification has no default callback, Snarl will assign the URL as the default callback.")
'            .Add new_BPrefsControl("label", "", "IMPORTANT: Care should be exercised when enabling this setting as a rogue application may create a notification with a malicious URL.  Clicking the notification would then cause your browser to access that URL.")

'            .Add new_BPrefsControl("banner", "", "Visual Aids")
'            .Add new_BPrefsControl("fancytoggle2", "", "Enable support for colour blindness?", "", g_ConfigGet(""), , False)

'            .Add new_BPrefsControl("fancytoggle2", "", "Notification", "", g_ConfigGet(""), , False)
'            .Add new_BPrefsControl("label", "", "Notification scaling:")
    
    
'        ' /* style defaults */
'
'        Case "sty"
'            With mPage
'                .Add new_BPrefsControl("banner", "", "Notification Title")
'                .Add new_BPrefsControl("fancyfontpicker", "style:title-font", gStyleDefaults.ValueOf("title-font"), "Font:", gStyleDefaults.ValueOf("title-font"))
'                .Add new_BPrefsControl("fancycolourpicker", "style:title-colour", , "Colour:", gStyleDefaults.ValueOf("title-colour"))
'    '            .Add new_BPrefsControl("fancycycle", "style:title-weight", "Normal|Bold", "Weight:", gStyleDefaults.ValueOf("title-weight"))
''                Set pm = New CTempMsg
''                With pm
''                    .Add "min", 0&
''                    .Add "max", 100&
''                    .Add "freq", 10&
''                    .Add "suffix", "%"
''
''                End With
''                .Add new_BPrefsControl("fancyslider", "style:title-opacity", "", "Opacity:", gStyleDefaults.ValueOf("title-opacity"), pm)
'                .Add prefskit_NewSlider("style:title-opacity", "Opacity:", 0, 100, 10, , Val(gStyleDefaults.ValueOf("title-opacity")), "%")
'                .Add prefskit_NewSlider("style:text-opacity", "Opacity:", 0, 100, 10, , Val(gStyleDefaults.ValueOf("text-opacity")), "%")
'
'                .Add new_BPrefsControl("banner", "", "Notification Text")
'                .Add new_BPrefsControl("fancyfontpicker", "style:text-font", gStyleDefaults.ValueOf("text-font"), "Font:", gStyleDefaults.ValueOf("text-font"))
'                .Add new_BPrefsControl("fancycolourpicker", "style:text-colour", , "Colour:", gStyleDefaults.ValueOf("text-colour"))
'    '            .Add new_BPrefsControl("fancycycle", "style:text-weight", "Normal|Bold", "Weight:", gStyleDefaults.ValueOf("text-weight"))
''                Set pm = New CTempMsg
''                With pm
''                    .Add "min", 0&
''                    .Add "max", 100&
''                    .Add "freq", 10&
''                    .Add "suffix", "%"
''
''                End With
''                .Add new_BPrefsControl("fancyslider", "style:text-opacity", "", "Opacity:", gStyleDefaults.ValueOf("text-opacity"), pm)
'                .Add prefskit_NewSlider("style:text-opacity", "Opacity:", 0, 100, 10, , Val(gStyleDefaults.ValueOf("text-opacity")), "%")
'                .Add new_BPrefsSeparator
'                .Add new_BPrefsControl("label", "", "These settings only apply to the default style, however other styles may also (and are encouraged to) support these settings.")
'
'            End With
    
        End Select
    End With

        ' /* spacing */

'        mPage.Add new_BPrefsSeparator()
'        mPage.Add new_BPrefsControl("banner", "", "Other")
'        mPage.Add new_BPrefsCombo("c>icon_theme", "", "Icon Theme:", , 28)
'        uUpdateIconThemes

        ' /* font smoothing */

'        mPage.Add new_BPrefsControl("banner", "", "Font Smoothing")
'        mPage.Add new_BPrefsControl("fancycycle", "font_smoothing", "MelonType#?0|None#?1|Antialiased#?2|ClearType#?3|Windows Default#?4", "Mode:", CStr(gPrefs.font_smoothing + 1))
'
'        Set pm = New CTempMsg
'        pm.Add "min", 1&
'        pm.Add "max", 100&
'        pm.Add "freq", 10&
'        pm.Add "suffix", "%"
'        mPage.Add new_BPrefsControl("fancyslider", "melontype_contrast", "", "Contrast:", CStr(gPrefs.melontype_contrast), pm)


End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

Dim pc As BControl
Dim pm As CTempMsg
Dim sz As String
Dim dw As Long
Dim i As Long

    Select Case Control.GetName()
    Case ""
        Exit Sub

    ' /* appearance */

    Case "default_style_list"
        uDisplayStyleChanged Control
        Exit Sub

    Case "default_scheme_list"
        uDisplaySchemeChanged Val(Value)
        Exit Sub

    Case "display_toolbar"
        Select Case Val(Value)
        Case 1
            If NOTNULL(mDisplayStyle) Then _
                 mDisplayStyle.ShowPrefs mPage.Panel.hWnd

        Case 2
            g_ConfigSet "default_style", LCase$(mDisplayStyle.Name & "/" & mDisplayScheme)

        End Select
        Exit Sub


    ' /* notification settings */


    Case "default_style"
        If (g_StyleRoster Is Nothing) Then _
            Exit Sub


    Case "default_scheme"
        MsgBox "!default_scheme!"
'        ' /* 'scheme' combo on 'visuals' page */
'        Set pm = New CTempMsg
'        pm.Add "what", "text"
'
'        If Control.Page.Find("default_style", pc) Then
'            ' /* get the style name */
'            sz = pc.DoExCmd(B_GET_ITEM, pm)
'            If sz <> "" Then
'                ' /* append the style name */
'                sz = LCase$(sz & "/" & Control.DoExCmd(B_GET_ITEM, pm))
'                ' /* set it */
'                g_ConfigSet "default_style", sz
'                Debug.Print "*** default scheme changed to '" & sz & "' ***"
'
'            End If
'        End If


    Case "icon_theme"
        g_ConfigSet Control.GetName, mThemes.TagAt(Val(Value)).Value
        g_LoadIconTheme


    Case "style.overflow_limit"
        Select Case Value
        Case "1"
            g_ConfigSet Control.GetName, "5"

        Case "2"
            g_ConfigSet Control.GetName, "7"

        Case "3"
            g_ConfigSet Control.GetName, "15"

        End Select


    Case "notification_scaling"
        Select Case Value
        Case "1"
            sz = "0.5"

        Case "2"
            sz = "1"

        Case "3"
            sz = "1.5"

        Case "4"
            sz = "2"

        End Select

        g_ConfigSet "scaling", sz



    ' /* redirection */

    Case "add_remove_redirect"
        If Value = "+" Then
            ' /* display the redirect picker window */
            Set thePicker = New TRedirectPicker
            thePicker.Go g_GetTopLevel(mPage.Panel.hWnd), ""

        Else
            ' /* remove the selected redirect */
            If mPage.Panel.Find("redirection_list", pc) Then
                i = Val(pc.GetValue)
                g_RemGlobalRedirect prefskit_GetItem(pc, "data", 0)
                ' /* refresh the list */
                g_UpdateRedirectList pc, gGlobalRedirectList, True
                ' /* and restore the selection */
                dw = prefskit_GetItemCount(pc)
                If i > dw Then
                    pc.SetValue CStr(dw)

                Else
                    pc.SetValue CStr(i)

                End If
            End If
        End If
        Exit Sub


    Case Else
        Debug.Print "TDisplaySubPage.ControlChanged(): setting " & g_Quote(Control.GetName) & " to " & g_Quote(Value)
        g_ConfigSet Control.GetName, Value
        
'    Case "default_position", "melontype_contrast", "global_opacity", "margin_spacing", "dropshadow_strength", "dropshadow_size", _
         "default_notification_sound", "default_priority_sound", "enable_sounds", "use_style_sounds", "prefer_style_sounds", _
         "use_style_icons", "show_timestamp", "auto_detect_url", "default_screen"

    End Select

    ' /* post-processing */

    Select Case Control.GetName()

    Case "default_duration"
        Select Case Val(Value)
        Case 0
            Control.SetText "Sticky"

        Case 1
            Control.SetText "1 sec"

        Case Else
            Control.SetText Value & " secs"

        End Select


'    Case "font_smoothing"
'        gPrefs.font_smoothing = Val(Control.GetValue()) - 1
'        If mPage.Find("melontype_contrast", pc) Then _
'            pc.SetEnabled (Val(Value) = 1)
'
'        fUpdate = True


'        mStyle.Info.Name = ""
'        g_StyleRoster.StyleAt Val(Value), mStyle
'        uUpdateSchemeList


'    Case "ft>use_dropshadow"
'        gPrefs.use_dropshadow = (Val(Value) <> 0)
'        prefskit_SafeEnable mPage.Panel, "dropshadow_strength", (Value = "1")
'        prefskit_SafeEnable mPage.Panel, "dropshadow_size", (Value = "1")

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    If Control.GetName = "redirection_list" Then
        mSelected = Control.GetValue()
        Set thePicker = New TRedirectPicker
        thePicker.Go g_GetTopLevel(mPage.Panel.hWnd), prefskit_GetItem(Control, "data"), Val(gGlobalRedirectList.ValueOf(prefskit_GetItem(Control, "data"), , False))

    End If

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)

    Select Case Control.GetName

    Case "default_style_list"
        If Notification = "refresh" Then _
            uUpdateDisplayList True

    End Select

End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 96

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

'Private Sub uUpdateStyleList()
'Dim pc As BControl
'
'    If (Not (mPage.Find("default_style", pc))) Or (g_StyleRoster Is Nothing) Then _
'        Exit Sub
'
'Dim sz As String
'Dim b As Boolean
'
'    pc.SetText g_StyleRoster.GetStyleList(b)
'    pc.SetEnabled b
''    pc.SetValue "1"
'
'End Sub
'
'Private Sub uUpdateSchemeList()
'Dim pc As BControl
'
'    If (Not (mPage.Find("default_scheme", pc))) Or (mStyle.Name = "") Or (g_StyleRoster Is Nothing) Then _
'        Exit Sub
'
'    ' /* set it */
'
''    pc.SetText g_StyleRoster.GetSchemeList(mStyle)
''    pc.SetEnabled True
''    pc.SetValue "1"
'
'End Sub




'Private Sub uUpdateIconThemes()
'
'    If (mPage Is Nothing) Then _
'        Exit Sub
'
'Dim pc As BControl
'Dim sz As String
'Dim i As Long
'Dim n As Long
'
'    If Not (mPage.Find("c>icon_theme", pc)) Then _
'        Exit Sub
'
'    sz = "(Default Theme)|"
'
'    If gIconThemes Then
'        For i = 1 To gIconThemes
'            If LCase$(gIconTheme(i).Name) = LCase$(gPrefs.icon_theme) Then _
'                n = i
'
'            sz = sz & gIconTheme(i).Name & "|"
'
'        Next i
'
'    End If
'
'    pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)
'    pc.SetValue CStr(n + 1)
'
'    ' /* set the icons */
'
'Dim pm As CTempMsg
'
'    Set pm = New CTempMsg
'    pm.Add "index", 1
'    pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\snarl.png"
'    pc.DoExCmd B_SET_ITEM, pm
'
'
'    If gIconThemes Then
'        For i = 1 To gIconThemes
'            Set pm = New CTempMsg
'            pm.Add "index", i + 1
'            pm.Add "image-file", gIconTheme(i).IconFile
'            pc.DoExCmd B_SET_ITEM, pm
'
'        Next i
'
'    End If
'
'
'End Sub

Private Function uGetIconThemeList() As String

    If (mThemes Is Nothing) Then
        g_Debug "TDisplaySubPage.uGetIconThemeList(): list is null", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pt As BTagItem

    With mThemes
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            uGetIconThemeList = uGetIconThemeList & pt.Name & "|"

        Loop
        
    End With

'    uGetIconThemeList = g_SafeLeftStr(uGetIconThemeList, Len(uGetIconThemeList) - 1)
    g_Debug "TDisplaySubPage.uGetIconThemeList(): list is: " & uGetIconThemeList

End Function

Private Function uIconThemeIndex() As Long
Dim sz As String

    uIconThemeIndex = 1

    sz = g_ConfigGet("icon_theme")
    If sz = "" Then _
        Exit Function

    If (mThemes Is Nothing) Then
        g_Debug "TDisplaySubPage.uIconThemeIndex(): list is null", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pt As BTagItem
Dim i As Long

    With mThemes
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            i = i + 1
            If pt.Value = sz Then
                uIconThemeIndex = i
                Exit Function
            
            End If

        Loop

    End With

End Function

Private Function uTranslateIconTheme(ByVal Path As String) As String

    Path = g_RemoveExtension(g_FilenameFromPath(Path))

Dim s() As String

    ' /* R2.6: use of the "@" splitter is now deprecated - a new way to define the author
    '    and copyright information will be implemented */

    s = Split(Path, "@")
    If UBound(s) = 1 Then
        uTranslateIconTheme = s(1) & " by " & s(0)

    Else
        uTranslateIconTheme = Path

    End If

End Function

Private Sub uGetIconThemes()
Dim sz As String

    g_Debug "TDisplaySubPage.uGetIconThemes()", LEMON_LEVEL_PROC_ENTER

    ' /* get icon themes */

    Set mThemes = new_BTagList()
    mThemes.Add new_BTagItem("[Default]", "")

    With New CFolderContent2
        If .SetTo(g_MakePath(App.Path) & "etc\icons") Then
            .Rewind
            Do While .GetNextFile(sz)
                If (g_IsFolder(sz)) And (g_GetExtension(sz, True) = "icons") Then _
                    mThemes.Add new_BTagItem(uTranslateIconTheme(sz), g_RemoveExtension(g_FilenameFromPath(sz)))

            Loop

        Else
            g_Debug "couldn't find \etc\icons", LEMON_LEVEL_CRITICAL

        End If

    End With

    If mThemes.CountItems < 2 Then _
        g_Debug "Only the default theme available", LEMON_LEVEL_WARNING

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Function uOverflowLimit() As String
Dim n As Long

    ' /* translate the number of lines into a cyclebutton value */

    n = g_SafeLong(g_ConfigGet("style.overflow_limit"))

    Select Case n
    Case 5
        uOverflowLimit = "1"

    Case 15
        uOverflowLimit = "3"

    Case Else
        uOverflowLimit = "2"

    End Select

End Function

Private Function uGetScreens(ByRef List As String) As Boolean
Dim sz As String
Dim pmi As T_MONITOR_INFO
Dim c As Long
Dim i As Long

    c = g_CountMonitors()
    If c > 0 Then
        uGetScreens = (c > 1)

        For i = 1 To c
            g_GetMonitorInfo i, pmi
            sz = sz & "Screen " & CStr(i) & IIf(pmi.IsPrimary, " (Primary)", "") & "|"

        Next i

        List = sz

    End If

End Function

Private Sub uUpdateDisplayList(Optional ByVal SelectDefault As Boolean)
Dim pc As BControl

    g_Debug "TDisplaySubPage.uUpdateDisplayList()", LEMON_LEVEL_PROC_ENTER
    If (Not (mPage.Find("default_style_list", pc))) Or (g_StyleRoster Is Nothing) Then
        g_Debug "missing control/roster", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If

Dim px As TStyle
Dim sz As String

    ' /* build the list content */
    With g_StyleRoster
        .Rewind
        Do While .GetNextStyle(px)
            With px
                If Not px.IsRedirect() Then _
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
        g_Debug "no styles", LEMON_LEVEL_WARNING Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    Else
        pc.SetEnabled True
        pc.SetText sz

    End If

Dim j As Long

    ' /* set the icons */
    With g_StyleRoster
        .Rewind
        Do While .GetNextStyle(px)
            If Not px.IsRedirect() Then
                j = j + 1
                sz = px.IconPath
                If Not g_Exists(sz) Then _
                    sz = g_MakePath(px.Path) & px.Name & ".png"

                ' /* use default icon if custom one not found */
                If Not g_Exists(sz) Then _
                    sz = g_MakePath(App.Path) & "etc\icons\style.png"

                prefskit_SetItem pc, j, "image-file", sz

            End If
        Loop
    End With

Dim i As Long

    If SelectDefault Then
        ' /* make sure current default style is selected */
        sz = style_GetStyleName(g_ConfigGet("default_style"))
        i = prefskit_FindInList(pc, sz)
        If i Then
            g_Debug "selecting style..."
            pc.SetValue CStr(i)
            If mPage.Find("default_scheme_list", pc) Then
                sz = style_GetSchemeName(g_ConfigGet("default_style"))
                i = prefskit_FindInList(pc, sz)
                If i Then
                    g_Debug "selecting scheme..."
                    pc.SetValue CStr(i)

                Else
                    g_Debug "can't find scheme " & g_Quote(sz) & " in list", LEMON_LEVEL_WARNING
                    pc.SetValue "1"
    
                End If
            Else
                g_Debug "scheme list control not found", LEMON_LEVEL_CRITICAL
            
            End If
        Else
            g_Debug "can't find style " & g_Quote(sz) & " in list", LEMON_LEVEL_WARNING
            pc.SetValue CStr(prefskit_FindInList(pc, "default"))

        End If
    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub uDisplayStyleChanged(ByRef pCombo As BControl)

    On Error Resume Next

    Set mDisplayStyle = g_StyleRoster.StyleAt(g_StyleRoster.IndexOf(prefskit_GetItem(pCombo, "data")))
    If ISNULL(mDisplayStyle) Then _
        Exit Sub

Dim pc As BControl

    If mPage.Find("display_scheme_label", pc) Then _
        pc.SetText "This style has " & CStr(mDisplayStyle.CountSchemes) & " scheme" & IIf(mDisplayStyle.CountSchemes = 1, "", "s") & ":"

    If mPage.Find("display_toolbar", pc) Then _
        prefskit_SetItem pc, 1, "enabled", IIf((mDisplayStyle.Flags And S_STYLE_IS_CONFIGURABLE), 1&, 0&)

    uUpdateDisplaySchemeList

End Sub

Private Sub uDisplaySchemeChanged(ByVal NewIndex As Long)

    If ISNULL(mDisplayStyle) Then _
        Exit Sub

    If mDisplayStyle.Name = "" Then _
        Exit Sub

    mDisplayScheme = mDisplayStyle.SchemeAt(NewIndex)
    If mDisplayScheme = "<Default>" Then _
        mDisplayScheme = ""

Dim sz As String

    sz = g_MakePath(mDisplayStyle.Path) & mDisplayScheme & ".preview"
    If Not g_Exists(sz) Then
        If uCreatePreviewImage() Then
            sz = uMakePreviewName()

        Else
            sz = g_MakePath(mDisplayStyle.Path) & "style.preview"

        End If

    End If

    ' /* if no preview image, then use default */

    If Not g_Exists(sz) Then _
        sz = g_MakePath(App.Path) & "etc\icons\blank.png"

Dim pc As BControl

    If mPage.Find("image_preview", pc) Then _
        prefskit_SetItem pc, , "image-file", sz


    frmAbout.bSetPrevewStyle LCase$(mDisplayStyle.Name & "/" & mDisplayScheme)

    ' /* update config */
        
'    If Not g_ConfigIsLocked Then _
        g_ConfigSet "default_style", LCase$(mDisplayStyle.Name & "/" & mDisplayScheme)

End Sub

Private Sub uUpdateDisplaySchemeList()
Dim pc As BControl

    If (Not (mPage.Find("default_scheme_list", pc))) Or (mDisplayStyle Is Nothing) Then _
        Exit Sub

Dim i As Long

    ' /* set the listbox content */
    pc.SetText mDisplayStyle.GetSchemeList(True)

    ' /* set the icons */
    With mDisplayStyle
        If .CountSchemes Then
            For i = 1 To .CountSchemes
                prefskit_SetItemObject pc, i, "image-object", .SchemeIconIcon(i)

            Next i
        End If

        pc.SetEnabled .CountSchemes > 1

    End With

    pc.SetValue "1"

End Sub

Private Function uCreatePreviewImage() As Boolean

    If (mDisplayStyle Is Nothing) Then _
        Exit Function

    If (mDisplayStyle.Flags And S_STYLE_IS_WINDOWLESS) Then _
        Exit Function

    DeleteFile uMakePreviewName()

Dim pInfo As T_NOTIFICATION_INFO

    With pInfo
        .Title = "Scheme Preview"
        .Text = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do..." ' mScheme
        .Timeout = 10
        .IconPath = mDisplayStyle.IconPath
        .StyleName = mDisplayStyle.Name
        .SchemeName = LCase$(mDisplayScheme)
        ' /* R2.4 DR7 */
        Set .ClassObj = New TAlert
        ' /* R2.4 DR8 */
        Set .Actions = new_BTagList()

        .OriginalContent = "value-date-packed::" & Format$(Now(), "YYYYMMDDHHNNSS")

    End With

    If (mDisplayStyle.Flags And S_STYLE_V42_CONTENT) Then _
        g_KludgeNotificationInfo pInfo

Dim pw As CSnarlWindow

    Set pw = New CSnarlWindow
    If pw.Create(pInfo, mDisplayStyle.CreateInstance(), mDisplayStyle.Flags, g_GetStylePath(mDisplayStyle.Name)) = M_OK Then
        pw.CaptureToFile uMakePreviewName()
        pw.Remove
        uCreatePreviewImage = True

    Else
        Debug.Print "create failed"

    End If

End Function

Private Function uMakePreviewName() As String

    If NOTNULL(mDisplayStyle) Then _
        uMakePreviewName = g_MakeValidPath(g_MakePath(mDisplayStyle.Path) & mDisplayStyle.Name & "-" & mDisplayScheme & ".prvw")
 
End Function

'Private Function uDoConfigureStyleWindow(ByRef pStyle As TStyle) As Boolean
'
'    If (pStyle.Engine Is Nothing) Then _
'        Exit Function
'
'Dim hWndConfig As Long
'
'    hWndConfig = pStyle.Engine.GetConfigWindow(pStyle.Name)
'    If IsWindow(hWndConfig) = 0 Then
'        g_Debug "TStylePage.ControlChanged(): ConfigRequested() failed (" & g_HexStr(hWndConfig) & ")"
''        MsgBox "This style cannot be configured.", vbInformation Or vbOKOnly, App.Title
'        Exit Function
'
'    End If
'
'    ' /* valid window */
'
'Dim rcConfig As RECT
'Dim rc As RECT
''Dim dw As Long
'
'    window_subclass hWndConfig, Me
'
'
''    dw = GetWindowLong(hWndConfig, GWL_STYLE)
''    dw = dw And (Not WS_MINIMIZEBOX)
''    SetWindowLong hWndConfig, GWL_STYLE, dw
'
'    SetWindowLong hWndConfig, GWL_HWNDPARENT, mPage.Panel.hWnd
'
'    GetWindowRect mPage.Panel.hWnd, rc
'    GetWindowRect hWndConfig, rcConfig
'    g_RectNormalise rcConfig
'
'    g_MoveWindow hWndConfig, _
'                 rc.Left + Fix(((rc.Right - rc.Left) - rcConfig.Right) / 2), _
'                 rc.Top + Fix(((rc.Bottom - rc.Top) - rcConfig.Bottom) / 2)
'
'    g_ShowWindow hWndConfig, True, True
'
'    EnableWindow mPage.Panel.hWnd, 0
'
'    uDoConfigureStyleWindow = True
'
'End Function

Private Function uScalingToIndex() As Long

    Select Case Val(g_ConfigGet("scaling"))
    Case 0.5
        uScalingToIndex = 1

    Case 1.5
        uScalingToIndex = 3

    Case 2
        uScalingToIndex = 4

    Case Else
        uScalingToIndex = 2

    End Select

End Function

Private Sub thePicker_Closed()

    Set thePicker = Nothing

End Sub

Private Sub thePicker_Saved(ByVal Name As String, ByVal Flags As SN_REDIRECTION_FLAGS)

    If gGlobalRedirectList.IndexOf(Name, False) = 0 Then
        g_AddGlobalRedirect Name, Flags

    Else
        gGlobalRedirectList.Update Name, CStr(Flags), False
    
    End If

Dim pc As BControl

    If mPage.Panel.Find("redirection_list", pc) Then
        g_UpdateRedirectList pc, gGlobalRedirectList, True

        If mSelected = "" Then _
            mSelected = CStr(prefskit_GetItemCount(pc))

        pc.SetValue mSelected

    End If

    mSelected = ""

End Sub


'            Set pm = New CTempMsg
'            pm.Add "min", 0&
'            pm.Add "max", 100&
'            pm.Add "freq", 10&
'            pm.Add "suffix", "%"
'            .Add new_BPrefsControl("fancyslider", "dropshadow_strength", "", "Strength:", g_ConfigGet("dropshadow_strength"), pm)
'            Set pm = New CTempMsg
'            pm.Add "min", 0&
'            pm.Add "max", 20&
'            pm.Add "freq", 2&
'            .Add new_BPrefsControl("fancyslider", "dropshadow_size", "", "Size:", g_ConfigGet("dropshadow_size"), pm)
'            Set pm = New CTempMsg
'            pm.Add "min", 0&
'            pm.Add "max", 60&
'            pm.Add "freq", 5&
'            pm.Add "label_size", 50&
'            .Add new_BPrefsControl("fancyslider", "default_duration", "", "Duration:", CStr(-1), pm)


'            i = Val(g_ConfigGet("default_duration"))
'            i = i \ 5
'            i = i * 5
'
'            If i = 0 Then
'                i = 1
'
'            ElseIf i = 5 Then
'                i = 2
'
'            ElseIf i = 10 Then
'                i = 3
'
'            ElseIf i = 30 Then
'                i = 4
'
'            ElseIf i = 60 Then
'                i = 5
'
'            Else
'                i = 3
'
'            End If
'        .Add new_BPrefsControl("fancycycle", "default_duration", "None (sticky)|5 seconds|10 seconds|30 seconds|1 minute", "Duration:", CStr(i))


'            Set pm = New CTempMsg
'            pm.Add "item-height", 36&
'            .Add new_BPrefsControl("combo", "default_style_list", , "Style:", , pm)
'            Set pm = New CTempMsg
'            pm.Add "item-height", 26&
'            .Add new_BPrefsControl("combo", "default_scheme_list", , "Scheme:", , pm)
'            Set pm = New CTempMsg
'            pm.Add "min", 1&
'            pm.Add "max", 100&
'            pm.Add "freq", 10&
'            pm.Add "suffix", "%"
'            .Add new_BPrefsControl("fancyslider", "global_opacity", "", "Opacity:", g_ConfigGet("global_opacity"), pm)

'            Set pm = New CTempMsg
'            pm.Add "min", -20&
'            pm.Add "max", 20&
'            pm.Add "freq", 5&
'            .Add new_BPrefsControl("fancyslider", "margin_spacing", "", "Spacing:", g_ConfigGet("margin_spacing"), pm)

