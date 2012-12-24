VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSmallPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Public Name As String
'
'Dim mPage As BPrefsPage
'Dim mInfoLabel As BControl
'
'Implements KPrefsPage
'
'Private Sub KPrefsPage_AllAttached()
'
'    Select Case Name
'
'    Case "gfx"
'        uUpdateStyleList        ' // only run *after* adding *both* combos!
'
'
'    End Select
'
'End Sub
'
'Private Sub KPrefsPage_Attached()
'Dim pm As CTempMsg
'Dim pc As BControl
'Dim i As Long
'
'    Set pm = New CTempMsg
'    pm.Add "height", 20
'
'    Set mInfoLabel = new_BPrefsControl("label", "")
'    mInfoLabel.SizeTo 0, 26
'
'    Select Case Name
'    Case "gen"
''        mPage.Add new_BPrefsControl("fancycycle", "ipb>timeout", "App Decides|Custom|Snarl Default", "Timeout:")
'        mPage.Add new_BPrefsControl("fancycycle", "ipb>timeout", "App Decides|Default|Custom", "Duration:")
'
'        ' /* R2.2: custom timeout slider */
'        Set pm = New CTempMsg
'        pm.Add "min", 0&
'        pm.Add "max", 60&
'        pm.Add "freq", 5&
'        pm.Add "label_size", 50&
'        mPage.Add new_BPrefsControl("fancyslider", "fs>gen.timeout", "", "", CStr(1), pm)
'
''        mPage.Add new_BPrefsControl("separator", "")
'
''        mPage.Add new_BPrefsControl("fancycycle", "cyc>timeout", "None (sticky)|5 seconds|10 seconds|30 seconds|1 minute", "Duration:", , , False)
'        mPage.Add new_BPrefsControl("fancycycle", "cyc>position", "Default#?-1|Top Left#?0|Top Right#?1|Bottom Left#?2|Bottom Right#?3", "Position:")
'        mPage.Add new_BPrefsControl("fancycycle", "cyc>suppression", "App Decides|Blocked|Not Blocked|Delayed", "Duplicates:")
'
'
'    Case "gfx"
''        mPage.Add new_bprefscontrol("fancybutton", "ipb>gfx", "Custom|Snarl Default", , , pm)
'        Set pm = New CTempMsg
'        pm.Add "text", "Use a specific style and scheme?"
'        pm.Add "align", 1
'        mPage.Add new_BPrefsControl("fancytoggle", "ft>gfx.custom", , , , pm)
'        mInfoLabel.SetText "If enabled, notifications of this class will use the following style and scheme:"
'        mPage.Add new_BPrefsCombo("cb>style", "", "Style:", , 20, , , , False)
'        mPage.Add new_BPrefsCombo("cb>scheme", "", "Scheme:", , 20, , , , False)
'
'
'    Case "aud"
''        mPage.Add new_bprefscontrol("fancybutton", "ipb>aud", "App Decides|Custom|Snarl Default", , , pm)
'        mPage.Add new_BPrefsControl("fancycycle", "ipb>aud.sound", "App Decides|Custom|None", "Sounds:")
'        mPage.Add new_BPrefsControl("fancyfilepicker", "fp>aud.sound", , "Custom:")
'        mPage.Add new_BPrefsControl("fancybutton2", "fb2>aud.test", "Test Sound")
'
'
'    Case "net"
''        mPage.Add new_BPrefsControl("fancytoggle2", "ft>forward", "Forward notifications?", , , , False)
'        mPage.Add new_BPrefsControl("fancycycle", "fc>fwd.when", "Never|Always|When Away|When Do Not Disturb Set", "When:")
'
''        mPage.Add new_BPrefsControl("label", "", "Style and Scheme to use:")
'        mPage.Add new_BPrefsCombo("cb>fwd_style", g_StyleRoster.GetNonWindowStyleList(), "Style:", , 20)
'        mPage.Add new_BPrefsCombo("cb>fwd_scheme", "", "Scheme:", , 20, , , , False)
'
'
''        mPage.Add new_BPrefsControl("edit", "ed>remote_addr", , "Host:")
''        mPage.Add new_BPrefsControl("fancycycle", "fc>host_type", "Snarl|Growl", "Type:")
'
''        mInfoLabel.SetText "If enabled, notifications of this class will be forwarded to the IP addresses set in the Network tab"
''        mPage.Add mInfoLabel
'
'        mPage.Add new_BPrefsControl("fancytoggle2", "ft>hide", "Suppress on-screen display?", , "0")
''        mPage.Add new_BPrefsControl("fancytoggle2", "ft>priority", "Only forward priority notifications?")
'
'
'    Case "adv"
'
'        mPage.Add new_BPrefsControl("fancytoggle2", "ft>adv.notify_only", "Show in view-only mode?")
'        mPage.Add new_BPrefsControl("label", "", "View-only mode makes the notification invisible to the mouse pointer, allowing you to access content beneath it.  View-only notifications cannot be interacted with.")
'
'        ' /* R2.2: priority notification */
'
'        mPage.Add new_BPrefsControl("fancytoggle2", "ft>adv.priority", "Treat as a priority notification?")
'
''        mPage.Add new_BPrefsControl("banner", "", "Duplicate Suppression")
''        mPage.Add new_BPrefsControl("separator", "")
'
''        mPage.Add new_BPrefsControl("fancycycle", "cyc>suppression", "App Decides|Blocked|Not Blocked|Delayed", "Duplicates:")
''        mPage.Add new_BPrefsControl("label", "", "If enabled, Snarl will attempt to block duplicate notifications.  If set to 'Delayed', Snarl will only block duplicates if they occur within a certain time period.")
'
'    End Select
'
'End Sub
'
'Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
'
'    If (gSelectedClass Is Nothing) Then _
'        Exit Sub
'
'Dim fUpdateConfig As Boolean
'Dim pc As BControl
'Dim pm As CTempMsg
'Dim sz As String
'Dim dw As Long
'
'    Select Case Control.GetName()
'
'    ' /* ----------------------------------- general page ----------------------------------- */
'
'    Case "ipb>timeout"
'        ' /* duration cycle button */
'        dw = Val(Value)
'
'        ' /* enable/disable the 'custom timeout' slider */
'
'        If mPage.Find("fs>gen.timeout", pc) Then _
'            pc.SetEnabled (dw = 3)
'
'        ' /* values translate to: App Decides (1) Default (2) Custom (3) */
'
''        If dw = 3 Then
''            ' /* custom duration */
''            gSelectedClass.AddFlag E_CLASS_CUSTOM_DURATION
''
''        Else
''            gSelectedClass.RemoveFlag E_CLASS_CUSTOM_DURATION
''            gSelectedClass.SetDuration IIf(dw = 1, E_APP_DECIDES_DURATION, E_DEFAULT_DURATION)
''
''        End If
'
'        fUpdateConfig = True
'
'
'    Case "fs>gen.timeout"
'        ' /* custom timeout slider */
'        Select Case Val(Value)
'        Case 0
'            Control.SetText "Sticky"
'
'        Case 1
'            Control.SetText "1 sec"
'
'        Case Else
'            Control.SetText Value & " secs"
'
'        End Select
'
'        gSelectedClass.SetDuration Val(Value)
'
'
'    Case "cyc>position"
'        ' /* position cycle button */
'        gSelectedClass.SetPosition (Val(Value) - 2)
'        fUpdateConfig = True
'
'
'
'    ' /* ----------------------------------- visuals page ----------------------------------- */
'
'
''    Case "ft>gfx.custom"
''
''        ' /* custom style toggle */
''
''        dw = Val(Value)
''        fUpdateConfig = True
''
''        If mPage.Find("cb>style", pc) Then _
''            pc.SetEnabled CBool(dw)
''
''        If mPage.Find("cb>scheme", pc) Then _
''            pc.SetEnabled CBool(dw)
''
''        If CBool(dw) Then
''            gSelectedClass.AddFlag E_CLASS_CUSTOM_STYLE
''
''        Else
''            gSelectedClass.RemoveFlag E_CLASS_CUSTOM_STYLE
''
''        End If
'
'
'    Case "cb>style"
'
'        ' /* 'style' combo */
'
'        uUpdateSchemeList
'
'
'    Case "cb>scheme"
'
'        ' /* 'scheme' combo */
'
'        Debug.Print ">>>> " & prefskit_GetItem(Control)
'
'        If mPage.Find("cb>style", pc) Then
'            ' /* get the selected style from the style combo */
'            sz = prefskit_GetItem(pc)
'            If sz <> "" Then
'                gSelectedClass.SetCustomStyle sz, prefskit_GetItem(Control)
'                fUpdateConfig = True
'
'            Else
'                Debug.Print "no style selected"
'
'            End If
'        End If
'
'
'    ' /* ----------------------------------- audio page ----------------------------------- */
'
'
'    Case "ipb>aud.sound"
'        ' /* cycle button - translates to: App Decides (1), Custom (2), None (3) */
'
'        Select Case Val(Value)
'        Case 1
'            ' /* app decides */
'            gSelectedClass.RemoveFlag (E_CLASS_CUSTOM_SOUND Or E_CLASS_NO_SOUND)
'            dw = 0
'
'        Case 2
'            ' /* custom sound */
'            gSelectedClass.RemoveFlag E_CLASS_NO_SOUND
'            gSelectedClass.AddFlag E_CLASS_CUSTOM_SOUND
'
'        Case 3
'            ' /* force no sound */
'            gSelectedClass.RemoveFlag E_CLASS_CUSTOM_SOUND
'            gSelectedClass.AddFlag E_CLASS_NO_SOUND
'
'        End Select
'
'        ' /* enable/disable the file picker */
'
'        If mPage.Panel.Find("fp>aud.sound", pc) Then _
'            pc.SetEnabled (Val(Value) = 2)
'
'        fUpdateConfig = True
'
'
'    Case "fp>aud.sound"
'        If Value <> gSelectedClass.CustomSound Then
'            gSelectedClass.SetCustomSound Value
'            sz = g_GetPath(Value)
'            If sz <> "" Then _
'                g_Prefs.last_sound_folder = sz
'
'            Debug.Print "** LAST SOUND FOLDER = " & sz & " ***"
'            fUpdateConfig = True
'
'        End If
'
'
'
'
'    ' /* ---------------------------------- advanced page ---------------------------------- */
'
'
'    Case "cyc>suppression"
'        ' /* duplicate suppression cycle button - translates to App Decides (1), Blocked (2), Not Blocked (3), Delayed (4) */
'
'        gSelectedClass.RemoveFlag (E_CLASS_ALLOW_DUPLICATES Or E_CLASS_DELAY_DUPLICATES Or E_CLASS_NO_DUPLICATES)
'
'        Select Case Val(Value)
'        Case 1
'            ' /* nothing to do! */
'
'        Case 2
'            gSelectedClass.AddFlag E_CLASS_NO_DUPLICATES
'
'        Case 3
'            gSelectedClass.AddFlag E_CLASS_ALLOW_DUPLICATES
'
'        Case 4
'            gSelectedClass.AddFlag E_CLASS_DELAY_DUPLICATES
'
'        End Select
'
'        fUpdateConfig = True
'
'
'    Case "ft>adv.priority"
'        ' /* 'priority notification' toggle */
'        If Val(Value) Then
'            gSelectedClass.AddFlag E_CLASS_PRIORITY_NOTIFICATION
'
'        Else
'            gSelectedClass.RemoveFlag E_CLASS_PRIORITY_NOTIFICATION
'
'        End If
'
'        fUpdateConfig = True
'
'
'    Case "ft>adv.notify_only"
'        ' /* 'notification only' toggle */
'        If Val(Value) Then
'            gSelectedClass.AddFlag E_CLASS_NOTIFICATION_ONLY
'
'        Else
'            gSelectedClass.RemoveFlag E_CLASS_NOTIFICATION_ONLY
'
'        End If
'
'        fUpdateConfig = True
'
'
'
'
'
'
'
'
'    Case "ipb>net"
'        dw = Val(Value)
'        If mPage.Find("lb>net", pc) Then _
'            pc.SetEnabled (dw = 2)
'
''        mInfoLabel.SetText IIf(dw = 1, "Notifications will not be forwarded", IIf(dw = 2, "Notifications will be forwarded to the following:", "Notifications will be forwarded to the list of IP addresses defined in the network tab"))
'
'    Case "ft>forward"
'        ' /* the forward/don't forward toggle button on the 'forwarding' page */
'        If mPage.Find("chk>hide", pc) Then _
'            pc.SetEnabled Val(Value)
'
'
'
'
'
'
'
'
'
'    ' /* advanced page */
'
'    Case "ft>adv.notify_only"
'        If Val(Value) = 0 Then
'            gSelectedClass.RemoveFlag E_CLASS_NOTIFICATION_ONLY
'
'        Else
'            gSelectedClass.AddFlag E_CLASS_NOTIFICATION_ONLY
'
'        End If
'
''        Debug.Print g_HexStr(gSelectedClass.PersistentFlags)
'
'    End Select
'
'
'    If fUpdateConfig Then _
'        g_WriteConfig
'
'
'End Sub
'
'Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
'
'    Select Case Control.GetName()
'    Case "fb2>aud.test"
'        If Not (gSelectedClass Is Nothing) Then _
'            g_PlaySound gSelectedClass.CustomSound
'
'    End Select
'
'End Sub
'
'Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
'End Sub
'
'Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
'Dim dw As Long
'
'    Set mPage = Page
'
'    Select Case Page.GetName()
'    Case "pg>netx" ', "pg>adv"
'        dw = 0
'
''    Case "pg>adv"
''        dw = 120
'
'    Case Else
'        dw = 90
'
'    End Select
'
'    mPage.SetMargin dw
'
'End Sub
'
'Private Sub KPrefsPage_Destroy()
'
'    Set mPage = Nothing
'
'End Sub
'
'Private Sub KPrefsPage_Detached()
'End Sub
'
'Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)
'End Sub
'
'Private Function KPrefsPage_hWnd() As Long
'End Function
'
'Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
'End Sub
'
'Private Sub uUpdateStyleList()
'Dim pc As BControl
'
'    If Not (mPage.Find("cb>style", pc)) Or (g_StyleRoster Is Nothing) Then _
'        Exit Sub
'
'    pc.SetText g_StyleRoster.GetStyleList()
'
'Dim pStyle As TStyle
'Dim pm As CTempMsg
'Dim sz As String
'Dim i As Long
'
'    Set pm = New CTempMsg
'
'    For i = 1 To g_StyleRoster.CountStyles()
'        Set pStyle = g_StyleRoster.StyleAt(i)
'        sz = pStyle.IconPath
'        If Not g_Exists(sz) Then _
'            sz = g_MakePath(App.Path) & "etc\icons\style.png"
'
'        pm.Replace "index", i
'        pm.Replace "image-file", sz
'        pc.DoExCmd B_SET_ITEM, pm
'
'    Next i
'
'End Sub
'
'Private Sub uUpdateSchemeList()
'Dim pc As BControl
'
'    If (Not (mPage.Find("cb>style", pc))) Or (g_StyleRoster Is Nothing) Then _
'        Exit Sub
'
'Dim px As TStyle
'
'    Set px = g_StyleRoster.StyleAt(Val(pc.GetValue()))
'    If (px Is Nothing) Then _
'        Exit Sub
'
'    g_Debug ">>> style = " & px.Name
'
'    If Not mPage.Find("cb>scheme", pc) Then _
'        Exit Sub
'
'    ' /* set the combo's content */
'
'    pc.SetText px.GetSchemeList()
'
'    ' /* set the icons */
'
'Dim pm As CTempMsg
'Dim i As Long
'
'    Set pm = New CTempMsg
'
'    For i = 1 To pc.DoExCmd(B_COUNT_ITEMS)
'        pm.Replace "index", i
'        pm.Replace "image-file", g_MakePath(px.Path) & px.SchemeAt(i) & ".icon"
'        pc.DoExCmd B_SET_ITEM, pm
'
'    Next i
'
'    ' /* select the first item */
'
'    pc.SetValue "1"
'
'''                    .EnableControl "cfg_style", (ps.Flags And SNARL_STYLE_HAS_SETTINGS)
'
'End Sub
'
'
'
'
'
'
