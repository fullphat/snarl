VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
Dim pc As BControl

    Select Case LCase$(mPage.GetName())
    Case "timer"
        If mPage.Find("timer_hours", pc) Then _
            pc.SetEnabled Not gConfig.timer_running

        If mPage.Find("timer_mins", pc) Then _
            pc.SetEnabled Not gConfig.timer_running

        uUpdateRemaining

    End Select

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg
Dim sz As String

    On Error Resume Next

    mPage.SetMargin 90

    Select Case LCase$(mPage.GetName())
    Case "general"

        With mPage
            
            .Add new_BPrefsControl("banner", "", "Startup")
            
'            Set pm = New CTempMsg
'            pm.Add "inset_by", CLng(Fix(.GetMargin() / 2))

            .Add new_BPrefsControl("fancytoggle2", "show_time_on_load", "Show current time?", , IIf(gConfig.show_time_on_load, "1", "0"))

            .Add new_BPrefsControl("banner", "", "Cookies")
            .Add new_BPrefsControl("fancytoggle2", "show_cookie", "Include random cookie in notification?", , IIf(gConfig.show_cookie, "1", "0"))

            sz = g_GetPath(gConfig.cookie_file)
            If sz <> "" Then
                Set pm = New CTempMsg
                pm.Add "path", sz

            End If

            .Add new_BPrefsControl("fancyfilepicker", "cookie_file", gConfig.cookie_file, "Cookie file:", , pm)

            .Add new_BPrefsControl("banner", "", "About " & App.Title)
            .Add new_BPrefsControl("label", "", "V" & CStr(App.Major) & "." & CStr(App.Revision), "Version:")
            .Add new_BPrefsControl("label", "", App.LegalCopyright)
            .Add new_BPrefsControl("fancybutton2", "go_website", "www.fullphat.net", "Website:")

        End With


    Case "timer"
        Set gTimerPrefsPage = mPage
        With mPage
'            .SetMargin 32

            .Add new_BPrefsControl("banner", "", "Duration")
            Set pm = New CTempMsg
            pm.Add "min", 0&
            pm.Add "max", 23&
            pm.Add "freq", 2&
            .Add new_BPrefsControl("fancyslider", "timer_hours", , "Hours:", CStr(gConfig.timer_hours), pm)

            Set pm = New CTempMsg
            pm.Add "min", 0&
            pm.Add "max", 59&
            pm.Add "freq", 5&
            .Add new_BPrefsControl("fancyslider", "timer_mins", , "Minutes:", CStr(gConfig.timer_mins), pm)

'            .Add new_BPrefsControl("banner", "", "Icon (Leave blank to use default)")

            .Add new_BPrefsControl("banner", "", "Options")
'            .Add new_BPrefsControl("fancytoggle2", "custom_timer_icon", "Use custom icon?", , IIf(gConfig.custom_timer_icon, "1", "0"))
            .Add new_BPrefsControl("fancyfilepicker", "timer_icon", gConfig.timer_icon, "Icon:", , pm)
            .Add new_BPrefsControl("label", "", "(Leave blank to use the default)")

            .Add new_BPrefsControl("fancyedit", "timer_message", gConfig.timer_message, "Message:", , pm)
            .Add new_BPrefsControl("label", "", "(Leave blank to use the default)")

            .Add new_BPrefsControl("fancytoggle2", "timer_10sec_warning", "Show notification during last 10 seconds?", , IIf(gConfig.timer_10sec_warning, "1", "0"))
'            .Add new_BPrefsControl("separator", "")

            .Add new_BPrefsControl("banner", "", "State")
            .Add new_BPrefsControl("fancytoggle", "start_stop_timer", "Paused|Running")
            .Add new_BPrefsControl("label", "timer_current", _
                                   Format$(gConfig.timer_hours, "00") & ":" & Format$(gConfig.timer_mins, "00") & ":00", _
                                   "Remaining:")



        End With
'            .AddControl "slider", "timer_hours", "", "Hours:", "0,12,1"
'            .SetValue "timer_hours", m_cdHour
'
'            .AddControl "slider", "timer_mins", "", "Minutes:", "0,59,5"
'            .SetValue "timer_mins", m_cdMin
'
'            If m_cdRunning Then
'                .EnableControl "timer_hours", False
'                .EnableControl "timer_mins", False
'
'            End If
'
'            .AddControl "file_picker", "cd_wav", m_cdWav, "Play:"
'            .AddControl "button", "cdbtn", IIf(m_cdRunning, "Stop", "Start"), , "128,28"
'            .AddControl "label", "lblcd", IIf(m_cdRunning, theHandler.Current(), "          ")
    
    
    Case "formats"

        With mPage

            .Add new_BPrefsControl("banner", "", "General Time Format")
            
            Set pm = New CTempMsg
            pm.Add "inset_by", CLng(Fix(.GetMargin() / 2))

            .Add new_BPrefsControl("fancytoggle2", "show_as_words", "Show as text?", , IIf(gConfig.show_as_words, "1", "0"), pm)
            
            .Add new_BPrefsControl("banner", "", "Numeric Time Format")
            .Add new_BPrefsControl("fancytoggle2", "use_12_hour_clock", "Use 12 hour clock?", , IIf(gConfig.use_12_hour_clock, "1", "0"), pm)
            .Add new_BPrefsControl("fancytoggle2", "two_digit_hour", "Two digit hour?", , IIf(gConfig.two_digit_hour, "1", "0"), pm)
            .Add new_BPrefsControl("fancytoggle2", "show_seconds", "Show seconds?", , IIf(gConfig.show_seconds, "1", "0"), pm)
            .Add new_BPrefsControl("fancytoggle2", "show_am_pm", "Show AM/PM?", , IIf(gConfig.show_am_pm, "1", "0"), pm)

'            .Add new_BPrefsControl("fancytoggle2", "custom_am_pm", "Custom AM/PM values?", , IIf(1, "1", "0"), pm, Not gConfig.show_as_words)

            .Add new_BPrefsControl("banner", "", "Custom AM/PM Mnemonics")
            Set pc = new_BPrefsControl("fancyedit", "am_val", gConfig.am_val, "AM:")
            pc.SizeTo 48, pc.Height
            .Add pc

            Set pc = new_BPrefsControl("fancyedit", "pm_val", gConfig.pm_val, "PM:")
            pc.SizeTo 48, pc.Height
            .Add pc

            .Add new_BPrefsControl("separator", "")
'            .Add new_BPrefsControl("label", "", "These settings affect the time format used in any notifications this extension displays")

            Set pc = new_BPrefsControl("fancybutton2", "test", "Test Settings")
            pc.SizeTo 128, 28
            .Add pc

        End With

    End Select


'    With mPage
'        .SetMargin 0
'
'        .Add new_BPrefsControl("label", "", "Notification timeout:")
'
'        Set pm = New CTempMsg
'        pm.Add "min", 0&
'        pm.Add "max", 60&
'        pm.Add "freq", 5&
'        pm.Add "label_size", 50&
'        .Add new_BPrefsControl("fancyslider", "fs>NotificationTimeout", "", "", CStr(mSettings.NotificationTimeout), pm)
'
'        .Add new_BPrefsControl("fancytoggle2", "ft>IncludeTitle", "Include title in notification?", , IIf(mSettings.IncludeTitle, "1", "0"))
'        .Add new_BPrefsControl("fancytoggle2", "ft>IncludeText", "Include text in notification?", , IIf(mSettings.IncludeText, "1", "0"))
'
''        .Add new_BPrefsControl("label", "", "Low power threshold:")
''
''        Set pm = New CTempMsg
''        pm.Add "min", 10&
''        pm.Add "max", 90&
''        pm.Add "freq", 10&
''        pm.Add "label_size", 50&
''        pm.Add "suffix", "%"
''        pm.Add "colour", rgba(255, 172, 172)
''        .Add new_BPrefsControl("fancyslider", "fs>LowPowerThreshold", "", "", CStr(mSettings.LowPowerThreshold), pm)
'
'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("fancytoggle2", "ft>MeterStyleMode", "Use Meter style friendly notifications?", , IIf(mSettings.MeterStyleMode, "1", "0"))
'        .Add new_BPrefsControl("label", "", "The Meter style expects a numeric value between 0 and 100 in the notification text.  This is then displayed graphically as a percentage by the selected scheme.")
'
'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("label", "", App.FileDescription & " V" & CStr(App.Major) & "." & CStr(App.Revision), , , , False)
'
'    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pc As BControl
Dim nm As Integer
Dim nh As Integer

    g_Debug "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()

    ' /* general */

    Case "show_time_on_load"
        gConfig.show_time_on_load = (Value = "1")
        g_WriteConfig

    Case "show_cookie"
        gConfig.show_cookie = (Value = "1")
        g_WriteConfig

    Case "cookie_file"
        gConfig.cookie_file = Value
        g_WriteConfig



    ' /* timer */

    Case "start_stop_timer"

        gConfig.timer_running = (Val(Value) <> 0)

        If gConfig.timer_running Then
            ' /* resumed/started timer so calculate minutes */
            If mPage.Find("timer_hours", pc) Then _
                nh = Val(pc.GetValue())

            If mPage.Find("timer_mins", pc) Then _
                nm = Val(pc.GetValue())

        End If

        If mPage.Find("timer_hours", pc) Then _
            pc.SetEnabled Not gConfig.timer_running

        If mPage.Find("timer_mins", pc) Then _
            pc.SetEnabled Not gConfig.timer_running

        ' /* notify our window of the change */

        SendMessage ghWnd, WM_START_COUNTDOWN, IIf(gConfig.timer_running, 1, 0), 0


    Case "timer_hours"
        uUpdateRemaining

    Case "timer_mins"
        uUpdateRemaining

    Case "timer_10sec_warning"
        gConfig.timer_10sec_warning = (Value = "1")
        g_WriteConfig
    
    Case "timer_icon"
        gConfig.timer_icon = Value
        g_WriteConfig

    Case "timer_message"
        gConfig.timer_message = Value
        g_WriteConfig

    ' /* formats */

    Case "show_as_words"
        gConfig.show_as_words = (Value = "1")
        g_WriteConfig
        uEnableTimeSettings Not gConfig.show_as_words

    Case "use_12_hour_clock"
        gConfig.use_12_hour_clock = (Value = "1")
        g_WriteConfig

    Case "show_am_pm"
        gConfig.show_am_pm = (Value = "1")
        g_WriteConfig

    Case "show_seconds"
        gConfig.show_seconds = (Value = "1")
        g_WriteConfig

    Case "two_digit_hour"
        gConfig.two_digit_hour = (Value = "1")
        g_WriteConfig

    Case "am_val"
        gConfig.am_val = Value

    Case "pm_val"
        gConfig.pm_val = Value

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)
Dim pc As BControl

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()

    Case "test"
        g_Notify True

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()

    If LCase$(mPage.GetName()) = "timer" Then _
        Set gTimerPrefsPage = Nothing

End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub uUpdateRemaining()
Dim pc As BControl
Dim nm As Integer
Dim nh As Integer

    ' /* get current hour value */
    If mPage.Find("timer_hours", pc) Then _
        nh = Val(pc.GetValue)

    ' /* get current minute value */
    If mPage.Find("timer_mins", pc) Then _
        nm = Val(pc.GetValue)

    ' /* set the label */
    If mPage.Find("timer_current", pc) Then _
        pc.SetText Format$(nh, "00") & ":" & Format$(nm, "00") & ":00"

    ' /* if the timer isn't running then disable the switch if 0:0 is specified */
    If Not gConfig.timer_running Then
        If mPage.Find("start_stop_timer", pc) Then _
            pc.SetEnabled Not ((nh = 0) And (nm = 0))

        ' /* tell our window handler the timeout has changed */
        SendMessage ghWnd, WM_SET_COUNTDOWN, 0, ByVal MAKELONG(nm, nh)

    End If

End Sub

Private Sub uEnableTimeSettings(ByVal Enabled As Boolean)
'Dim pc As BControl
'
'    If mPage.Find("use_12_hour_clock", pc) Then _
'        pc.SetEnabled Enabled
'
'    If mPage.Find("two_digit_hour", pc) Then _
'        pc.SetEnabled Enabled
'
'    If mPage.Find("show_seconds", pc) Then _
'        pc.SetEnabled Enabled
'
'    If mPage.Find("show_am_pm", pc) Then _
'        pc.SetEnabled Enabled
'
'    If mPage.Find("custom_am_pm", pc) Then _
'        pc.SetEnabled Enabled


End Sub



