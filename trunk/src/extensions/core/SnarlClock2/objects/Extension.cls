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
    '/  Created:        3-Apr-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* our constants */

Const EXT_DATE = "10-Feb-2012"
Const EXT_REL = "3.02"

Const UID_HOURLY = "_clock+_hourly"
Const UID_TIMER = "_clock+_timer"

Const CLASS_HOUR_CHIME = "hrch"
Const CLASS_COUNTDOWN_COMPLETE = "ccom"
Const CLASS_REMINDER = "rmdr"
Const CLASS_ALARM = "alrm"

Enum CLOCK_FORMAT_TYPES
    CFT_NUMERIC = 0
    CFT_TEXT = 1
'    CFT_SPECIAL = 2

End Enum

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

Dim mLastErr As String
Dim mConfigFile As CConfFile

Dim mPanel As BPrefsPanel           ' // our prefs panel
Dim mhWndOwner As Long              ' // handle of the window which launched our config (typically Snarl's Prefs panel)
Dim mPassword As String

    ' /* hourly chime */
Dim mLastHour As Integer

    ' /* cookies */
Dim mCookie() As String
Dim mCookies As Long

    ' /* timer */
Dim mTimer As Date
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1
Dim mTimerRunning As Boolean

Dim WithEvents theWatchTimer As BTimer
Attribute theWatchTimer.VB_VarHelpID = -1
Dim WithEvents theKeyTimer As BTimer
Attribute theKeyTimer.VB_VarHelpID = -1
Dim WithEvents theCornerTimer As BTimer
Attribute theCornerTimer.VB_VarHelpID = -1

    ' /* reminders */
Dim WithEvents theReminderPanel As TReminderPanel
Attribute theReminderPanel.VB_VarHelpID = -1
Dim mRemConfig As CConfFile3
Dim mRemList As BTagList
Dim WithEvents theAlarmAndReminderTimer As BTimer
Attribute theAlarmAndReminderTimer.VB_VarHelpID = -1

    ' /* alarms */
Dim mAlarmConfig As CConfFile3
Dim mAlarmList As BTagList
Dim WithEvents theAlarmPanel As TReminderPanel
Attribute theAlarmPanel.VB_VarHelpID = -1

Implements KPrefsPanel              ' // we're a panel handler
Implements KPrefsPage               ' // and a page handler
Implements ISnarlExtension

Private Function ISnarlExtension_GetConfigWindow() As Long
Dim pPage As BPrefsPage
Dim pm As CTempMsg
Dim sz As String
Dim pc As BControl


    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle App.Title & " Preferences"
            .SetWidth 390
            .SetWindow 1                    ' // so calling Go() doesn't show the window

            Set pPage = new_BPrefsPage("General", uPanelIcon("general"), Me)

            ' /* General */

            With pPage
                .SetMargin 90
                .SetWidth 390
                .Add new_BPrefsControl("banner", "", "Startup")
                .Add new_BPrefsControl("fancytoggle2", "show_time_on_load", "Show current time?", , mConfigFile.ValueOf("show_time_on_load"))
                .Add new_BPrefsControl("fancytoggle2", "show_cookie_on_load", "Display random cookie?", , mConfigFile.ValueOf("show_cookie_on_load"))

                .Add new_BPrefsControl("banner", "", "Cookies")
                .Add new_BPrefsControl("fancytoggle2", "show_cookie", "Include random cookie in notification?", , mConfigFile.ValueOf("show_cookie"))
                .Add new_BPrefsControl("fancyfilepicker", "cookie_file", mConfigFile.ValueOf("cookie_file"), "Cookies:")

                .Add new_BPrefsControl("banner", "", "Other")
                .Add new_BPrefsControl("fancytoggle2", "show_time_on_key", "Show time when the Tab key is held?", , mConfigFile.ValueOf("show_time_on_key"))
                .Add new_BPrefsControl("fancytoggle2", "show_time_on_hotspot", "Show time when the pointer moves to the:", , mConfigFile.ValueOf("show_time_on_hotspot"))
                .Add new_BPrefsControl("fancycycle", "hotspot_corner", "Top-left corner|Top-right corner|Bottom-left corner|Bottom-right corner", , mConfigFile.ValueOf("hotspot_corner"), , (mConfigFile.ValueOf("show_time_on_hotspot") = "1"))

            End With

            .AddPage pPage

            ' /* Timer */

            Set pPage = new_BPrefsPage("Timer", uPanelIcon("timer"), Me)
            With pPage
                .SetMargin 90
                .Add new_BPrefsControl("banner", "", "Duration")
                Set pm = New CTempMsg
                pm.Add "min", 0&
                pm.Add "max", 23&
                pm.Add "freq", 2&
                .Add new_BPrefsControl("fancyslider", "timer_hours", , "Hours:", mConfigFile.ValueOf("timer_hours"), pm)

                Set pm = New CTempMsg
                pm.Add "min", 0&
                pm.Add "max", 59&
                pm.Add "freq", 5&
                .Add new_BPrefsControl("fancyslider", "timer_mins", , "Minutes:", mConfigFile.ValueOf("timer_mins"), pm)

                .Add new_BPrefsControl("fancytoggle", "start_stop_timer", "Paused|Running")
                .Add new_BPrefsControl("label", "timer_current", _
                                       Format$(Val(mConfigFile.ValueOf("timer_hours")), "00") & ":" & Format$(Val(mConfigFile.ValueOf("timer_mins")), "00") & ":00", _
                                       "Remaining:")


                .Add new_BPrefsControl("banner", "", "Options")
'            .Add new_BPrefsControl("fancytoggle2", "custom_timer_icon", "Use custom icon?", , IIf(mconfigfile.custom_timer_icon, "1", "0"))
                .Add new_BPrefsControl("fancyfilepicker", "timer_icon", mConfigFile.ValueOf("timer_icon"), "Icon:", , pm)

                .Add new_BPrefsControl("fancyedit", "timer_message", mConfigFile.ValueOf("timer_message"), "Message:", , pm)
                .Add new_BPrefsControl("label", "", "(Leave either field blank to use the default)")

                .Add new_BPrefsControl("fancytoggle2", "timer_10sec_warning", "Show notification during last 10 seconds?", "", mConfigFile.ValueOf("timer_10sec_warning"))
'                .Add new_BPrefsControl("fancytoggle2", "show_meter", "Meter-friendly?", , mConfigFile.ValueOf("show_meter"))
'                .Add new_BPrefsControl("label", "", "Enabling this option causes " & App.Title & " to include text in the notification which some styles can display graphically.")

            End With

            .AddPage pPage



            ' /* alarms */

            Set pPage = new_BPrefsPage("Alarms", uPanelIcon("alarms"), Me)
            With pPage
                .SetMargin 0
'                .Add new_BPrefsControl("banner", "", "Alarms")

                Set pm = New CTempMsg
                pm.Add "checkboxes", 1&
                pm.Add "plusminus", 1&
                pm.Add "item-height", 36&

                Set pc = new_BPrefsControl("listbox", "alarms_list", , , , pm)
                pc.SizeTo 0, (6 * 36) + 2
                .Add pc
                .Add new_BPrefsControl("fancyplusminus", "add_alarm")
    
            End With

            .AddPage pPage




            ' /* reminders */

            Set pPage = new_BPrefsPage("Reminders", uPanelIcon("reminders"), Me)
            With pPage
                .SetMargin 10
'                .Add new_BPrefsControl("banner", "", "Reminders")

                Set pm = New CTempMsg
'                pm.Add "checkboxes", 1&
                pm.Add "plusminus", 1&
                pm.Add "item-height", 36&

                Set pc = new_BPrefsControl("listbox", "reminders_list", , , , pm)
                pc.SizeTo 0, (6 * 36) + 4
                .Add pc
                .Add new_BPrefsControl("fancyplusminus", "add_reminder")

'                .Add new_BPrefsControl("banner", "", "Options")
                .Add new_BPrefsControl("label", "", "Show notification:")
                .Add new_BPrefsControl("fancycycle", "reminder_offset", "On the event|5 minutes before|10 minutes before|30 minutes before|1 hour before|1 day before", , CStr(uGetReminderOffset() + 1))

            End With

            .AddPage pPage



            ' /* formats */

            Set pPage = new_BPrefsPage("Formats", uPanelIcon("format"), Me)
            With pPage
                .SetMargin 24
                .Add new_BPrefsControl("banner", "", "General Time Format")

                Set pm = New CTempMsg
                pm.Add "inset_by", CLng(Fix(.GetMargin() / 2))

'                .Add new_BPrefsControl("fancytoggle2", "show_as_words", "Show as text?", , IIf(mconfigfile.show_as_words, "1", "0"), pm)
                .Add new_BPrefsControl("fancycycle", "format_type", "Numeric|Text", , CStr(Val(mConfigFile.ValueOf("format_type")) + 1))

                .Add new_BPrefsControl("banner", "", "Numeric Time Format")
                .Add new_BPrefsControl("fancytoggle2", "use_12_hour_clock", "Use 12 hour clock?", , mConfigFile.ValueOf("use_12_hour_clock"))
                .Add new_BPrefsControl("fancytoggle2", "two_digit_hour", "Two digit hour?", , mConfigFile.ValueOf("two_digit_hour"))
                .Add new_BPrefsControl("fancytoggle2", "show_seconds", "Show seconds?", , mConfigFile.ValueOf("show_seconds"))
                .Add new_BPrefsControl("fancytoggle2", "show_am_pm", "Show AM/PM?", , mConfigFile.ValueOf("show_am_pm"), , (mConfigFile.ValueOf("use_12_hour_clock") = "1"))
    
    '            .Add new_BPrefsControl("fancytoggle2", "custom_am_pm", "Custom AM/PM values?", , IIf(1, "1", "0"), pm, Not mconfigfile.show_as_words)

'                .Add new_BPrefsControl("banner", "", "Custom Abbreviations")
'                Set pc = new_BPrefsControl("fancyedit", "am_val", mConfigFile.ValueOf("am_val"), "AM:")
'                pc.SizeTo 48, pc.Height
'                .Add pc
'
'                Set pc = new_BPrefsControl("fancyedit", "pm_val", mConfigFile.ValueOf("pm_val"), "PM:")
'                pc.SizeTo 48, pc.Height
'                .Add pc

                .Add new_BPrefsControl("separator", "")
    '            .Add new_BPrefsControl("label", "", "These settings affect the time format used in any notifications this extension displays")

                Set pc = new_BPrefsControl("fancybutton2", "test", "Test Settings")
                pc.SizeTo 128, 28
                .Add pc
    
            End With

            .AddPage pPage

            
            ' /* about */

            Set pPage = new_BPrefsPage("About", uPanelIcon("about"), Me)
            With pPage
                .SetMargin 90
                .Add new_BPrefsControl("banner", "", "About " & App.Title)
                .Add new_BPrefsControl("label", "", "V" & CStr(App.Major) & "." & CStr(App.Revision), "Version:")
                .Add new_BPrefsControl("fancybutton2", "go_website", "www.fullphat.net", "Website:")

            
            End With

            .AddPage pPage
            
            
            ' /* create it */

            .Go
            g_SetWindowIconToAppResourceIcon .hWnd

        End With

    End If

    ISnarlExtension_GetConfigWindow = mPanel.hWnd

End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "full phat products"
        .Copyright = App.LegalCopyright
        .Date = EXT_DATE
        .Description = App.FileDescription
        .Flags = SNARL_EXTN_IS_CONFIGURABLE Or SNARL_EXTN_WANTS_PULSE
        .IconPath = g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png"
        .Name = App.Title
        .Path = App.Path
        .Release = EXT_REL
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With

End Sub

Private Function ISnarlExtension_Initialize() As Long
Dim szPath As String
Dim i As Integer

    On Error Resume Next

    l3OpenLog "%APPDATA%\" & App.Title & ".log"
    g_Debug App.Title & " running from '" & g_MakePath(App.Path) & "'"
    g_Debug ""

    mLastErr = ""
    mPassword = ""

    ' /* does snarl's /etc/ path exist? */

    If Not snGetConfigPath(szPath) Then
        g_Debug "Extension.Initialize(): snGetConfigPath() failed", LEMON_LEVEL_CRITICAL
        mLastErr = "Couldn't get Snarl user path"
        ISnarlExtension_Initialize = 201
        Exit Function

    End If

    g_Debug "Extension.Initialize(): snGetConfigPath() is '" & szPath & "'"

'                Else
'                    ' /* set from previous version */
'                    If .FindEntry("show_as_words", szData) Then _
'                        gConfig.format_type = Val(szData)
'
'                End If
'
'
'                If .FindEntry("debug_mode", szData) Then _
'                    gConfig.debug_enabled = (szData = "1")

    ' /* create dynamic password */

    For i = 1 To 32
        Randomize Timer
        mPassword = mPassword & Chr$((Rnd * (255 - 48)) + 48)

    Next i

    g_Debug "Extension.Initialize(): done"

    ' /* ok */

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()

    ' /* hourly chime */

    If (Minute(Now()) = 0) And (Second(Now()) = 0) And (Hour(Now()) <> mLastHour) Then
        snDoRequest "notify?app-sig=" & App.ProductName & "&replace-uid=" & UID_HOURLY & "&uid=" & UID_HOURLY & _
                    "&class=" & CLASS_HOUR_CHIME & "&title=Hourly Reminder&text=" & uGetTimeString() & _
                    "&icon=" & g_MakePath(App.Path) & "hourly.png" & "&password=" & mPassword & "&log=0" & _
                    "&value-date-packed=" & uGetPackedDate(Now())

        mLastHour = Hour(Now)

    End If

End Sub

Private Sub ISnarlExtension_Start()
Dim hr As Long

    g_Debug "ISnarlExtension_Start", LEMON_LEVEL_PROC

    ' /* register */
    hr = snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png", mPassword)
    If hr < 0 Then
        g_Debug "ISnarlExtension_Start(): registration failed: " & CStr(hr), LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    ' /* add classes */

    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & CLASS_HOUR_CHIME & "&name=" & "Hourly chime / Ad hoc display"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & CLASS_COUNTDOWN_COMPLETE & "&name=" & "Countdown complete"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & CLASS_REMINDER & "&name=" & "Reminder"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & CLASS_ALARM & "&name=" & "Alarm"

Dim szPath As String
Dim ps As CConfSection
Dim pr As TReminder

    snGetConfigPath szPath

    ' /* (re)read reminders */

    Set mRemList = new_BTagList()
    Set mRemConfig = New CConfFile3
    With mRemConfig
        .SetFile g_MakePath(szPath) & "clock+reminders.conf"
        .Load

        .Rewind
        Do While .GetNextSection(ps)
            Set pr = New TReminder
            If pr.SetFrom(ps) Then _
                mRemList.Add pr

        Loop

    End With

Dim pa As TAlarm

    ' /* (re)read alarms */

    Set mAlarmList = new_BTagList()
    Set mAlarmConfig = New CConfFile3
    With mAlarmConfig
        .SetFile g_MakePath(szPath) & "clock+alarms.conf"
        .Load
        .Rewind
        Do While .GetNextSection(ps)
            Set pa = New TAlarm
            If pa.SetFrom(ps) Then _
                mAlarmList.Add pa

        Loop

    End With



    ' /* (re)read config */

    If g_Exists(g_MakePath(szPath) & ".snarlclock2") Then
        MoveFile g_MakePath(szPath) & ".snarlclock2", g_MakePath(szPath) & "clock+.conf"

'    ElseIf g_Exists(g_MakePath(szPath) & "extensions\snarlclock\snarlclock.conf") Then
'        MoveFile g_MakePath(szPath) & "extensions\snarlclock\snarlclock.conf", g_MakePath(szPath) & "clock+.conf"
    
    End If

    Set mConfigFile = New CConfFile
    With mConfigFile
        .SetTo g_MakePath(szPath) & "clock+.conf"

        .AddIfMissing "cookie_file", g_MakePath(App.Path) & "cookies.txt"
        .AddIfMissing "format_type", "1"        ' // text
        .AddIfMissing "show_cookie", "1"
        .AddIfMissing "show_time_on_load", "1"
        .AddIfMissing "use_12_hour_clock", "1"
        .AddIfMissing "show_am_pm", "1"
        .AddIfMissing "show_seconds", "0"
        .AddIfMissing "two_digit_hour", "0"
        .AddIfMissing "am_val", "am"
        .AddIfMissing "pm_val", "pm"
        .AddIfMissing "timer_icon", ""
        .AddIfMissing "timer_10sec_warning", "1"
        .AddIfMissing "timer_message", ""
        .AddIfMissing "timer_hours", "0"
        .AddIfMissing "timer_mins", "10"

        .AddIfMissing "show_meter", "1"
        .AddIfMissing "show_cookie_on_load", "1"

        .AddIfMissing "show_time_on_key", "0"
        .AddIfMissing "show_time_on_hotspot", "0"
        .AddIfMissing "hotspot_corner", "1"
        .AddIfMissing "reminder_offset", "0"


        ' /* hotspot or key watch? */

        uSetHotSpotWatch

        ' /* show time on load? */

        If .ValueOf("show_time_on_load") = "1" Then _
            uShowCurrentTime

        ' /* get the cookies */

        uLoadCookies .ValueOf("cookie_file")

        ' /* seperate cookie notification */

Dim sz As String

        If (.ValueOf("show_cookie_on_load") = "1") And (uGetRandomCookie(sz)) Then _
            snarl_ez_notify App.ProductName, "", _
                            "", sz, _
                            g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "cookie.png", _
                            , , _
                            mPassword

    End With

    Set theAlarmAndReminderTimer = new_BTimer(5000)

End Sub

Private Sub ISnarlExtension_Stop()

    Set theWatchTimer = Nothing
    Set theKeyTimer = Nothing
    Set theCornerTimer = Nothing
    Set theAlarmAndReminderTimer = Nothing
    snarl_unregister App.ProductName, mPassword

End Sub

Private Sub ISnarlExtension_TidyUp()

    If Not (mPanel Is Nothing) Then _
        mPanel.Quit

End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    EnableWindow mhWndOwner, -1
    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
Dim pc As BControl

    If mPanel.Find("timer_hours", pc) Then _
        pc.SetEnabled Not mTimerRunning

    If mPanel.Find("timer_mins", pc) Then _
        pc.SetEnabled Not mTimerRunning

    uUpdateRemaining
    uUpdateReminderList
    uUpdateAlarmList

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pc As BControl
Dim nm As Integer
Dim nh As Integer
Dim i As Long

'    g_Debug "[" & mpanel.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()

    Case "format_type"
        ' /* format type needs a little tweaking before setting */
        mConfigFile.Update Control.GetName(), CStr(Val(Value) - 1)
        mConfigFile.Save
        uEnableTimeSettings (Val(Value) = 0)


    Case "start_stop_timer"
        ' /* timer master control */

        mTimerRunning = (Val(Value) <> 0)

        ' /* enable/disable the sliders */

        prefskit_SafeEnable mPanel, "timer_hours", Not mTimerRunning
        prefskit_SafeEnable mPanel, "timer_mins", Not mTimerRunning

        If mTimerRunning Then
            ' /* resumed/started timer so calculate minutes */

            If mPanel.Find("timer_hours", pc) Then _
                nh = Val(pc.GetValue())

            If mPanel.Find("timer_mins", pc) Then _
                nm = Val(pc.GetValue())

            mTimer = TimeSerial(nh, nm, Second(mTimer))

            snDoRequest "hide?app-sig=" & App.ProductName & "&uid=" & UID_TIMER & "&password=" & mPassword
            Set theTimer = new_BTimer(1000)

        Else
            ' /* stop the timer */
            Set theTimer = Nothing

        End If


    Case "timer_hours"
        ' /* timer hour slider */
        uUpdateRemaining


    Case "timer_mins"
        ' /* timer minute slider */
        uUpdateRemaining


    Case "add_reminder"
        ' /* add/remove reminder */
        If Value = "+" Then
            Set theReminderPanel = New TReminderPanel
            theReminderPanel.Go mPanel.hWnd, New TReminder, True

        Else
            i = Val(prefskit_GetValue(mPanel, "reminders_list"))
            If i > 0 Then
                mRemConfig.RemoveSection mRemConfig.IndexOf(mRemList.TagAt(i).Name)
                mRemConfig.Save
                mRemList.Remove i
                uUpdateReminderList

            End If
        End If


    Case "add_alarm"
        ' /* add/remove reminder */
        If Value = "+" Then
            Set theAlarmPanel = New TReminderPanel
            theAlarmPanel.IsAlarmPanel = True
            theAlarmPanel.Go mPanel.hWnd, New TAlarm, True

        Else
            i = Val(prefskit_GetValue(mPanel, "alarms_list"))
            If i > 0 Then
                mAlarmConfig.RemoveSection mAlarmConfig.IndexOf(mAlarmList.TagAt(i).Name)
                mAlarmConfig.Save
                mAlarmList.Remove i
                uUpdateAlarmList

            End If
        End If


    Case "reminder_offset"
        ' /* amount of time applied to reminder notifications */
        mConfigFile.Update Control.GetName(), CStr(Val(Value) - 1)
        mConfigFile.Save


    Case Else
        mConfigFile.Update Control.GetName(), Value
        mConfigFile.Save

    End Select



    ' /* post-processing */

    Select Case Control.GetName()
    Case "show_time_on_hotspot"
        prefskit_SafeEnable mPanel, "hotspot_corner", (Value = "1")
        uSetHotSpotWatch

    Case "show_time_on_key"
        uSetHotSpotWatch

    Case "use_12_hour_clock"
        prefskit_SafeEnable mPanel, "show_am_pm", (Value = "1")
    
    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

'    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()

    Case "test"
        uShowCurrentTime

'        snarl_ez_notify App.ProductName, CLASS_HOUR_CHIME, _
                        "Hourly Reminder", uGetTimeString(), _
                        g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "hourly.png", , , mPassword


    Case "reminders_list"
        Set theReminderPanel = New TReminderPanel
        theReminderPanel.Go mPanel.hWnd, mRemList.TagAt(Val(prefskit_GetValue(mPanel, Control.GetName())))

    Case "alarms_list"
        Set theAlarmPanel = New TReminderPanel
        theAlarmPanel.IsAlarmPanel = True
        theAlarmPanel.Go mPanel.hWnd, mAlarmList.TagAt(Val(prefskit_GetValue(mPanel, Control.GetName())))

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
Dim pa As TAlarm
Dim sz As String
Dim i As Long

    Select Case Control.GetName()

    Case "alarms_list"
        If Notification = "checked" Then
            i = Val(prefskit_GetValue(mPanel, Control.GetName))
            Set pa = mAlarmList.TagAt(i)
            sz = prefskit_GetItemEx(mPanel, Control.GetName, "checked")
            pa.SetEnabled (sz = "1")
            mAlarmConfig.Save

        End If

    End Select

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

Private Sub uUpdateRemaining()
Dim pc As BControl
Dim nm As Integer
Dim nh As Integer

    ' /* get current hour value */
    If mPanel.Find("timer_hours", pc) Then _
        nh = Val(pc.GetValue)

    ' /* get current minute value */
    If mPanel.Find("timer_mins", pc) Then _
        nm = Val(pc.GetValue)

    ' /* set the label */
    If mPanel.Find("timer_current", pc) Then _
        pc.SetText Format$(nh, "00") & ":" & Format$(nm, "00") & ":00"

    ' /* if the timer isn't running then disable the switch if 0:0 is specified */
    If Not mTimerRunning Then
        If mPanel.Find("start_stop_timer", pc) Then _
            pc.SetEnabled Not ((nh = 0) And (nm = 0))

        mTimer = TimeSerial(nh, nm, 0)

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

Private Sub theAlarmPanel_Closed()

    Set theAlarmPanel = Nothing

End Sub

Private Sub theAlarmPanel_Save(Item As melon4.BTagItem)
Dim pa As TAlarm

    If mAlarmList.IndexOf(Item.Name) = 0 Then
        ' /* is a new alarm */
        mAlarmList.Add Item

        ' /* add to persistent config */
        Set pa = Item
        mAlarmConfig.Add pa.Config

    Else
        ' /* existing alarm */

    End If

    ' /* save config */

    mAlarmConfig.Save

    ' /* update list */
    
    uUpdateAlarmList

End Sub

Private Sub theCornerTimer_Pulse()

    uShowCurrentTime

End Sub

Private Sub theKeyTimer_Pulse()

    uShowCurrentTime

End Sub

Private Sub theReminderPanel_Closed()

    Set theReminderPanel = Nothing

End Sub

Private Sub theReminderPanel_Save(Item As melon4.BTagItem)
Dim pr As TReminder

    If mRemList.IndexOf(Item.Name) = 0 Then
        ' /* is a new reminder */
        mRemList.Add Item

        ' /* add to persistent config */
        Set pr = Item
        mRemConfig.Add pr.Config

    Else
        ' /* existing reminder */

    End If

    ' /* save config */

    mRemConfig.Save

    ' /* update list */
    
    uUpdateReminderList

End Sub

Private Sub theAlarmAndReminderTimer_Pulse()

    ' /* check reminders */

Dim pr As TReminder
Dim i As Long

    i = uGetReminderOffset()

    With mRemList
        .Rewind
        Do While .GetNextTag(pr) = B_OK
            If pr.IsDue(i) Then
                ' /* it's due now so create the notification */
                If snarl_ez_notify(App.ProductName, CLASS_REMINDER, pr.Description, pr.DueDelta, pr.Icon(True), pr.Priority, , mPassword, pr.Config.Name) > 0 Then
                    pr.SetWasDisplayed
                    uUpdateReminderList

                End If
            End If
        Loop

    End With

    ' /* check alarms */

Dim pa As TAlarm

    With mAlarmList
        .Rewind
        Do While .GetNextTag(pa) = B_OK
            If pa.IsDue() Then
                ' /* it's due now so create the notification */
                If snarl_ez_notify(App.ProductName, CLASS_ALARM, IIf(pa.Label = "", "Alarm", pa.Label), "", g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "alarm.png", pa.Priority, , _
                                   mPassword, pr.Config.Name) > 0 Then _
                    pa.SetAsDone

            End If
        Loop

    End With

End Sub

Private Sub theTimer_Pulse()
Dim szIcon As String
Dim szMsg As String
Static i As Long

    ' /* doing a countdown */

    mTimer = DateAdd("s", -1, mTimer)

    If (Hour(mTimer) = 0) And (Minute(mTimer) = 0) Then
        
        szIcon = mConfigFile.ValueOf("timer_icon")
        If szIcon = "" Then _
            szIcon = g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "timer.png"
        
        Select Case Second(mTimer)
        Case 0
            ' /* we're done */

            snDoRequest "hide?app-sig=" & App.ProductName & "&uid=" & UID_TIMER & "&password=" & mPassword
            szMsg = mConfigFile.ValueOf("timer_message")

            snarl_ez_notify App.ProductName, CLASS_COUNTDOWN_COMPLETE, _
                            IIf(szMsg = "", "Countdown done!", szMsg), _
                            "", _
                            szIcon, _
                            , , _
                            mPassword, _
                            UID_TIMER

            Set theTimer = Nothing

        Case 10
            ' /* show the T-10 countdown notification? */

            If mConfigFile.ValueOf("timer_10sec_warning") = "1" Then _
                snarl_ez_notify App.ProductName, CLASS_COUNTDOWN_COMPLETE, _
                                Second(mTimer) & " second" & IIf(Second(mTimer) = 1, "", "s") & " to go", _
                                IIf(mConfigFile.ValueOf("show_meter") = "1", "100", ""), _
                                szIcon, _
                                , , _
                                mPassword, _
                                UID_TIMER


        Case Is < 10
            If mConfigFile.ValueOf("timer_10sec_warning") = "1" Then _
                snDoRequest "notify?app-sig=" & App.ProductName & "&password=" & mPassword & "&uid=" & UID_TIMER & _
                            "&update-uid=" & UID_TIMER & _
                            "&title=" & CStr(Second(mTimer)) & " second" & IIf(Second(mTimer) = 1, "", "s") & " to go" & _
                            IIf(mConfigFile.ValueOf("show_meter") = "1", "&text=" & CStr(Second(mTimer) * 10), "")

        End Select

    End If

Static pc As BControl

    If Not (mPanel Is Nothing) Then

        ' /* update the panel controls */

'    g_Debug "uUpdateTimerControls(): " & CStr(mTimer)
        If (Hour(mTimer) = 0) And (Minute(mTimer) = 0) And (Second(mTimer) = 0) Then
            ' /* countdown completed: update toggle */
            If mPanel.Find("start_stop_timer", pc) Then
                pc.SetValue "0"
                pc.SetEnabled False

            End If
        End If

        ' /* update minute slider */

        If mPanel.Find("timer_mins", pc) Then _
            pc.SetValue CStr(Minute(mTimer))

        ' /* update hours slider */

        If mPanel.Find("timer_hours", pc) Then _
            pc.SetValue CStr(Hour(mTimer))

        ' /* label */

        If mPanel.Find("timer_current", pc) Then _
            pc.SetText Format$(mTimer, "hh:mm:ss")

    End If

End Sub

Private Function uPanelIcon(ByVal Name As String) As MImage

    Set uPanelIcon = load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & Name & ".png")

End Function

Private Function uGetTimeString() As String
Dim txt As String
Dim pt As CTime
Dim i As Integer

    ' /* construct the time format */

    g_Debug "uShowMsg()", LEMON_LEVEL_PROC

    Set pt = New CTime
    pt.SetTo Now()

Dim lType As CLOCK_FORMAT_TYPES
Dim bAmPm As Boolean
Dim b12hr As Boolean
Dim bSecs As Boolean
Dim b2dh As Boolean

    lType = g_SafeLong(mConfigFile.ValueOf("format_type"))
    bSecs = (mConfigFile.ValueOf("show_seconds") = "1")
    b12hr = (mConfigFile.ValueOf("use_12_hour_clock") = "1")
    b2dh = (mConfigFile.ValueOf("two_digit_hour") = "1")
    bAmPm = (mConfigFile.ValueOf("show_am_pm") = "1")

    txt = "It's now "

    If lType = CFT_TEXT Then
        ' /* as text */
        txt = txt & LCase$(pt.AsText(Not b12hr, bSecs))

    Else
        ' /* numerically */
        If b12hr Then
            i = pt.Hour12

        Else
            i = pt.Hour

        End If

        ' /* construct hh:mm */

        txt = txt & IIf(b2dh, Format$(i, "00"), CStr(i)) & ":" & Format$(pt.Minutes(), "00")

        ' /* add seconds */

        If bSecs Then _
            txt = txt & ":" & Format$(pt.Seconds(), "00")

        ' /* am/pm setting only applies for 12 hour clock */

        If (b12hr) And (bAmPm) Then
            If pt.IsAfternoon() Then
                txt = txt & mConfigFile.ValueOf("pm_val")

            Else
                txt = txt & mConfigFile.ValueOf("am_val")

            End If
        End If
    End If

Dim a As Long

    ' /* R2.06: suppress in-line cookie if using the packed time format */

    If (mConfigFile.ValueOf("show_cookie") = "1") And (mCookies > 0) Then
        Randomize Timer
        a = Rnd * mCookies
        txt = txt & vbCrLf & vbCrLf & mCookie(a)

    ElseIf (mConfigFile.ValueOf("show_cookie") = "1") And (mCookies = 0) Then
        g_Debug "uGetTimeString(): show_cookie option set but no cookies available", LEMON_LEVEL_WARNING

    End If

    uGetTimeString = txt

End Function

Private Sub uLoadCookies(ByVal Path As String)

    On Error Resume Next

    If Not g_Exists(Path) Then
        g_Debug "uLoadCookies(): '" & Path & "' does not exist", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim sz As String
Dim n As Integer

    mCookies = 0
    ReDim mCookie(0)

    n = FreeFile()
    Err.Clear
    Open Path For Input As #n
    If Err.Number = 0 Then
        Do While Not EOF(n)
            Line Input #n, sz
            If sz <> "" Then
                If Left$(sz, 1) <> "#" Then
                    mCookies = mCookies + 1
                    ReDim Preserve mCookie(mCookies)
                    mCookie(mCookies) = sz

                End If
            End If
        Loop
        Close #n
        g_Debug "uLoadCookies(): " & CStr(mCookies) & " loaded"

    Else
        g_Debug "uLoadCookies(): bad path '" & Path & "'", LEMON_LEVEL_CRITICAL

    End If

End Sub

Private Function uGetRandomCookie(ByRef Cookie As String) As Boolean

    If mCookies = 0 Then _
        Exit Function

    Randomize Timer
    Cookie = mCookie(Rnd * mCookies)
    uGetRandomCookie = (Cookie <> "")

End Function

Private Sub uSetHotSpotWatch()

    Set theWatchTimer = Nothing

    If (mConfigFile.ValueOf("show_time_on_key") = "1") Or (mConfigFile.ValueOf("show_time_on_hotspot") = "1") Then _
        Set theWatchTimer = new_BTimer(100)

End Sub

Private Sub theWatchTimer_Pulse()
Dim pt As POINTAPI
Dim rc As RECT

    If mConfigFile.ValueOf("show_time_on_hotspot") = "1" Then
        rc.Right = g_ScreenWidth()
        rc.Bottom = g_ScreenHeight()
        GetCursorPos pt

        Select Case Val(mConfigFile.ValueOf("hotspot_corner"))
        Case 1      ' // top-left
            rc.Right = 15
            rc.Bottom = 15

        Case 2      ' // top-right
            rc.Left = rc.Right - 15
            rc.Bottom = 15

        Case 3      ' // bottom-left
            rc.Right = 15
            rc.Top = rc.Bottom - 15

        Case 4      ' // bottom-right
            rc.Left = rc.Right - 15
            rc.Top = rc.Bottom - 15

        End Select

        If g_PointInRect(rc, pt) Then
            ' /* pointer is in the hotspot corner */
            If (theCornerTimer Is Nothing) Then
                ' /* lag timer isn't running so start now */
                Debug.Print "## pointer is in hotspot corner"
                Set theCornerTimer = new_BTimer(250, True)

            End If

        ElseIf Not (theCornerTimer Is Nothing) Then
            ' /* pointer has exited hotspot corner */
            Debug.Print "## pointer left hotspot corner"
            Set theCornerTimer = Nothing
            snDoRequest "hide?app-sig=" & App.ProductName & "&password=" & mPassword & "&uid=" & UID_HOURLY

        End If
    End If

    ' /* key press/hold */

    If mConfigFile.ValueOf("show_time_on_key") = "1" Then
        If (g_IsPressed(vbKeyTab)) And (Not g_IsPressed(VK_MENU)) Then
            If (theKeyTimer Is Nothing) Then
                Debug.Print "key pressed"
                Set theKeyTimer = new_BTimer(250, True)

            End If

        ElseIf Not (theKeyTimer Is Nothing) Then
            Debug.Print "key was released"
            Set theKeyTimer = Nothing
            snDoRequest "hide?app-sig=" & App.ProductName & "&password=" & mPassword & "&uid=" & UID_HOURLY

        End If
    End If

End Sub

Private Sub uShowCurrentTime()

'    snarl_ez_notify App.ProductName, CLASS_HOUR_CHIME, _
'                    App.Title, uGetTimeString(), g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png", , , _
'                    mPassword, UID_HOURLY

    snarl_notify App.ProductName, CLASS_HOUR_CHIME, UID_HOURLY, mPassword, App.Title, uGetTimeString(), _
                 g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png", , , , , _
                 "value-date-packed=" & uGetPackedDate(Now()) & "&log=0"

End Sub

Private Sub uUpdateReminderList()

    If (mPanel Is Nothing) Then _
        Exit Sub

Dim pr As TReminder
Dim sz As String

    With mRemList
        .Rewind
        Do While .GetNextTag(pr) = B_OK
            sz = sz & pr.Description & IIf(g_IsIDE(), " (due " & Format$(pr.NextDue, "d mmm yyyy hh:mm") & ")", "") & "#?0#?" & pr.Details & "|"
        
        Loop

    End With

    prefskit_SafeSetText mPanel, "reminders_list", g_SafeLeftStr(sz, Len(sz) - 1)

    ' /* icons */

Dim i As Long

    With mRemList
        If .CountItems Then
            For i = 1 To .CountItems
                Set pr = .TagAt(i)
                prefskit_SetItemEx mPanel, "reminders_list", "image-file", i, pr.Icon(True)

            Next i
        End If
    End With

End Sub

Private Sub uUpdateAlarmList()
Dim pa As TAlarm
Dim sz As String

    With mAlarmList
        .Rewind
        Do While .GetNextTag(pa) = B_OK
            sz = sz & pa.Details & "#?0#?" & pa.Label & "|"

        Loop

    End With

    prefskit_SafeSetText mPanel, "alarms_list", g_SafeLeftStr(sz, Len(sz) - 1)

    ' /* icons */

Dim i As Long

    With mAlarmList
        If .CountItems Then
            For i = 1 To .CountItems
                Set pa = .TagAt(i)
                prefskit_SetItemEx mPanel, "alarms_list", "image-file", i, g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "alarm.png"
                prefskit_SetItemEx mPanel, "alarms_list", "checked", i, IIf(pa.IsEnabled, 1&, 0&)

            Next i
        End If
    End With

End Sub

'Private Sub uRecalcReminders()
'Dim pr As TReminder
'
'    With mRemList
'        .Rewind
'        Do While .GetNextTag(pr) = B_OK
'            pr.CalcNextDue
'
'        Loop
'
'    End With
'
'    uUpdateReminderList
'
'End Sub

Private Function uGetReminderOffset() As Long

    uGetReminderOffset = g_SafeLong(mConfigFile.ValueOf("reminder_offset"))

End Function

Private Function uGetPackedDate(ByVal aDate As Date) As String

    With New CTime
        .SetTo aDate
        
        ' /* YYYYMMDD */
        uGetPackedDate = Format$(.Year, "0000") & Format$(.Month, "00") & Format$(.Day, "00")

        ' /* HH */
        If mConfigFile.ValueOf("use_12_hour_clock") = "1" Then
            uGetPackedDate = uGetPackedDate & Format$(.Hour12, "00")

        Else
            uGetPackedDate = uGetPackedDate & Format$(.Hour, "00")

        End If

        ' /* MM */
        uGetPackedDate = uGetPackedDate & Format$(.Minutes, "00")

        ' /* SS */
        uGetPackedDate = uGetPackedDate & IIf(mConfigFile.ValueOf("show_seconds") = "1", Format$(.Seconds, "00"), "XX")

    End With

End Function
