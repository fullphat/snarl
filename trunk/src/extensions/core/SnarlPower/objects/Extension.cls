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
    '/  Description:    SnarlPower Extension main object
    '/
    '/  Created:        17-Feb-2010
    '/
    '/  © 2010 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /*
    '
    '
    '
    ' */

Private Const OBJ_DATE = "28-Jun-2010"
Private Const CLASS_POWER_CONNECT = "pwrc"
Private Const CLASS_POWER_DISCONNECT = "pwrd"
Private Const CLASS_POWER_LOW = "pwrl"

Private Const SW_SHOW = 5


Dim mAppToken As Long
Dim mLastErr As String

Dim mConfigFile As ConfigFile
Dim mConfig As ConfigSection

Dim mPanel As BPrefsPanel           ' // our prefs panel

Dim mNotificationId As Long
Dim mLowPower As Long

Dim WithEvents theHandler As TWindow
Attribute theHandler.VB_VarHelpID = -1

Implements KPrefsPanel              ' // we're a panel handler
Implements KPrefsPage               ' // we're also a page handler
Implements ISnarlExtension

Private Function ISnarlExtension_GetConfigWindow() As Long

'    mIgnoreChange = True

Dim pPage As BPrefsPage
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetTitle App.Title & " Preferences"
            .SetWidth 380
            .SetHandler Me
            .SetWindow 1

            Set pPage = new_BPrefsPage("", Nothing, Me)
            With pPage
                .SetMargin 72

                .Add new_BPrefsControl("banner", "", "Low power threshold:")
                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 90&
                pm.Add "freq", 10&
                pm.Add "label_size", 50&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "LowPowerThreshold", "", "", mConfig.GetValueWithDefault("LowPowerThreshold", "25"), pm)
                .Add new_BPrefsControl("fancytoggle2", "ShowWarningEvenIfCharging", "Show notification even if charging?", , mConfig.GetValueWithDefault("ShowWarningEvenIfCharging", "0"))

                .Add new_BPrefsControl("banner", "", "Appearance")
                .Add new_BPrefsControl("fancytoggle2", "ShowText", "Include text?", , mConfig.GetValueWithDefault("ShowText", "0"))
                .Add new_BPrefsControl("fancytoggle2", "MeterStyleMode", "Meter friendly notifications?", , mConfig.GetValueWithDefault("MeterStyleMode", "0"), , mConfig.GetValueWithDefault("ShowText", "0") = "1")

                .Add new_BPrefsControl("banner", "", "About " & App.Title)
                .Add new_BPrefsControl("label", "", "V" & CStr(App.Major) & "." & CStr(App.Revision), "Version:")
                .Add new_BPrefsControl("label", "", App.LegalCopyright)
                .Add new_BPrefsControl("fancybutton2", "go_website", "www.fullphat.net", "Website:")

'        .Add new_BPrefsControl("fancybutton2", "low", "Test Low Power")
''        .Add new_BPrefsControl("fancybutton2", "disc", "Test Disconnected")
''        .Add new_BPrefsControl("fancybutton2", "conn", "Test Connected")

            End With

            .AddPage pPage
            .Go

'            g_SetWindowIconToAppResourceIcon .hWnd

        End With
    End If

    ISnarlExtension_GetConfigWindow = mPanel.hWnd

End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "full phat products"
        .Copyright = App.LegalCopyright
        .Date = OBJ_DATE
        .Description = App.FileDescription
        .Flags = SNARL_EXTN_IS_CONFIGURABLE
        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Name = App.Title
        .Path = g_MakePath(App.Path)
        .Release = "1.0 RC1"
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With

End Sub

Private Function ISnarlExtension_Initialize() As Long

    On Error Resume Next

    l3OpenLog "%APPDATA%\" & App.Title & ".log"

    g_Debug "ISnarlExtension.Initialize()", LEMON_LEVEL_PROC_ENTER
    
    g_Debug "pwd is '" & g_MakePath(App.Path) & "'"

    ' /* check misc.resource */

Dim pl As BMsgLooper

    g_Debug "checking misc.resource..."

    Set pl = New BMsgLooper
    If Err.Number <> 0 Then
        mLastErr = "Needs misc.resource V48 or later"
        g_Debug "failed"
        ISnarlExtension_Initialize = 200
        Exit Function

    End If

    g_Debug "ok"

    ' /* does snarl's /etc/ path exist? */

Dim szPath As String

    g_Debug "checking settings folder..."

    If Not snext_getuserpath(szPath) Then
        mLastErr = "Settings folder missing"
        ISnarlExtension_Initialize = 201
        Exit Function

    End If

    g_Debug "ok"

    Set mConfigFile = New ConfigFile
    mConfigFile.File = g_MakePath(szPath) & "etc\.snarlpower"

    ' /* create window */

    g_Debug "creating window..."

    Set theHandler = New TWindow
    If Not theHandler.IsValid Then
        ' /* no window */
        Set theHandler = Nothing
        mLastErr = "Couldn't create message loop"
        ISnarlExtension_Initialize = 202
        Exit Function

    End If

Dim b As Boolean
Dim i As Long

    i = get_battery_charge_percent(b)
    g_Debug "battery_percent: " & i & " " & b
    g_Debug "power_status: " & get_power_status()

    g_Debug "ok", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
End Sub

Private Sub ISnarlExtension_Start()

    g_Debug "ISnarlExtension.Start()", LEMON_LEVEL_PROC_ENTER

    mAppToken = sn41RegisterApp(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png")

g_Debug mAppToken

    If mAppToken Then
        sn41AddClass mAppToken, CLASS_POWER_CONNECT, "Power connected"
        sn41AddClass mAppToken, CLASS_POWER_DISCONNECT, "Power disconnected"
        sn41AddClass mAppToken, CLASS_POWER_LOW, "Low power warning"

    Else
        g_Debug "No app token!", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If

    g_Debug "loading config..."

Dim i As Long

    With mConfigFile
        .Load

        i = .FindSection("general")
        If i = 0 Then
            Set mConfig = .AddSectionObj("general")
            g_Debug "adding [general] section..."

        Else
            Set mConfig = .SectionAt(i)

        End If

        ' /* defaults */

        g_Debug "checking config..."

        With mConfig
            If .IndexOf("LowPowerThreshold") = 0 Then _
                .Add "LowPowerThreshold", "25"

            If .IndexOf("MeterStyleMode") = 0 Then _
                .Add "MeterStyleMode", "1"                      ' // was 0 in previous release

            If .IndexOf("ShowWarningEvenIfCharging") = 0 Then _
                .Add "ShowWarningEvenIfCharging", "0"

            ' /* new in 1.0 */

            If .IndexOf("ShowText") = 0 Then _
                .Add "ShowText", "0"

            ' /* our window needs to know this value */

            theHandler.SetLowPowerThreshold Val(.GetValueWithDefault("LowPowerThreshold", "25"))

        End With

        .Save

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub ISnarlExtension_Stop()

    g_Debug "ISnarlExtension.Stop()", LEMON_LEVEL_PROC_ENTER

    sn41UnregisterApp mAppToken
    mAppToken = 0

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub ISnarlExtension_TidyUp()

    g_Debug "ISnarlExtension.TidyUp()", LEMON_LEVEL_PROC_ENTER

    If Not (mPanel Is Nothing) Then _
        mPanel.Quit

    Set theHandler = Nothing

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub theHandler_OnLowPowerWarning()

    g_Debug "theHandler.OnLowPowerWarning()", LEMON_LEVEL_PROC_ENTER

    ' /* show even if charging? */

    If (get_power_status = B_POWER_STATE_CONNECTED) And (mConfig.GetValueWithDefault("ShowWarningEvenIfCharging", "0") = "0") Then _
        Exit Sub

Dim sz As String

    If mConfig.GetValueWithDefault("ShowText", "0") = "1" Then
        g_Debug "'showtext' enabled..."
        sz = CStr(get_battery_charge_percent())

        If mConfig.GetValueWithDefault("MeterStyleMode", "0") = "0" Then _
            sz = "Low Power Warning\nCurrent battery charge is " & sz & "%"

    End If

    mLowPower = sn41EZNotify(mAppToken, _
                             CLASS_POWER_LOW, _
                             "", _
                             sz, _
                             0, _
                             "!battery-low", _
                             1)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub theHandler_OnPowerStatusChanged(ByVal Connected As Boolean)

    g_Debug "theHandler.OnPowerStatusChanged()", LEMON_LEVEL_PROC_ENTER

Dim sz As String

    If mConfig.GetValueWithDefault("ShowText", "0") = "1" Then
        g_Debug "'showtext' enabled..."

        sz = CStr(get_battery_charge_percent())
        If mConfig.GetValueWithDefault("MeterStyleMode", "0") = "0" Then _
            sz = IIf(Connected, "Power Connected", "Power Disconnected") & "\nCurrent battery charge is " & sz & "%"

    End If

    If sn41IsVisible(mNotificationId) Then _
        sn41Hide mNotificationId

    If Connected Then

        ' /* remove low power warning if it's still on screen */

        If sn41IsVisible(mLowPower) Then _
            sn41Hide mLowPower

        mNotificationId = sn41EZNotify(mAppToken, _
                                       CLASS_POWER_CONNECT, _
                                       "", _
                                       sz, _
                                       , _
                                       "!power-connected" _
                                       )

    Else
        mNotificationId = sn41EZNotify(mAppToken, _
                                       CLASS_POWER_DISCONNECT, _
                                       "", _
                                       sz, _
                                       0, _
                                       "!power-disconnected" _
                                       )

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()

'    mIgnoreChange = False

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pc As BControl

    g_Debug " ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    mConfig.Update Control.GetName, Value
    mConfigFile.Save

    Select Case Control.GetName
    Case "ShowText"
        If mPanel.Find("MeterStyleMode", pc) Then _
            pc.SetEnabled (Value = "1")

    Case "LowPowerThreshold"
        theHandler.SetLowPowerThreshold Val(Value)

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Debug.Print "ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()

    Case "low"
        theHandler_OnLowPowerWarning

    Case "disc"
        theHandler_OnPowerStatusChanged False

    Case "conn"
        theHandler_OnPowerStatusChanged True

    Case "go_website"
        ShellExecute 0, "open", "http://www.fullphat.net", vbNullString, vbNullString, SW_SHOW

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
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



