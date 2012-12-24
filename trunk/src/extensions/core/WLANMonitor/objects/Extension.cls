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
    '/  Created:        15-Feb-2010
    '/
    '/  © 2010 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* our constants */

Const EXT_NAME = "WiFi Monitor"
Const EXT_DATE = "6-June-2011"
Const EXT_REL = "1.1"

    ' /*
    '
    '
    '   IMPORTANT: This is a fully functional Snarl extension - if you're planning on
    '   writing your own extension you should use the template extension in the Snarl SDK.
    '
    ' */

Dim mLastErr As String

Dim mPanel As BPrefsPanel           ' // our prefs panel
Dim mhWndOwner As Long              ' // handle of the window which launched our config (typically Snarl's Prefs panel)

Dim mService As SWbemServices
Dim mInstance() As TWLANInstance
Dim mCount As Long

Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

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
            .SetTitle EXT_NAME & " Preferences"
            .SetWidth 380
            .SetWindow 1                    ' // so calling Go() doesn't show the window

            Set pPage = new_BPrefsPage("", , Me)

            ' /* General */

            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("banner", "", "WiFi Interfaces")

                Set pm = New CTempMsg
                pm.Add "checkboxes", 1
                pm.Add "item-height", 38
                Set pc = new_BPrefsControl("listbox", "interface_list", , , "1", pm)
                pc.SizeTo 0, 160
                .Add pc

                .Add new_BPrefsControl("label", "", "Note that both virtual and physical WiFi interfaces are listed here.")


                .Add new_BPrefsControl("banner", "", "Scan Frequency")
                Set pm = New CTempMsg
                pm.Add "min", 1&
                pm.Add "max", 60&
                pm.Add "freq", 5&
                pm.Add "label_size", 50&
                .Add new_BPrefsControl("fancyslider", "scan_frequency", , "", gConfig.ValueOf("scan_frequency"), pm)

                .Add new_BPrefsControl("banner", "", "Sensitivity")
'                Set pm = New CTempMsg
'                pm.Add "min", 1&
'                pm.Add "max", 4&
'                pm.Add "freq", 1&
'                .Add new_BPrefsControl("fancyslider", "sensitivity", , "", IIf(gConfig.ValueOf("sensitivity") = "1", "4", "1"), pm)

                .Add new_BPrefsControl("fancycycle", "sensitivity", "Very Low|Low|Medium|High")


'            Set pm = New CTempMsg
'            pm.Add "inset_by", CLng(Fix(.GetMargin() / 2))

'                .Add new_BPrefsControl("fancytoggle2", "show_time_on_load", "Show current time?", , IIf(gConfig.show_time_on_load, "1", "0"))
'                .Add new_BPrefsControl("fancytoggle2", "show_cookie_on_load", "Display random cookie?", , IIf(gConfig.show_cookie_on_load, "1", "0"))
'
'                .Add new_BPrefsControl("banner", "", "Cookies")
'                .Add new_BPrefsControl("fancytoggle2", "show_cookie", "Include random cookie in notification?", , IIf(gConfig.show_cookie, "1", "0"))
'
'                sz = g_GetPath(gConfig.cookie_file)
'                If sz <> "" Then
'                    Set pm = New CTempMsg
'                    pm.Add "path", sz
'
'                End If
'
'                .Add new_BPrefsControl("fancyfilepicker", "cookie_file", gConfig.cookie_file, "Cookie file:", , pm)
'
'                .Add new_BPrefsControl("banner", "", "About " & EXT_NAME)
'                .Add new_BPrefsControl("label", "", "V" & CStr(App.Major) & "." & CStr(App.Revision), "Version:")
'                .Add new_BPrefsControl("fancybutton2", "go_website", "www.fullphat.net", "Website:")

            End With

            .AddPage pPage

'
'                Set pm = New CTempMsg
'                pm.Add "min", 0&
'                pm.Add "max", 59&
'                pm.Add "freq", 5&
'                .Add new_BPrefsControl("fancyslider", "timer_mins", , "Minutes:", CStr(gConfig.timer_mins), pm)
'                .Add new_BPrefsControl("fancytoggle2", "timer_10sec_warning", "Show notification during last 10 seconds?", "Other:", IIf(gConfig.timer_10sec_warning, "1", "0"))
'                Set pc = new_BPrefsControl("fancyedit", "pm_val", gConfig.pm_val, "PM:")
'                pc.SizeTo 48, pc.Height
'                .Add pc

            ' /* create it */

            .Go
            g_SetWindowIconToAppResourceIcon .hwnd

        End With

    End If

    ISnarlExtension_GetConfigWindow = mPanel.hwnd

End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "full phat products"
        .Copyright = App.LegalCopyright
        .Date = EXT_DATE
        .Description = App.FileDescription
        .Flags = SNARL_EXTN_IS_CONFIGURABLE
        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Name = EXT_NAME
        .Path = App.Path
        .Release = EXT_REL
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With

End Sub

Private Function ISnarlExtension_Initialize() As Long

    l3OpenLog "%APPDATA%\" & App.Title & ".log"
    g_Debug EXT_NAME & " running from '" & g_MakePath(App.Path) & "'"
    g_Debug ""

    g_Debug "ISnarlExtension_Initialize"
    mLastErr = ""

    ' /* does snarl's /etc/ path exist? */

Dim szPath As String

    If Not snExt_GetUserPath(szPath) Then
        mLastErr = "Couldn't get Snarl user path"
        ISnarlExtension_Initialize = 201
        Exit Function

    End If

    ' /* we got WMI? */

Dim pLocator As SWbemLocator

    Err.Clear
    Set pLocator = New SWbemLocator
    If (Err.Number <> 0) Or (pLocator Is Nothing) Then
        g_Debug "failed to get locator object", LEMON_LEVEL_CRITICAL
        mLastErr = "WMI initialization error"
        ISnarlExtension_Initialize = 201
        Exit Function

    End If

    g_Debug "got WMI locator..."

    Err.Clear
    Set mService = pLocator.ConnectServer(, "root\wmi")
    If (Err.Number <> 0) Or (mService Is Nothing) Then
        g_Debug "failed to get root service", LEMON_LEVEL_CRITICAL
        mLastErr = "WMI initialization error"
        ISnarlExtension_Initialize = 201
        Exit Function

    End If

    g_Debug "got WMI root service..."

    ' /* do a quick scan to see if we actually have any devices */

    uScan

'    If mCount = 0 Then
'        g_Debug "no devices found", LEMON_LEVEL_CRITICAL
'        mLastErr = "No WiFi cards found"
'        ISnarlExtension_Initialize = 201
'        Exit Function
'
'    End If
'
End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
End Sub

Private Sub ISnarlExtension_Start()

    On Error Resume Next

    g_Debug "ISnarlExtension_Start"

    ' /* register */
    gAppToken = sn41RegisterApp("application/vnd.k23-" & App.Title, EXT_NAME, g_MakePath(App.Path) & "icon.png")
    If gAppToken = 0 Then
        g_Debug "failed to register with Snarl", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    g_Debug "registered with Snarl ok..."

'    sn41AddClass gapptoken, CLASS_HOUR_CHIME, "Hourly chime"

    ' /* load config */

Dim szPath As String
Dim bSave As Boolean
    
    snExt_GetUserPath szPath

    Set gConfig = New CConfFile
    With gConfig
        .SetTo g_MakePath(szPath) & "etc\." & App.Title
        If Not .Exists("scan_frequency") Then
            .Add "scan_frequency", "2"
            bSave = True

        End If

        If Not .Exists("sensitivity") Then
            .Add "sensitivity", "3"
            bSave = True

        End If

        If bSave Then _
            .Save

    End With

    ' /* first scan */

    uScan

    Set theTimer = New BTimer
    theTimer.SetTo Val(gConfig.ValueOf("scan_frequency")) * 1000

End Sub

Private Sub ISnarlExtension_Stop()

    g_Debug "ISnarlExtension_Stop"

    If Not (theTimer Is Nothing) Then _
        theTimer.SetEnabled False

    sn41UnregisterApp gAppToken

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
Dim i As Long

    uUpdateList

    If mPanel.Find("sensitivity", pc) Then _
        pc.SetValue CStr(5 - Val(gConfig.ValueOf("sensitivity")))

    i = Val(gConfig.ValueOf("scan_frequency"))
    If mPanel.Find("scan_frequency", pc) Then _
        pc.SetText CStr(i) & " sec" & IIf(i = 1, "", "s")

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim i As Long

'    g_Debug "ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName

    Case "sensitivity"
        ' /* swap the values around so 5 == Low, etc. */
        gConfig.Update Control.GetName, CStr(5 - Val(Value))
        gConfig.Save

    Case "scan_frequency"
        gConfig.Update Control.GetName, Value
        gConfig.Save

        i = Val(Value)
        Control.SetText CStr(i) & " sec" & IIf(i = 1, "", "s")
        theTimer.SetTo i * 1000

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

'    Case "test"
'        sn41EZNotify 0, "", "Time format preview", "It's now " & g_GetTimeString(gConfig, False), -1, g_MakePath(App.Path) & "icon.png"

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Select Case Control.GetName
    Case "interface_list"
        Select Case Notification
        Case "update_list"
            uUpdateList

        Case "checked"
            ' /* set the enabled/disabled state in the config */
            gConfig.Update "*" & prefskit_GetItem(Control, , Val(Control.GetValue)), prefskit_GetItem(Control, "checked", Val(Control.GetValue))
            gConfig.Save

        End Select

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

Private Sub theTimer_Pulse()

    uScan

End Sub

Private Sub uScan()
Dim pCollection As SWbemObjectSet

    On Error Resume Next
' Where active=true
    Err.Clear
    Set pCollection = mService.ExecQuery("SELECT * FROM MSNdis_80211_ReceivedSignalStrength")
    If (Err.Number <> 0) Or (pCollection Is Nothing) Then _
        Exit Sub

Dim pItem As SWbemObject
Dim pProp As SWbemProperty
Dim i As Long

    For Each pItem In pCollection
        Set pProp = pItem.Properties_.Item("InstanceName")
        If Not (pProp Is Nothing) Then
            i = uIndexOf(pProp.Value)
            If i = 0 Then
                uAdd pItem

            ElseIf gConfig.ValueOf("*" & pProp.Value) = "1" Then
                ' /* only check if it's enabled in the config */
                mInstance(i).Check pItem

            End If

        End If

'        For Each pProp In pItem.Properties_
'            Debug.Print pProp.Name & " --> " & pProp.Value
'
'        Next

    Next

End Sub

Private Function uIndexOf(ByVal Name As String) As Long
Static i As Long

    If mCount Then
        For i = 1 To mCount
            If mInstance(i).Name = Name Then
                uIndexOf = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Sub uAdd(ByRef Instance As SWbemObject)

    mCount = mCount + 1
    ReDim Preserve mInstance(mCount)
    Set mInstance(mCount) = New TWLANInstance
    mInstance(mCount).SetTo Instance

Dim i As E_WLAN_STRENGTH

    i = mInstance(mCount).Strength
    Debug.Print "Added " & mInstance(mCount).Name & " > " & i

Dim sz As String

    ' /* is it in our config? */

    sz = "*" & mInstance(mCount).Name
    If Not gConfig.Exists(sz) Then
        gConfig.Add sz, "1"
        gConfig.Save

    End If

End Sub

Private Sub uUpdateList()
Dim pc As BControl

    If Not (mPanel.Find("interface_list", pc)) Then _
        Exit Sub

Dim iPrev As String

    iPrev = pc.GetValue()           ' // so we can re-selected the currently selected item...
    pc.DoExCmd "lock"

Dim sz As String
Dim pm As CTempMsg
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            sz = sz & mInstance(i).Name & "|"

        Next i

        ' /* set the content */

        pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

    ' /* set the checkboxes and icons */

        For i = 1 To mCount
            Set pm = New CTempMsg
            pm.Add "index", i
            pm.Add "checked", CLng(Val(gConfig.ValueOf("*" & mInstance(i).Name)))
            pm.Add "image-file", g_MakePath(App.Path) & "icon.png"
            pc.DoExCmd B_SET_ITEM, pm

        Next i

        pc.SetValue iPrev

    Else
        ' /* no instances so clear the list */
        pc.SetText ""

    End If

    pc.DoExCmd "unlock"

End Sub

