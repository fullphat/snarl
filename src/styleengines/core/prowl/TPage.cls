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

Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub KPrefsPage_AllAttached()
Dim pc As BControl

    If mPage.Find("Timeout", pc) Then _
        pc.SetValue CStr(gConfig.Timeout)

End Sub

Private Sub KPrefsPage_Attached()
Dim pm As CTempMsg

    With mPage

        Select Case LCase$(mPage.GetName())
        Case "general"
            .SetMargin 80

'            .Add new_BPrefsControl("banner", "", "User Key")
            .Add new_BPrefsControl("fancyedit", "UserKey", gConfig.UserKey, "User key:")
            .Add new_BPrefsControl("label", "", "This is the user key you generated at http://prowl.weks.net/")

'            .Add new_BPrefsControl("banner", "", "Options")
            .Add new_BPrefsControl("fancytoggle2", "OnlyShowPriorityNotifications", "Only show priority notifications?", "Options:", IIf(gConfig.OnlyShowPriorityNotifications, "1", "0"))
            .Add new_BPrefsControl("fancytoggle2", "ReplaceCRLFs", "Replace CR/LFs with spaces?", "", IIf(gConfig.ReplaceCRLFs, "1", "0"))

'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("fancybutton2", "restart_style_roster", "Restart Style Roster")

        Case "network"
            .SetMargin 80

            Set pm = New CTempMsg
            pm.Add "min", 1&
            pm.Add "max", 9&
            pm.Add "freq", 1&
            pm.Add "label_size", 64&
            .Add new_BPrefsControl("fancyslider", "Timeout", "", "Timeout:", CStr(-1), pm)
            .Add new_BPrefsControl("label", "", "This is the maximum time to wait while sending the notification to the Prowl servers.  You should try to keep this setting as low as possible.")

            .Add new_BPrefsControl("fancytoggle2", "UseProxyServer", "Use a proxy server?", "Proxy:", IIf(gConfig.UseProxyServer, "1", "0"))
            .Add new_BPrefsControl("fancyedit", "ProxyServer", gConfig.ProxyServer, "Name:")
            .Add new_BPrefsControl("fancyedit", "ProxyServerPort", CStr(gConfig.ProxyServerPort), "Port:")
            .Add new_BPrefsControl("label", "", "Leave the following blank if proxy autentication is not required:")
            .Add new_BPrefsControl("fancyedit", "ProxyUsername", gConfig.ProxyUsername, "Username:")

            Set pm = New CTempMsg
            pm.Add "password", -1
            .Add new_BPrefsControl("fancyedit", "ProxyPassword", gConfig.ProxyPassword, "Password:", , pm)

        Case "about"
            .SetMargin 32
            .Add new_BPrefsControl("banner", "", "About Prowl")
            .Add new_BPrefsControl("label", "", "Prowl is an iPod Touch/iPhone push message service created by Zachary West.  To get Prowl to work, you must:")
            .Add new_BPrefsControl("label", "", "o Own an iPod Touch or iPhone")
            .Add new_BPrefsControl("label", "", "o Purchase and install the iPhone Prowl client from the App Store")
            .Add new_BPrefsControl("label", "", "o Create a (free) account at prowl.weks.net")
            .Add new_BPrefsControl("label", "", "o Generate your unique user key at the same site")
            .Add new_BPrefsControl("label", "", "o Enter this user key in the text box above")

            .Add new_BPrefsControl("separator", "")
            .Add new_BPrefsControl("label", "", "More detailed instructions are available at http://prowl.weks.net")
            .Add new_BPrefsControl("fancybutton2", "go_prowl_site", "Visit prowl.weks.net")

        End Select

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Select Case Control.GetName()
    Case "UserKey"
        gConfig.UserKey = Value

    Case "OnlyShowPriorityNotifications"
        gConfig.OnlyShowPriorityNotifications = (Value = "1")

    Case "ReplaceCRLFs"
        gConfig.ReplaceCRLFs = (Value = "1")


    Case "UseProxyServer"
        gConfig.UseProxyServer = (Value = "1")

    Case "ProxyServer"
        gConfig.ProxyServer = Value

    Case "ProxyServerPort"
        If g_IsNumeric(Value) Then _
            gConfig.ProxyServerPort = Val(Value)

    Case "ProxyUsername"
        gConfig.ProxyUsername = Value

    Case "ProxyPassword"
        gConfig.ProxyPassword = Value

    Case "Timeout"
        Control.SetText Value & " second" & IIf(Value <> "1", "s", "")
        gConfig.Timeout = Val(Value)

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName()
    Case "go_prowl_site"
        ShellExecute mPage.Panel.hWnd, "open", "http://prowl.weks.net/", vbNullString, vbNullString, 5

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

