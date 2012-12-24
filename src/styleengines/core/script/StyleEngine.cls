VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const LIB_DATE = "24-Feb-2012"

Dim mLastErr As String
Dim mPanel As BPrefsPanel
Dim mTest As IStyleInstance
Dim mStyles As BTagList

Implements IStyleEngine
Implements KPrefsPanel
Implements KPrefsPage

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = mStyles.CountItems

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance
Dim ps As TScriptStyle
Dim pi As TInstance

    Set ps = mStyles.TagAt(mStyles.IndexOf(StyleName, False))
    If ISNULL(ps) Then
        g_Debug "IStyleEngine.CreateInstance(): failed on style '" & StyleName & "'", LEMON_LEVEL_CRITICAL

    Else
        g_Debug "IStyleEngine.CreateInstance(): creating new instance of '" & StyleName & "'..."
        Set pi = New TInstance
        pi.SetTo ps
        Set IStyleEngine_CreateInstance = pi

    End If

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = "Supports VBScript-based styles (Alpha release for comment)"

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long
Dim ps As TScriptStyle

    Set ps = mStyles.TagAt(mStyles.IndexOf(StyleName, False))
    If NOTNULL(ps) Then
        If MsgBox("This is a scripted style.  Would you like to open the script so you can edit it?", vbQuestion Or vbYesNo, "Scripted Styles") = vbYes Then _
            ShellExecute 0, "open", "notepad.exe", ps.ScriptPath, vbNullString, 5

    End If

'Dim pp As BPrefsPage
'Dim pm As CTempMsg
'
'    If (mPanel Is Nothing) Then
'
''        mPanelStyle = StyleName
'
'        Set mPanel = New BPrefsPanel
'        With mPanel
'            .SetHandler Me
'            .SetTitle "Prowl Preferences"
'            .SetWidth 300
'            .SetWindow 1
'
'            Set pp = new_BPrefsPage("General", load_image_obj(g_MakePath(App.Path) & "general.png"), Me)
'            With pp
'                .SetMargin 88
'
''            .Add new_BPrefsControl("banner", "", "User Key")
''                .Add new_BPrefsControl("fancyedit", "UserKey", gConfig.UserKey, "User key:")
''                .Add new_BPrefsControl("label", "", "This is the user key you generated at http://prowl.weks.net/")
''
'''            .Add new_BPrefsControl("banner", "", "Options")
''                .Add new_BPrefsControl("fancytoggle2", "OnlyShowPriorityNotifications", "Only show priority notifications?", "Options:", IIf(gConfig.OnlyShowPriorityNotifications, "1", "0"))
''                .Add new_BPrefsControl("fancytoggle2", "ReplaceCRLFs", "Replace CR/LFs with spaces?", "", IIf(gConfig.ReplaceCRLFs, "1", "0"))
''                .Add new_BPrefsControl("fancytoggle2", "RedactSensitive", "Redact text in sensitive notifications?", "", IIf(gConfig.RedactSensitive, "1", "0"))
''
''                .Add new_BPrefsControl("fancyedit", "AppText", gConfig.AppText, "Application:")
''                .Add new_BPrefsControl("label", "", "Prowl supports an entry which indicates the sending application.  You can use the special values %APP% and %COMPUTER% to indicate the sending application and the name of this computer respectively")
''
''                .Add new_BPrefsControl("separator", "")
''                .Add new_BPrefsControl("fancybutton2", "test", "Test Settings")
'
'            End With
'            .AddPage pp
'
'            Set pp = new_BPrefsPage("Network", load_image_obj(g_MakePath(App.Path) & "network.png"), Me)
'            With pp
'                .SetMargin 88
'
'                Set pm = New CTempMsg
'                pm.Add "min", 1&
'                pm.Add "max", 9&
'                pm.Add "freq", 1&
'                pm.Add "label_size", 64&
'                .Add new_BPrefsControl("fancyslider", "Timeout", "", "Timeout:", CStr(-1), pm)
'                .Add new_BPrefsControl("label", "", "This is the maximum time to wait while sending the notification to the Prowl servers.  You should try to keep this setting as low as possible.")
''
''                .Add new_BPrefsControl("fancytoggle2", "UseProxyServer", "Use a proxy server?", "Proxy:", IIf(gConfig.UseProxyServer, "1", "0"))
''                .Add new_BPrefsControl("fancyedit", "ProxyServer", gConfig.ProxyServer, "Name:")
''                .Add new_BPrefsControl("fancyedit", "ProxyServerPort", CStr(gConfig.ProxyServerPort), "Port:")
''                .Add new_BPrefsControl("label", "", "Leave the following blank if proxy autentication is not required:")
''                .Add new_BPrefsControl("fancyedit", "ProxyUsername", gConfig.ProxyUsername, "Username:")
'
''                Set pm = New CTempMsg
''                pm.Add "password", -1
''                .Add new_BPrefsControl("fancyedit", "ProxyPassword", gConfig.ProxyPassword, "Password:", , pm)
'
'            End With
'            .AddPage pp
'
'            Set pp = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & "icon.png"), Me)
'            With pp
'                .SetMargin 32
'                .Add new_BPrefsControl("banner", "", "About Prowl")
'                .Add new_BPrefsControl("label", "", "Prowl is an iPod Touch/iPhone push message service created by Zachary West.  To get Prowl to work, you must:")
'                .Add new_BPrefsControl("label", "", "o Own an iPod Touch or iPhone")
'                .Add new_BPrefsControl("label", "", "o Purchase and install the iPhone Prowl client from the App Store")
'                .Add new_BPrefsControl("label", "", "o Create a (free) account at prowl.weks.net")
'                .Add new_BPrefsControl("label", "", "o Generate your unique user key at the same site")
'                .Add new_BPrefsControl("label", "", "o Enter this user key in the text box above")
'
'                .Add new_BPrefsControl("separator", "")
'                .Add new_BPrefsControl("label", "", "More detailed instructions are available at http://prowl.weks.net")
'                .Add new_BPrefsControl("fancybutton2", "go_prowl_site", "Visit prowl.weks.net")
'
'            End With
'            .AddPage pp
'
'            .Go
''            g_SetWindowIconToAppResourceIcon mPanel.hWnd
'
'        End With
'    End If
'
'    IStyleEngine_GetConfigWindow = mPanel.hWnd

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    l3OpenLog "%APPDATA%\full phat\snarl\" & "script_style_engine.log", True
    g_Debug "IStyleEngine.Initialize()", LEMON_LEVEL_PROC_ENTER
    g_Debug "script.styleengine V" & CStr(App.Major) & "." & CStr(App.Revision) & " (" & LIB_DATE & ")"

    ' /* return code isn't really important so long as it's an error code (i.e. <> 0) */

    IStyleEngine_Initialize = M_FAILED

    mLastErr = "Needs misc2.resource V49 or better"
    If Not uMiscResourceOk() Then
        g_Debug "misc2.resource not found", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    Set mStyles = new_BTagList()

    ' /* get scripts */


    ' /* if in the IDE, current path */

    If g_IsIDE() Then _
        uLoadScriptsFrom g_MakePath(App.Path)

Dim sz As String

    ' /* all users */

    If style_GetSnarlStylesPath2(True, sz) Then
        uLoadScriptsFrom sz

    Else
        g_Debug "no all users styles folder"

    End If

    ' /* current user */

    If style_GetSnarlStylesPath2(False, sz) Then
        uLoadScriptsFrom sz

    Else
        g_Debug "no user styles folder"

    End If


    If mStyles.CountItems = 0 Then
        g_Debug "no styles found", LEMON_LEVEL_WARNING

    Else
        g_Debug CStr(mStyles.CountItems) & " style(s) loaded", LEMON_LEVEL_INFO

    End If


'    g_Debug "setting defaults..."
'
'    With gConfig
'        .OnlyShowPriorityNotifications = False
'        .UseProxyServer = False
'        .ProxyServerPort = 80
'        .ReplaceCRLFs = True
'        .Timeout = 5
'
'        .AppText = "%APP% on %COMPUTER%"
'        .RedactSensitive = False
'
'    End With

    ' /* load config here */

'Dim sz As String
Dim i As Long

'    With New ConfigFile
'        .File = style_GetSnarlConfigPath("prowl")
'        g_Debug "looking for config '" & .File & "'"
'        If .Load Then
'            g_Debug "got config..."
'
'            i = .FindSection("general")
'            If i Then
'                g_Debug "got [general] section..."
'
'                With .SectionAt(i)
'                    If .Find("UserKey", sz) Then _
'                        gConfig.UserKey = sz
'
'                    If .Find("OnlyShowPriorityNotifications", sz) Then _
'                        gConfig.OnlyShowPriorityNotifications = (Val(sz) <> 0)
'
'                    If .Find("ReplaceCRLFs", sz) Then _
'                        gConfig.ReplaceCRLFs = (Val(sz) <> 0)
'
'                    If .Find("AppText", sz) Then _
'                        gConfig.AppText = sz
'
'                    If .Find("RedactSensitive", sz) Then _
'                        gConfig.RedactSensitive = (Val(sz) <> 0)
'
'                End With
'
'            Else
'                g_Debug "missing [general] section", LEMON_LEVEL_WARNING
'
'            End If
'
'            i = .FindSection("network")
'            If i Then
'                g_Debug "got [network] section..."
'
'                With .SectionAt(i)
'                    If .Find("UseProxyServer", sz) Then _
'                        gConfig.UseProxyServer = (Val(sz) <> 0)
'
'                    If .Find("ProxyServer", sz) Then _
'                        gConfig.ProxyServer = sz
'
'                    If .Find("ProxyServerPort", sz) Then _
'                        gConfig.ProxyServerPort = Val(sz)
'
'                    If .Find("ProxyUsername", sz) Then _
'                        gConfig.ProxyUsername = sz
'
'                    If .Find("ProxyPassword", sz) Then _
'                        gConfig.ProxyPassword = sz
'
'                    If .Find("Timeout", sz) Then _
'                        gConfig.Timeout = Val(sz)
'
'                End With
'
'            Else
'                g_Debug "missing [network] section", LEMON_LEVEL_WARNING
'
'            End If
'
'            ' /* validate loaded settings */
'
'            If gConfig.Timeout < 1 Then
'                g_Debug "invalid timeout '" & CStr(gConfig.Timeout) & "' - resetting to '1", LEMON_LEVEL_INFO
'                gConfig.Timeout = 1
'
'            ElseIf gConfig.Timeout > 9 Then
'                g_Debug "invalid timeout '" & CStr(gConfig.Timeout) & "' - resetting to '9", LEMON_LEVEL_INFO
'                gConfig.Timeout = 9
'
'            End If
'
'        Else
'            g_Debug "failed to load", LEMON_LEVEL_WARNING
'
'        End If
'
'    End With

    mLastErr = ""
    IStyleEngine_Initialize = M_OK
    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function IStyleEngine_LastError() As String

    IStyleEngine_LastError = mLastErr

End Function

Private Function IStyleEngine_Name() As String

    IStyleEngine_Name = App.Title

End Function

Private Function IStyleEngine_Path() As String

    IStyleEngine_Path = App.Path

End Function

Private Function IStyleEngine_Revision() As Long

    IStyleEngine_Revision = App.Revision

End Function

Private Sub IStyleEngine_StyleAt(ByVal Index As Long, Style As libSnarlStyles.style_info)

    If (Index < 1) Or (Index > mStyles.CountItems) Then
        g_Debug "IStyleEngine_StyleAt(): not supplying a style for index #" & CStr(Index), LEMON_LEVEL_WARNING
        Exit Sub

    End If

Dim ps As TScriptStyle

    Set ps = mStyles.TagAt(Index)
    LSet Style = ps.Info

    g_Debug "IStyleEngine_StyleAt(): style at index #" & CStr(Index) & " is '" & Style.Name & "' " & Style.Major & "." & Style.Minor

End Sub

Private Sub IStyleEngine_TidyUp()
Dim ps As TScriptStyle

    With mStyles
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            ps.TidyUp
            
        Loop

    End With

    Set mStyles = Nothing

    g_Debug "IStyleEngine_TidyUp(): done"
    l3CloseLog

End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function

Private Sub KPrefsPage_AllAttached()
Dim pc As BControl

'    If mPanel.Find("Timeout", pc) Then _
        pc.SetValue CStr(gConfig.Timeout)

End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

'    Select Case Control.getName()
'    Case "UserKey"
''        gConfig.UserKey = Value
'
'    Case "OnlyShowPriorityNotifications"
'        gConfig.OnlyShowPriorityNotifications = (Value = "1")
'
'    Case "ReplaceCRLFs"
'        gConfig.ReplaceCRLFs = (Value = "1")
'
'
'    Case "UseProxyServer"
'        gConfig.UseProxyServer = (Value = "1")
'
'    Case "ProxyServer"
'        gConfig.ProxyServer = Value
'
'    Case "ProxyServerPort"
'        If g_IsNumeric(Value) Then _
'            gConfig.ProxyServerPort = Val(Value)
'
'    Case "ProxyUsername"
'        gConfig.ProxyUsername = Value
'
'    Case "ProxyPassword"
'        gConfig.ProxyPassword = Value
'
'    Case "Timeout"
'        Control.SetText Value & " second" & IIf(Value <> "1", "s", "")
'        gConfig.Timeout = Val(Value)
'
'
'    Case "AppText"
'        gConfig.AppText = Value
'
'    Case "RedactSensitive"
'        gConfig.RedactSensitive = (Value = "1")
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
'Dim pn As notification_info
'
'    Select Case Control.getName()
'    Case "go_prowl_site"
'        ShellExecute mPanel.hWnd, "open", "http://prowl.weks.net/", vbNullString, vbNullString, 5
'
'    Case "test"
'        With pn
'            .Title = "Prowl Style"
'            .Text = "title::Test Notification#?text::This is a test notification from the Prowl Style"
'
'        End With
'
'        Set mTest = New StyleInstance
'        mTest.UpdateContent pn
'
'        MsgBox "A test notification has been sent", vbInformation Or vbOKOnly, "Prowl Style"
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
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

'    g_WriteConfig
    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Function uMiscResourceOk() As Boolean
Dim pt As BTimer

    On Error Resume Next
    Err.Clear

    Set pt = New BTimer
    If (Err.Number = 0) And (NOTNULL(pt)) Then _
        uMiscResourceOk = True

End Function

Private Sub uLoadScriptsFrom(ByVal Path As String)
Dim ps As TScriptStyle
Dim sz As String

    g_Debug "StyleEngine.uLoadScriptsFrom()", LEMON_LEVEL_PROC_ENTER

    Path = g_MakePath(Path) & "script"

    With New CFolderContent2
        If .SetTo(Path) Then
            g_Debug "looking in '" & Path & "'..."
            .Rewind
            Do While .GetNextFile(sz)
                If (g_IsFolder(sz)) And (g_Exists(g_MakePath(sz) & "script.vbs")) Then
                    Set ps = New TScriptStyle
                    If ps.InitFrom(sz) Then
                        If mStyles.IndexOf(ps.Name, False) = 0 Then
                            mStyles.Add ps
                            
                        Else
                            g_Debug "not added: " & g_Quote(ps.Name) & " already loaded", LEMON_LEVEL_WARNING

                        End If
                    End If
                End If
            Loop

        Else
            g_Debug Path & " doesn't exist", LEMON_LEVEL_WARNING

        End If
        
    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

