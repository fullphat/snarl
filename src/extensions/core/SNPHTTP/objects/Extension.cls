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
    '/  Created:        8-Apr-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

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
    '       'url' - return fully formatted URL to website (e.g. 'http://www.fullphat.net/')
    '
    '
    '   IMPORTANT: This is a fully functional Snarl extension - if you're planning on
    '   writing your own extension you should use the template extension in the Snarl SDK.
    '
    ' */

Private Const LIB_DATE = "23-Feb-2011"
Private Const LIB_RELEASE = "1.3"
Private Const EXT_NAME = "SNP/HTTP"
Private Const EXT_COMMENT = "Provides support for SNP 2.0 over HTTP"

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Dim mLastErr As String
Dim mhWnd As Long

Dim mConfig As CConfFile
Dim mPanel As BPrefsPanel           ' // our prefs panel

Dim mConn() As TConnection
Dim mCount As Long

Dim mPortChanged As Boolean
Dim mPortIsBad As Boolean

Implements KPrefsPanel              ' // we're a panel handler
Implements KPrefsPage               ' // we're also a page handler

Implements ISnarlExtension
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

End Function

Private Function ISnarlExtension_GetConfigWindow() As Long
Dim pPage As BPrefsPage
Dim pCtl As BControl

    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetWindow 1
            .SetTitle "SNP/HTTP Preferences"
            .SetWidth 178

            Set pPage = new_BPrefsPage("", , Me)
            With pPage
                .SetMargin 64
'                .SetWidth 200
'                .Add new_BPrefsControl("label", "", "This extension manages applications that register with Snarl using SNP 2.0 sent via HTTP.")

                .Add new_BPrefsControl("banner", "", "Connection")
                Set pCtl = new_BPrefsControl("fancyedit", "port", mConfig.GetValueWithDefault("port", "80"), "Port:")
                pCtl.SizeTo 48, pCtl.GetSize.y
                .Add pCtl

                .Add new_BPrefsControl("label", "", "The standard port for HTTP traffic is 80.  You can change this if it conflicts with an existing application (e.g. a web server).")

                If mPortIsBad Then
                    .Add new_BPrefsSeparator
                    .Add new_BPrefsControl("label", "", "SNP/HTTP was unable to bind to port to " & mConfig.GetValueWithDefault("port", "80") & ".  It's possible that another service is using this port.")

                End If

                .Add new_BPrefsControl("banner", "", "Formatting")
                .Add new_BPrefsControl("fancytoggle2", "html_response", "Include HTML in response?", , mConfig.GetValueWithDefault("html_response", "0"))

'                .Add new_BPrefsControl("banner", "", "Privacy")
'                .Add new_BPrefsControl("fancytoggle2", "require_password", "Require a password for application registration?", , mConfig.GetValueWithDefault("require_password", "0"))

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
        .Author = "fullphat products"
        .Copyright = App.LegalCopyright
        .Date = LIB_DATE
        .Description = EXT_COMMENT
        .Flags = SNARL_EXTN_IS_CONFIGURABLE
        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Name = EXT_NAME
        .Path = App.Path
        .Release = LIB_RELEASE
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With


End Sub

Private Function ISnarlExtension_Initialize() As Long

    On Error Resume Next

    l3OpenLog "%APPDATA%\full phat\snarl\snphttp.log", True
    g_Debug "pwd: " & g_MakePath(App.Path)

    If snDoRequest("version") = 0 Then
        g_Debug "Needs Snarl V42 or better", LEMON_LEVEL_CRITICAL
        mLastErr = "Requires Snarl V42 or greater"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

    ' /* load our config here so we can fail if we can't listen on the specified port */

'Dim szUser As String
Dim szPath As String

    Set mConfig = New CConfFile
    If snGetConfigPath(szPath) Then _
        mConfig.SetTo szPath & ".snphttp"

    Set theSocket = New CSocket

    If Not uStartListener Then
        mLastErr = "Couldn't create listener on port " & mConfig.GetValueWithDefault("port", "80")
        ISnarlExtension_Initialize = -1

    End If

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
End Sub

Private Sub ISnarlExtension_Start()
End Sub

Private Sub ISnarlExtension_Stop()
End Sub

Private Sub ISnarlExtension_TidyUp()

    theSocket.CloseSocket
    Set theSocket = Nothing

End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

    If mPortChanged Then
        If Not uStartListener Then
    
        End If
    
    End If

End Sub

Private Sub KPrefsPanel_Ready()

    mPortChanged = False

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
'Dim pc As BControl
'Dim pm As CTempMsg
'Dim sz As String
'
'    On Error Resume Next
'
'    With mPage
'        .SetMargin 0
'
'        .Add new_BPrefsControl("fancytoggle2", "ft>WatchSpecificApps", "Watch specific applications?", , IIf(mSettings.WatchSpecificApps, "1", "0"))
'
'        Set pm = New CTempMsg
'        pm.Add "item-height", 24&
'        Set pc = new_BPrefsControl("listbox", "lb>apps", sz, , , pm)
'        pc.SizeTo 0, 100
'        mPage.Add pc
'
'        .Add new_BPrefsControl("fancyplusminus", "fpm>apps")
'
''        Set pm = New CTempMsg
''        pm.Add "min", 0&
''        pm.Add "max", 60&
''        pm.Add "freq", 5&
''        pm.Add "label_size", 50&
''        .Add new_BPrefsControl("fancyslider", "fs>NotificationTimeout", "", "", CStr(0), pm)
'
'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("label", "", App.FileDescription & " V" & CStr(App.Major) & "." & CStr(App.Revision), , , , False)
'
'        If g_IsPressed(VK_LCONTROL) Then _
'            .Add new_BPrefsControl("fancytoggle2", "ft>DebugMode", "Debug Mode?", , IIf(mSettings.DebugMode, "1", "0"))
'
'    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)

    mConfig.Update Control.GetName, Value
    mConfig.Save

    Select Case Control.GetName

    Case "port"
        mPortChanged = True

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

'    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

'    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

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

Private Sub theSocket_OnConnectionRequest(ByVal requestID As Long)

    mCount = mCount + 1
    ReDim Preserve mConn(mCount)
    Set mConn(mCount) = New TConnection
    mConn(mCount).Accept requestID, mConfig

End Sub

Private Function uStartListener() As Boolean

    mPortIsBad = True

    If (mConfig Is Nothing) Or (theSocket Is Nothing) Then _
        Exit Function

Dim n As Long

    n = Val(mConfig.GetValueWithDefault("port", "80"))

    If n = 0 Then
        g_Debug "uStartListener(): invalid port 0 corrected to 80", LEMON_LEVEL_WARNING
        n = 80

    End If

    If theSocket.State = sckListening Then _
        theSocket.CloseSocket

    If theSocket.Bind(n) Then
        theSocket.Listen
        uStartListener = True
        mPortIsBad = False

    Else
        g_Debug "uStartListener(): couldn't bind socket", LEMON_LEVEL_CRITICAL

    End If

End Function
