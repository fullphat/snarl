VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Snarl"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   2940
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Other icons: Gnome themes by various artists"
      Height          =   255
      Index           =   5
      Left            =   1980
      TabIndex        =   6
      Top             =   2340
      Width           =   4155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UI icons: PixeloPhilia by Ömer ÇETIN (aka ~omercetin) "
      Height          =   255
      Index           =   4
      Left            =   1980
      TabIndex        =   5
      Top             =   2040
      Width           =   4155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Snarl icon by Paul Davey (aka Mattahan)"
      Height          =   255
      Index           =   3
      Left            =   1980
      TabIndex        =   4
      Top             =   1740
      Width           =   4155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.getsnarl.info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008A504A&
      Height          =   255
      Left            =   2370
      TabIndex        =   3
      Top             =   3060
      Width           =   1710
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   600
      Picture         =   "frmAbout.frx":1042
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   430
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   60
      Picture         =   "frmAbout.frx":1CA8
      Top             =   180
      Width           =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Snarl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1980
      TabIndex        =   2
      Top             =   120
      Width           =   4395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A notification system for Windows"
      Height          =   255
      Index           =   1
      Left            =   1980
      TabIndex        =   1
      Top             =   780
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2005-2012 full phat products"
      Height          =   255
      Index           =   2
      Left            =   1980
      TabIndex        =   0
      Top             =   1140
      Width           =   4155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   2880
      Left            =   0
      Top             =   0
      Width           =   6450
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           frmAbout.frm
    '/
    '/  Description:    Displays the product info and handles various other tasks
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As Any, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function EmptyWorkingSet Lib "psapi.dll" (ByVal hProcess As Long) As Long

Private Declare Function WTSRegisterSessionNotification Lib "Wtsapi32" (ByVal hWnd As Long, ByVal THISSESS As Long) As Long
Private Declare Function WTSUnRegisterSessionNotification Lib "Wtsapi32" (ByVal hWnd As Long) As Long

Private Const NOTIFY_FOR_ALL_SESSIONS As Long = 1

Private Const WM_WTSSESSION_CHANGE As Long = &H2B1
Private Const WTS_SESSION_LOCK As Long = 7
Private Const WTS_SESSION_UNLOCK As Long = 8

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETSCREENSAVERRUNNING = &H72


Dim mSysKeyPrefs As Long
Dim mSysKeyTest As Long

Dim mTrayIcon As BNotifyIcon

Dim m_About As String
Dim m_SelectedApp As String         ' // current selected application in listbox

Dim mPrefs As T_CONFIG
Dim mCurAlert As TAlert

Dim mPanel As BPrefsPanel
Dim mAppsPage As TAppsPage

    ' /* listening sockets */
Dim WithEvents GrowlUDPSocket As CSocket        ' // 9887 (UDP)
Attribute GrowlUDPSocket.VB_VarHelpID = -1
Dim mSNPListener As CSnarlListener              ' // 9887 (TCP) on 0.0.0.0
Dim mGNTPListener As CSnarlListener             ' // 23053 (TCP) on 0.0.0.0
Dim mJSONListener As CSnarlListener             ' // 9889 (TCP) on 0.0.0.0
Dim mMelonListener As CSnarlListener            ' // 5233 (TCP) on 0.0.0.0

Dim mClickThruOver As CSnarlWindow
Dim mMenuOpen As Boolean
Dim mDownloadId As Long
Dim WithEvents theReadyTimer As BTimer
Attribute theReadyTimer.VB_VarHelpID = -1

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long

End Type

Private Declare Function GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO) As Boolean

Dim WithEvents theIdleTimer As BTimer
Attribute theIdleTimer.VB_VarHelpID = -1

    ' /* R2.4.2 */
Dim mKeyCloseAll As Long
Dim mKeyClose As Long

Dim WithEvents theAppList As TAppsPopUpWindow
Attribute theAppList.VB_VarHelpID = -1

'Dim mMarkMissedOnClose As Boolean

    ' /* icon ids */
Private Const SN_II_NORMAL = 1&
Private Const SN_II_BUSY = 30&
Private Const SN_II_STOPPED = 40&
Private Const SN_II_MISSED = 50&
Private Const SN_II_AWAY = 60&

    ' /* R2.5.1 */
Dim WithEvents theGarbageTimer As BTimer
Attribute theGarbageTimer.VB_VarHelpID = -1
Private Const HISTORY_PAGE = 6

Dim mTestStyleAndScheme As String


Dim WithEvents theClassPanel As TConfigureClassPanel
Attribute theClassPanel.VB_VarHelpID = -1

Implements MMessageSink
Implements KPrefsPanel
Implements KPrefsPage
Implements MWndProcSink

Implements IDropTarget

Private Sub Form_Load()
Dim sz As String
Dim pm As OMMenu
Dim n As Integer

    On Error Resume Next

    g_HideFromView Me.hWnd

    ' /* R2.4 DR7: check for Calibri and default to Tahoma */

'    For n = Label3.LBound To Label3.UBound
'        Label3(n).Font.Name = "Calibri"
'        If Label3(n).Font.Name <> "Calibri" Then
'            Label3(n).Font.Name = "Tahoma"
'            Label3(n).Font.Size = Label3(n).Font.Size - 1
'
'        End If
'
'    Next n

    With Me.Font
        .Name = "Segoe UI"
        If .Name <> "Segoe UI" Then _
            .Name = "Tahoma"

    End With

    ' /* R2.4 DR8: register for TS session events */

    g_Debug "registering for terminal server session events..."
    WTSRegisterSessionNotification Me.hWnd, NOTIFY_FOR_ALL_SESSIONS

    ' /* register the hotkeys */

    g_Debug "setting hotkeys..."
    Me.bSetHotkeys

    If g_ConfigGet("use_notification_hotkey") = "1" Then _
        Me.bSetNotificationHotkey True

    ' /* pre-load our 'About' text */

    g_Debug "_load: pre-loading readme..."
    n = FreeFile()
    err.Clear
    Open g_MakePath(App.Path) & "read-me.rtf" For Input As #n
    If err.Number = 0 Then
        Do While Not EOF(n)
            Line Input #n, sz
            m_About = m_About & sz & vbCrLf

        Loop
        Close #n
    End If

    AddSubClass Me.hWnd, Me

    ' /* create the tray icon */

    Set mTrayIcon = New BNotifyIcon
    mTrayIcon.SetTo Me.hWnd, WM_SNARL_TRAY_ICON
    AddTrayIcon

    ' /* create our JSON listener */

    If g_ConfigGet("listen_for_json") = "1" Then _
        EnableJSON True

    ' /* create our Snarl listener */

    If g_ConfigGet("listen_for_snarl") = "1" Then _
        EnableSNP True

    ' /* set dynamic version info */

    Label3(0).Caption = "Snarl " & g_Version()
    g_Debug "_load: Version = " & g_Version(), LEMON_LEVEL_INFO

    ' /* create the idle input timer */

    Set theIdleTimer = new_BTimer(250)

    ' /* create the garbage collection timer */

    Set theGarbageTimer = new_BTimer(60000)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = True Then _
        Label1.Font.Underline = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Me.Hide
        Cancel = -1             ' // close gadget

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Long

    g_Debug "_Unload()", LEMON_LEVEL_PROC
    Me.Hide

    If NOTNULL(theClassPanel) Then _
        theClassPanel.Quit

    ' /* R2.4 DR8: unregister session events */

    WTSUnRegisterSessionNotification Me.hWnd

    ' /* stop the idle timer */

    Set theIdleTimer = Nothing

    ' /* close our Snarl listeners */

    If g_ConfigGet("listen_for_snarl") = "1" Then _
        EnableSNP False

    ' /* close our JSON listener */

    If g_ConfigGet("listen_for_json") = "1" Then _
        EnableJSON False

    ' /* quit the prefs panel, if it's open */

    If Not (mPanel Is Nothing) Then
        g_Debug "_Unload(): closing prefs window..."
        mPanel.Quit
        Set mPanel = Nothing
        g_Debug "_Unload(): prefs window closed"

    End If

    Set mTrayIcon = Nothing

    g_Debug "_Unload(): unsubclassing window..."
    RemoveSubClass Me.hWnd

    uUnregisterHotkeys
    Me.bSetNotificationHotkey False

End Sub

Private Sub GrowlUDPSocket_OnDataArrival(ByVal bytesTotal As Long)

    g_Debug "GrowlUDPSocket.OnDataArrival()", LEMON_LEVEL_PROC_ENTER
    g_Debug "received " & CStr(bytesTotal) & " byte(s)..."

Dim b() As Byte

    GrowlUDPSocket.GetData b(), vbArray + vbByte

    g_Debug "processing request..."
    g_ProcessGrowlUDP b(), bytesTotal, GrowlUDPSocket.RemoteHostIP

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub IDropTarget_DragEnter(ByVal pDataObject As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)
Dim pdo As CDropContentLite

    Set pdo = New CDropContentLite
    If pdo.SetTo(pDataObject) Then
        If pdo.HasFormat(CF_HDROP) Then _
            pdwEffect = DROPEFFECT_COPY

    End If

End Sub

Private Sub IDropTarget_DragLeave()
End Sub

Private Sub IDropTarget_DragOver(ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)
'Dim pdo As CDropContentLite
'
'    Set pdo = New CDropContentLite
'    If pdo.SetTo(pDataObject) Then
'        If pdo.HasFormat(CF_HDROP) Then _
'            pdwEffect = DROPEFFECT_COPY
'
'    End If

End Sub

Private Sub IDropTarget_Drop(ByVal pDataObject As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)
Dim pdo As CDropContentLite
Dim lpDrop As STGMEDIUM

    Set pdo = New CDropContentLite
    If pdo.SetTo(pDataObject) Then
        If pdo.GetData(CStr(CF_HDROP), lpDrop) Then
            uDoFileDrop lpDrop.Data
            pdo.Release lpDrop

        End If
    End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = True Then _
        Label1.Font.Underline = False

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = True Then _
        Label1.Font.Underline = False

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    If Control.GetName = "" Then _
        Exit Sub

    Select Case Control.GetName

    Case "ts>display", "advanced_tab_strip"
        ' /* don't write this to config! */
        Exit Sub

    ' /* [About] */

    Case "ftb>web_stuff"
        Select Case Val(Value)
        Case 1
            ' /* site */
            ShellExecute 0, "open", "http://www.getsnarl.info/", vbNullString, vbNullString, SW_SHOW

        Case 2
            ' /* forum */
            'ShellExecute 0, "open", "http://sourceforge.net/forum/?group_id=191100", vbNullString, vbNullString, SW_SHOW
            ShellExecute 0, "open", "http://groups.google.co.uk/group/snarl-discuss?hl=en", vbNullString, vbNullString, SW_SHOW

        Case 3
            ' /* blog */
            ShellExecute 0, "open", "http://www.snarl-development.blogspot.com/", vbNullString, vbNullString, SW_SHOW

        End Select
        Exit Sub

    End Select

    ' /* other controls - are there any now? */
    Debug.Print "frmAbout: setting '" & Control.GetName; "' to '" & Value & "'"
    g_ConfigSet Control.GetName, Value

    ' /* post-processing */
    Select Case Control.GetName()


    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim hWnd As Long
Dim sz As String
Dim ps As TStyle
Dim szText As String
Dim szIcon As String
Dim i As Long

    Select Case Control.GetName()

    Case "update_now"
        g_DoManualUpdateCheck

    Case "cycle_config"
        g_ConfigInit

    Case "test_display_settings"

        If g_IsPressed(VK_SHIFT) Then
            If g_StyleRoster.Find(style_GetStyleName(LCase$(mTestStyleAndScheme)), ps) Then
                If Not ps.IsRedirect Then
                    For i = 1 To ps.CountSchemes
                        ps.DoSchemePreview ps.SchemeAt(i), False, 50, False

                    Next i
                End If
            End If

        Else
            g_NotificationRoster.Hide 0, "_display_settings_test", App.ProductName, ""
            g_NotificationRoster.Hide 0, "_display_settings_test_priority", App.ProductName, ""

            If g_StyleRoster.Find(style_GetStyleName(LCase$(mTestStyleAndScheme)), ps) Then
                ' /* store the current style */
                sz = g_ConfigGet("default_style")
                ' /* temporarily switch to the selected style */
                g_ConfigSet "default_style", mTestStyleAndScheme
                szIcon = IIf(ps.IconPath = "", g_MakePath(App.Path) & "etc\icons\style.png", ps.IconPath)
                szText = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
                ' /* test normal */
                g_PrivateNotify , "Settings Test", szText, 0, szIcon, , "!test", , SN_NF_REMOTE Or SN_NF_SECURE, True, "_display_settings_test", 50, True

                ' /* test priority */
                If (ps.Flags And S_STYLE_SINGLE_INSTANCE) = 0 Then _
                    g_PrivateNotify , "Settings Test (Priority)", szText, 0, szIcon, 1, "!test", , SN_NF_REMOTE Or SN_NF_SECURE, True, "_display_settings_test_priority", 50, True
                
                ' /* restore the current style */
                g_ConfigSet "default_style", sz

            End If
        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static fIgnoreNext As Boolean
Dim dw As Long

    Select Case uMsg

    Case WM_COPYDATA, WM_MANAGE_SNARL, WM_SNARLTEST
        ' /* backwards compatability - send these directly to our handling window */
        ReturnValue = SendMessage(ghWndMain, uMsg, wParam, ByVal lParam)
        MWndProcSink_WndProc = True

    Case WM_HOTKEY
        Select Case LoWord(wParam)
        Case mSysKeyPrefs
            Me.NewDoPrefs

        Case mSysKeyTest
            bDoSysInfoNotification

        Case mKeyClose
            g_NotificationRoster.CloseMostRecent

        Case mKeyCloseAll
            g_NotificationRoster.CloseMultiple 0

        Case Else
            g_Debug "ISubClassed.WndProc(): Spurious WM_HOTKEY received: " & _
                    g_HexStr(HiWord(wParam), 4) & " " & g_HexStr(LoWord(wParam), 4), LEMON_LEVEL_WARNING

        End Select

        MWndProcSink_WndProc = True

    Case WM_SNARL_COMMAND
        ' /* this message shouldn't arrive here anymore, being directed to TMainWindow instead */
        Me.NewDoPrefs

    Case MSG_QUIT, WM_CLOSE
        PostQuitMessage 0

    Case WM_SNARL_TRAY_ICON
        Select Case lParam
        Case WM_RBUTTONUP
            If Not fIgnoreNext Then
                uDoMainMenu

            Else
                fIgnoreNext = False

            End If

        Case WM_LBUTTONDBLCLK
'            If g_NotificationRoster.HaveMissedNotifications Then
'                Me.bShowMissedPanel
'
'            Else
                Me.NewDoPrefs
'
'            End If

        End Select

    Case WM_ENTERMENULOOP
        mMenuOpen = True

    Case WM_EXITMENULOOP
        mMenuOpen = False

    Case WM_CLOSE
        Unload Me
        MWndProcSink_WndProc = True

    Case WM_INSTALL_SNARL
        If LoWord(wParam) = SNARL_CALLBACK_INVOKED Then _
            ShellExecute hWnd, "open", g_MakePath(App.Path) & gUpdateFilename, vbNullString, vbNullString, SW_SHOW

    Case RegisterWindowMessage("TaskbarCreated")
        ' /* R2.4 DR8 */
        AddTrayIcon

    Case WM_WTSSESSION_CHANGE
        Select Case wParam
        Case WTS_SESSION_LOCK
            Debug.Print "WM_WTSSESSION_CHANGE: =locked= " & Now()
            If g_ConfigGet("away_when_locked") = "1" Then _
                g_SetPresence SN_PF_AWAY_COMPUTER_LOCKED

        Case WTS_SESSION_UNLOCK
            Debug.Print "WM_WTSSESSION_CHANGE: =unlocked= " & Now()
            If g_ConfigGet("away_when_locked") = "1" Then _
                g_ClearPresence SN_PF_AWAY_COMPUTER_LOCKED

        End Select

    End Select

End Function

Private Sub uDoMainMenu()

    ' /* R2.31: only if admin says so */

    If gSysAdmin.InhibitMenu Then
        g_Debug "frmAbout.uDoMainMenu(): blocked by admin", LEMON_LEVEL_WARNING
        Exit Sub

    End If

    ' /* track the menu */

    SetForegroundWindow Me.hWnd

Dim pi As OMMenuItem

    With New OMMenu
        If gDebugMode Then
            .AddItem .CreateItem("sos", "SOS...")
            .AddSeparator

        End If

        .AddItem .CreateItem("about", "About Snarl...")
        .AddSeparator

        .AddItem .CreateItem("nc", "Notification Centre")
        .AddSeparator

        .AddItem .CreateItem("hide_all", "Hide All Notifications")
        .AddItem .CreateItem("sticky", "Sticky Notifications", , , (g_ConfigGet("sticky_snarls") = "1"))
        .AddSeparator

        .AddItem .CreateItem("dnd", "Do Not Disturb", , , g_IsPresence(SN_PF_DND_USER))
        .AddSeparator
        .AddItem .CreateItem("restart", "Restart Snarl", , g_IsRunning)

        If g_IsRunning Then
            .AddItem .CreateItem("stop", "Stop Snarl")

        Else
            .AddItem .CreateItem("start", "Start Snarl")

        End If

        .AddSeparator
        .AddItem .CreateItem("prefs", "Settings...", , Not gSysAdmin.InhibitPrefs)
        .AddItem .CreateItem("missed", "Missed Notifications", , (g_NotificationRoster.RealMissedCount > 0), , , , uMissedNotificationsSubmenu())
        .AddItem .CreateItem("app_list", "Snarl Apps...", , (g_AppRoster.CountSnarlApps > 0))
'        .AddItem .CreateItem("", "Snarl Apps", , , , , , g_AppRoster.SnarlAppsMenu())
        .AddSeparator
        .AddItem .CreateItem("quit", "Quit Snarl", , Not gSysAdmin.InhibitQuit)

        Set pi = .Track(Me.hWnd)

    End With

    PostMessage Me.hWnd, WM_NULL, 0, ByVal 0&

Dim pa As TApp

    If Not (pi Is Nothing) Then
        Select Case pi.Name
        Case "quit"
            PostQuitMessage 0
            Exit Sub

        Case "about"
            frmAbout.Show

        Case "restart"
            g_SetRunning False
            DoEvents
            Sleep 1500
            DoEvents
            g_SetRunning True

        Case "start"
            g_SetRunning True

        Case "stop"
            g_SetRunning False

        Case "prefs"
            Me.NewDoPrefs

        Case "nc"
            With g_NotificationRoster.NC
                If .IsVisible Then
                    .Hide
                
                Else
                    .Show
                
                End If
            End With


'        Case "app_mgr"
'            ShellExecute 0, "open", g_MakePath(App.Path) & "SNARLAPP_Manager.exe", vbNullString, vbNullString, SW_SHOW

        Case "sticky"
            g_ConfigSet "sticky_snarls", IIf(g_ConfigGet("sticky_snarls") = "1", "0", "1")

        Case "dnd"
            If g_IsPresence(SN_PF_DND_USER) Then
                ' /* clear it */
                g_ClearPresence SN_PF_DND_USER

            Else
                ' /* set it */
                g_SetPresence SN_PF_DND_USER
'                g_NotificationRoster.ResetMissedCount

            End If

        Case "missed"
            Me.bShowMissedPanel

        Case "hide_all"
            ' /* R2.4.2 */
            If Not (g_NotificationRoster Is Nothing) Then _
                g_NotificationRoster.CloseMultiple 0

        Case "sos"
            ' /* R2.4.2 DR3 */
            SOS_invoke New TSOSHandler

Dim i As Long

        Case "app_list"
            If (theAppList Is Nothing) Then
                Set theAppList = New TAppsPopUpWindow
                theAppList.Create 26
                
                With g_AppRoster
                    If .CountApps Then
                        For i = 1 To .CountApps
                            If .AppAt(i).IncludeInMenu Then _
                                theAppList.AddItem .AppAt(i).Name, .AppAt(i).Signature, .AppAt(i).CachedIcon

                        Next i
                    End If
                End With

                theAppList.Show

            End If

        Case Else

            ' /* missed list */

Dim pn As TNotification

            If g_BeginsWith(pi.Name, "!missed") Then
                Set pn = g_NotificationRoster.MissedList.TagAt(Val(g_SafeRightStr(pi.Name, Len(pi.Name) - 7)))
                If NOTNULL(pn) Then _
                    g_NotificationRoster.Add pn.Info, Nothing, False, True

            End If

'            If g_SafeLeftStr(pi.Name, 1) = "!" Then
'                Set pa = g_AppRoster.AppAt(Val(g_SafeRightStr(pi.Name, Len(pi.Name) - 1)))
'                pa.Activated
'
'            End If

'            sz = g_SafeLeftStr(pi.Name, 3)
'            szData = g_SafeRightStr(pi.Name, Len(pi.Name) - 3)
'
'            Select Case sz
'            Case "cfg"
'                ' /* Snarl App -> Settings... szData is App Roster index */
'                g_AppRoster.SnarlAppDo Val(szData), SNARLAPP_SHOW_PREFS
'
'            Case "abt"
'                ' /* Snarl App -> About... szData is App Roster index */
'                g_AppRoster.SnarlAppDo Val(szData), SNARLAPP_SHOW_ABOUT
'
'            End Select

        End Select
    End If

'    If update_config Then _
        g_WriteConfig

End Sub

Public Sub NewDoPrefs(Optional ByVal PageToSelect As Integer)

    ' /* R2.31: only if admin says we can... */

    If gSysAdmin.InhibitPrefs Then
        g_Debug "frmAbout.NewDoPrefs(): access blocked by admin", LEMON_LEVEL_WARNING
        MsgBox "Access to Snarl's preferences has been blocked by your system administrator.", vbInformation Or vbOKOnly, App.Title
        Exit Sub

    End If

Dim pp As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then

'        mCurrentStyleAndScheme = g_ConfigGet("default_style")

        g_Debug "frmAbout.NewDoPrefs(): creating panel..."

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            mPanel.SetTitle "Snarl Preferences"
            mPanel.SetWidth 540

            ' /* apps */
            g_Debug "frmAbout.NewDoPrefs(): apps page..."
            Set mAppsPage = New TAppsPage
            .AddPage new_BPrefsPage("Applications", load_image_obj(g_MakePath(App.Path) & "etc\icons\apps.png"), mAppsPage)

            ' /* notifications page */
            Set pp = new_BPrefsPage("Notifications", load_image_obj(g_MakePath(App.Path) & "etc\icons\notifications.png"), Me)
            With pp
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "height", 380
                Set pc = new_BPrefsControl("tabstrip", "ts>display", , , , pm)
                BTabStrip_AddPage pc, "Appearance", new_BPrefsPage("appearance1", , New TDisplaySubPage)
                BTabStrip_AddPage pc, "Behaviour", new_BPrefsPage("behaviour", , New TDisplaySubPage)
                BTabStrip_AddPage pc, "Layout", new_BPrefsPage("appearance2", , New TDisplaySubPage)
'                BTabStrip_AddPage pc, "Layout", new_BPrefsPage("layout", , New TDisplaySubPage)
                BTabStrip_AddPage pc, "Sounds", new_BPrefsPage("sounds", , New TDisplaySubPage)
'                BTabStrip_AddPage pc, "Redirection", new_BPrefsPage("redirection", , New TDisplaySubPage)
                BTabStrip_AddPage pc, "Advanced", new_BPrefsPage("advanced", , New TDisplaySubPage)

                .Add pc
                .Add new_BPrefsControl("fancybutton2", "test_display_settings", "Test Settings")

            End With
            .AddPage pp
            
            ' /* network */
            Set pp = new_BPrefsPage("Gateway", load_image_obj(g_MakePath(App.Path) & "etc\icons\gateway.png"), Me)
            With pp
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "height", 412
                Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
'                BTabStrip_AddPage pc, "General", new_BPrefsPage("net-general", , New TNetSubPage)
'                BTabStrip_AddPage pc, "Redirection", new_BPrefsPage("redirection", , New TDisplaySubPage)
                BTabStrip_AddPage pc, "Forwarding", new_BPrefsPage("net-clients", , New TNetSubPage)
                BTabStrip_AddPage pc, "Subscriptions", new_BPrefsPage("net-subs", , New TNetSubPage)
'                BTabStrip_AddPage pc, "Subscribers", new_BPrefsPage("net-subscribers", , New TNetSubPage)
''                BTabStrip_AddPage pc, "Listeners", new_BPrefsPage("net-listeners", , New TNetSubPage)
                .Add pc

            End With
            .AddPage pp

            ' /* presence */

'            g_Debug "frmAbout.NewDoPrefs(): presence page..."
'            Set pp = new_BPrefsPage("Presence", load_image_obj(g_MakePath(App.Path) & "etc\icons\presence.png"), Me)
'            With pp
'                .SetMargin 0
'                Set pm = New CTempMsg
'                pm.Add "height", 412
'                Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
'                BTabStrip_AddPage pc, "Active", new_BPrefsPage("pre-active", , New TNetSubPage)
'                BTabStrip_AddPage pc, "Away", new_BPrefsPage("pre-away", , New TNetSubPage)
'                BTabStrip_AddPage pc, "Busy", new_BPrefsPage("pre-busy", , New TNetSubPage)
'                .Add pc
'
'            End With
'            .AddPage pp

            ' /* addons */
            Set pp = new_BPrefsPage("AddOns", load_image_obj(g_MakePath(App.Path) & "etc\icons\extensions.png"), Me)
            With pp
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "height", 412
                Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
                BTabStrip_AddPage pc, "Displays", new_BPrefsPage("sty-display", , New TNetSubPage)
                BTabStrip_AddPage pc, "Redirects", new_BPrefsPage("sty-redirect", , New TNetSubPage)
                BTabStrip_AddPage pc, "Extensions", new_BPrefsPage("sty-extensions", , New TExtPage)
                BTabStrip_AddPage pc, "Style Engines", new_BPrefsPage("sty-engines", , New TNetSubPage)
                .Add pc

            End With
            .AddPage pp

            ' /* settings page */
            g_Debug "frmAbout.NewDoPrefs(): general page..."
            Set pp = new_BPrefsPage("Options", load_image_obj(g_MakePath(App.Path) & "etc\icons\general.png"), Me)
            With pp
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "height", 412
                Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
                BTabStrip_AddPage pc, "General", new_BPrefsPage("gen-basic", , New TNetSubPage)
                BTabStrip_AddPage pc, "Presence", new_BPrefsPage("gen-presence", , New TNetSubPage)
                BTabStrip_AddPage pc, "Network", new_BPrefsPage("net-general", , New TNetSubPage)
                BTabStrip_AddPage pc, "Security", new_BPrefsPage("gen-security", , New TNetSubPage)
                BTabStrip_AddPage pc, "Advanced", new_BPrefsPage("gen-advanced", , New TNetSubPage)

                If gDebugMode Then _
                    BTabStrip_AddPage pc, "Debug", new_BPrefsPage("gen-debug", , New TNetSubPage)

                .Add pc

            End With
            .AddPage pp

            ' /* R2.4.2 DR3: history */
            g_Debug "frmAbout.NewDoPrefs(): history page..."
            Set pp = new_BPrefsPage("History", load_image_obj(g_MakePath(App.Path) & "etc\icons\history.png"), Me)
            With pp
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "height", 412
                Set pc = new_BPrefsControl("tabstrip", "history_tabs", , , , pm)
                BTabStrip_AddPage pc, "Displayed", new_BPrefsPage("his-all", , New TNetSubPage)
                BTabStrip_AddPage pc, "Missed", new_BPrefsPage("his-missed", , New TNetSubPage)
                .Add pc

            End With
            .AddPage pp

            ' /* About page */
            g_Debug "frmAbout.NewDoPrefs(): about page..."
            Set pp = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & "etc\icons\about.png"), Me)
            With pp
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\snarl.png"
                pm.Add "image-height", 32
                pm.Add "valign", "centre"
                .Add new_BPrefsControl("labelex", "", "Snarl " & App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ")", , , pm)

                Set pm = New CTempMsg
                pm.Add "file", g_MakePath(App.Path) & "read-me.rtf"
                Set pc = new_BPrefsControl("rtf", "rtf")
                pc.DoExCmd "load", pm
                pc.SizeTo 0, 260
                .Add pc

                .Add new_BPrefsControl("fancytoolbar", "ftb>web_stuff", "Snarl Website|Discussion Group|Blog")

                Set pm = New CTempMsg
                pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\open_source.jpg"
                pm.Add "image-height", 48
                pm.Add "valign", "centre"
                .Add new_BPrefsControl("labelex", "", " Released under the Simplified BSD Licence.", , , pm)

            End With
            .AddPage pp

            ' /* Debug page */

'            If gDebugMode Then
'
'                g_Debug "frmAbout.NewDoPrefs(): debug page..."
'                Set pp = new_BPrefsPage("Debug", load_image_obj(g_MakePath(App.Path) & "etc\icons\debug.png"), Me)
'
'                With pp
'                    .SetMargin 96
'                    .Add new_BPrefsControl("banner", "", "Debugging")
'                    .Add new_BPrefsControl("fancybutton2", "go_lemon", "Open debug log")
'            '        .Add new_BPrefsControl("label", "", "The log file can be useful for debugging purposes.")
'
'                    .Add new_BPrefsControl("fancybutton2", "go_garbage", "Garbage collection", , , , g_IsWinXPOrBetter())
'
'            '        .Add new_BPrefsControl("separator", "")
'                    .Add new_BPrefsControl("banner", "", "Configuration")
'                    .Add new_BPrefsControl("fancybutton2", "open_config", "Open config folder")
'                    .Add new_BPrefsControl("label", "", "Opens the current config folder in Explorer so the various configuration files can be edited manually.")
'
'            '        .Add new_BPrefsControl("fancybutton2", "cycle_config", "Reload Config File")
'            '        .Add new_BPrefsControl("label", "", "Reloads the current configuration file.")
'
'            '        .Add new_BPrefsControl("separator", "")
'                    .Add new_BPrefsControl("banner", "", "Diagnostics")
'                    .Add new_BPrefsControl("fancybutton2", "test", "Test notification")
'                    .Add new_BPrefsControl("label", "", "Sends a special test message to the Snarl engine which should result in a notification appearing.  This message is sent using the same mechanism a 3rd party application would use and therefore should prove (or otherwise) that the Snarl notification engine is running correctly.")
'
'            '        .Add new_BPrefsControl("separator", "")
'            '        .Add new_BPrefsControl("fancybutton2", "restart_style_roster", "Restart Style Roster")
'
'                End With
'
'                .AddPage pp
'
'            End If

            g_Debug "frmAbout.NewDoPrefs(): displaying..."
            .Go
            g_SetWindowIconToAppResourceIcon .hWnd
            g_ShowWindow .hWnd, True, True
'            SetForegroundWindow .hWnd

            g_WindowToFront .hWnd, True

            g_Debug "frmAbout.NewDoPrefs(): done"

        End With

    End If

    If (PageToSelect > 0) And (PageToSelect <= mPanel.CountPages) Then _
        mPanel.SetPage PageToSelect

End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    ' /* prefs panel has been closed */

'    If mMarkMissedOnClose Then
'        g_NotificationRoster.MarkMissed
'        mMarkMissedOnClose = False
'
'    End If

    RevokeDragDrop mPanel.hWnd

    Set mPanel = Nothing
    Set mAppsPage = Nothing

    g_WriteConfig

End Sub

Private Sub KPrefsPanel_Ready()
Dim pc As BControl

    ' /* panel is now ready and visible, so select the first item in the registered apps combo - this then
    '    cascades a changed event down which configures all the other controls on that page */

    If mPanel.Find("cb>apps", pc) Then _
        pc.SetValue "1"

    ' /* R2.4 DR8: set here so we pick up the custom label change */

    If mPanel.Find("idle_minutes", pc) Then _
        pc.SetValue g_ConfigGet("idle_minutes")

    ' /* find our current style and select it in the 'Display' sub page */

'Dim i As Long
'Dim px As TStyle
'Dim j As Long

'    Debug.Print gPrefs.default_style

'    If Not (g_StyleRoster Is Nothing) Then
'        i = g_StyleRoster.IndexOf(style_GetStyleName(g_ConfigGet("default_style")))
'        If i Then
'            Set px = g_StyleRoster.StyleAt(i)
'            j = px.SchemeIndex(style_GetSchemeName(g_ConfigGet("default_style")))
'
'            If j Then
'                ' /* R2.4 RC1: select default style and scheme in [Styles] page*/
'                prefskit_SetValue mPanel, "installed_styles", CStr(i)
'                prefskit_SetValue mPanel, "installed_schemes", CStr(j)
'
''            prefskit_SetValue mPanel, "default_style", CStr(i)
''            prefskit_SetValue mPanel, "default_scheme", CStr(j)
''
''            If mPanel.Find("default_style", pc) Then _
''                pc.SetValue CStr(i)
''
''            If mPanel.Find("default_scheme", pc) Then _
''                pc.SetValue CStr(j)
'
'            End If
'        End If
'    End If

'    If mPanel.Find("busy_style", pc) Then
'        ' /* set the icons *
'        g_StyleRoster.SetNonWindowStyleIcons2 pc
'        ' /* select the right item */
'
'    End If
'
'    If mPanel.Find("away_style", pc) Then
'        g_StyleRoster.SetNonWindowStyleIcons2 pc
'        ' /* set the icons *
'        g_StyleRoster.SetNonWindowStyleIcons2 pc
'        ' /* select the right item */
'
'    End If

    bUpdateHistoryList
    bUpdateMissedList

'    If mPanel.Find("melontype_contrast", pc) Then _
'        pc.SetEnabled (gPrefs.font_smoothing = E_MELONTYPE)

    RegisterDragDrop mPanel.hWnd, Me

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub Label1_Click()

    ShellExecute Me.hWnd, "open", "http://www.getsnarl.info/", vbNullString, vbNullString, SW_SHOW

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = False Then _
        Label1.Font.Underline = True

End Sub

Private Property Get MMessageSink_Name() As String

    MMessageSink_Name = "p>snarl"

End Property

Private Function MMessageSink_Received(message As melon.MMessage) As Boolean
End Function

Friend Sub bUpdateExtList()

    If (mPanel Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    If (mPanel.Find("lb>extensions", pc)) Then _
        pc.Notify "update_list", Nothing

End Sub

Public Sub AddTrayIcon()
Dim hIcon As Long

    If (mTrayIcon Is Nothing) Or (g_ConfigGet("show_tray_icon") = "0") Or (gSysAdmin.HideIcon) Then _
        Exit Sub

    hIcon = LoadImage(App.hInstance, 1&, IMAGE_ICON, 16, 16, 0)
    If hIcon = 0 Then _
        hIcon = Me.Icon.Handle

    mTrayIcon.Remove "tray_icon"

    mTrayIcon.Add "tray_icon", hIcon, "Snarl"

End Sub

'Friend Sub bMissedNotificationsChanged()
'
'    On Error Resume Next
'
'    If (mTrayIcon Is Nothing) Then _
'        Exit Sub
'
'Dim hIcon As Long
'
'    If g_NotificationRoster.ActualMissedCount > 0 Then
'
'        hIcon = LoadResPicture(50, vbResIcon).Handle
'        If hIcon = 0 Then _
'            hIcon = Me.Icon.Handle
'
'        mTrayIcon.Update "tray_icon", hIcon, "Snarl - " & CStr(g_NotificationRoster.ActualMissedCount) & " missed notification" & IIf(g_NotificationRoster.ActualMissedCount = 1, "", "s")
'
'    Else
'        hIcon = LoadImage(App.hInstance, 1&, IMAGE_ICON, 16, 16, 0)
'        If hIcon = 0 Then _
'            hIcon = Me.Icon.Handle
'
'        mTrayIcon.Update "tray_icon", hIcon, "Snarl"
'
'    End If
'
'End Sub

Private Function uIsAlertEnabled(ByVal ConfigString As String) As Boolean
Dim sz() As String

    On Error Resume Next

    sz() = Split(ConfigString, "#?")
    uIsAlertEnabled = Val(sz(0))

End Function

Friend Function bSetHotkeys(Optional ByVal KeyCode As Long = 0) As Boolean

    ' /* return True if the prefs hotkey was registered ok */

    If g_ConfigGet("use_hotkey") = "0" Then
        ' /* hotkeys not enabled */
        uUnregisterHotkeys
        bSetHotkeys = True
        Exit Function

    End If

    If KeyCode = 0 Then
        g_Debug "bSetHotKeys(): registering existing hotkey (" & g_ConfigGet("hotkey_prefs") & ")", LEMON_LEVEL_INFO
        KeyCode = Val(g_ConfigGet("hotkey_prefs"))

    End If

Dim hSysKey As Long

    ' /* attempt to register the CTRL+keycode combo: if this fails, we fail */

    hSysKey = register_system_key(Me.hWnd, KeyCode, B_SYSTEM_KEY_CONTROL)
    g_Debug "bSetHotkeys(): register_system_key('prefs'): " & (hSysKey <> 0)
    If hSysKey = 0 Then _
        Exit Function

    ' /* registered okay, so unregister the existing hotkeys */

    uUnregisterHotkeys
    mSysKeyPrefs = hSysKey

    ' /* attempt to register the CTRL+SHIFT+keycode combo as well - don't fail if this fails though */

    hSysKey = register_system_key(Me.hWnd, KeyCode, B_SYSTEM_KEY_SHIFT Or B_SYSTEM_KEY_CONTROL)
    g_Debug "bSetHotkeys(): register_system_key('test'): " & (hSysKey <> 0)
    If hSysKey <> 0 Then _
        mSysKeyTest = hSysKey

    bSetHotkeys = True

End Function

Private Sub uUnregisterHotkeys()

    g_Debug "uUnregisterHotKeys(): unregister_system_key('prefs'): " & unregister_system_key(Me.hWnd, mSysKeyPrefs)
    g_Debug "uUnregisterHotKeys(): unregister_system_key('test'): " & unregister_system_key(Me.hWnd, mSysKeyTest)

    mSysKeyPrefs = 0
    mSysKeyTest = 0

End Sub

Friend Sub bUpdateAppList()

    If ISNULL(mPanel) Then _
        Exit Sub

Dim pc As BControl

    If mPanel.Find("cb>apps", pc) Then _
        pc.Notify "update_list", Nothing

End Sub

Friend Sub bUpdateClassList(ByVal AppToken As Long)

    If NOTNULL(theClassPanel) Then _
        theClassPanel.Refresh

End Sub

'Private Sub uUpdateStyleList()
'Dim pc As BControl
'
'    If Not (mPanel Is Nothing) Then
'        If mPanel.Find("installed_styles", pc) Then _
'            pc.Notify "update_list", Nothing
'
'    End If
'
'End Sub

'Friend Sub bUpdateRemoteComputerList()
'Dim pc As BControl
'
'    If Not (mPanel Is Nothing) Then
'        If mPanel.Find("lb>forward", pc) Then _
'            pc.Notify "update_list", Nothing
'
'    End If
'
'End Sub

Friend Sub bDoSysInfoNotification()
Static wreec As Long
Dim szMetric As String
Dim szMelon As String
Dim dFreq As Double
Dim hKey As Long
Dim dw As Long
Dim cb As Long

    ' /* empty working set */

    EmptyWorkingSet GetCurrentProcess()

    ' /* read melon version from registry */

    If RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\melon", hKey) = ERROR_SUCCESS Then
        If RegQueryValueEx(hKey, "DisplayVersion", 0, dw, ByVal 0&, cb) = ERROR_SUCCESS Then
            If dw = REG_SZ Then
                szMelon = String$(cb, 0)
                If RegQueryValueEx(hKey, "DisplayVersion", 0, ByVal 0&, ByVal szMelon, cb) = ERROR_SUCCESS Then _
                    szMelon = Left$(szMelon, cb - 1)

            End If
        End If
    End If

    If g_IsPressed(vbKeyTab) Then
        wreec = wreec + 1
        If wreec = 10 Then
            g_PrivateNotify , Decode64("V2UgbWFkZSBpdC4uLg==", False), Decode64("Q29yZTogQ2hyaXMsIFRvbWFzDQpBUEk6IENocmlzLCBUb2tlLCBTdmVuDQpPdGhlcjogSmV0b24gQWxpamksIEpvbnVzIENvbnJhZCwgSWNlYm9iLCBMdWlzIExhdmVuYSwgU2FtIExpc3RvcGFkIElJLCBTaGF3biBNY1RlYXIsIE1eMywgTWF4IE5vcnJpcywgUGFrbywgRGFuaWVsIFBlbmtpbiwgUHN5DQpUaGFua3MgdG86IEtlbiBCZXJyeSwgUGF1bCBEYXZleQ==", False), 0, , , , , SN_NF_SECURE
            Exit Sub

        End If

        wreec = Min(wreec, 675)

    End If

Dim pci As B_CPU_INFO

    get_cpu_info 1, pci
    dw = processor_count()

Dim pmi As T_MONITOR_INFO
Dim szAddr As String
Dim szScr As String

    With pci

        dFreq = .Speed
        If dFreq > 1000# Then
            dFreq = dFreq / 1000
            szMetric = "GHz"

        Else
            szMetric = "MHz"

        End If

'                        IIf(dw > 1, CStr(dw) & " x ", "") & Format$(dFreq, "0.0#") & " " & szMetric & " CPU" & vbCrLf & _

        szAddr = get_ip_address_table()
        szAddr = Replace$(szAddr, "0.0.0.0", "")
        szAddr = Replace$(szAddr, "127.0.0.1", "")
        szAddr = Replace$(szAddr, "  ", "")
        szAddr = Replace$(szAddr, " ", "; ")

        g_CountMonitors
        If g_GetPrimaryMonitorInfo(pmi) Then _
            szScr = "Screen: " & CStr(pmi.rcPhysical.Right - pmi.rcPhysical.Left) & "x" & CStr(pmi.rcPhysical.Bottom - pmi.rcPhysical.Top)

        g_PrivateNotify , g_GetUserName() & " on " & g_GetComputerName(), _
                        g_GetOSName() & " " & g_GetServicePackName() & vbCrLf & _
                        g_FileSizeToStringEx2(g_GetPhysMem(True), "GB", " ", "0.0") & " (" & g_FileSizeToStringEx2(g_GetPageMem(True) + g_GetPhysMem(True), "GB", " ", "0.0") & " total) RAM" & vbCrLf & _
                        szScr & vbCrLf & _
                        "IP: " & szAddr & vbCrLf & _
                        "Snarl " & App.Major & "." & App.Revision & " (" & App.Comments & ")" & vbCrLf & "melon " & IIf(szMelon <> "", szMelon, "??"), _
                        -1, _
                        g_MakePath(App.Path) & "etc\icons\snarl.png", , , , , , _
                        "_snarl_system_info"

    End With

End Sub

Private Sub theAppList_Closed()

    Set theAppList = Nothing

End Sub

Private Sub theAppList_Selected(ByVal Signature As String)
Dim pa As TApp

    If g_AppRoster.PrivateFindBySignature(Signature, pa) Then _
        pa.Activated

End Sub

Private Sub theGarbageTimer_Pulse()

    If (ISNULL(g_AppRoster)) Or (g_ConfigGet("garbage_collection") = "0") Then _
        Exit Sub

Dim t As Long

    t = GetTickCount()

Dim pa As TApp
Dim i As Long

    Debug.Print "running garbage collection..."

    With g_AppRoster
        If .CountApps Then
            For i = .CountApps To 1 Step -1
                Set pa = .AppAt(i)
                If Not pa.KeepAlive Then
                    ' /* app shouldn't remain registered if it disappears */
                    If pa.Pid > 0 Then
                        ' /* win32 */
                        If Not g_IsProcessRunning(pa.Pid) Then
                            g_Debug "GarbageCollection: '" & pa.Name & "' (" & CStr(pa.Pid) & ") has gone"
                            .Remove i

                        End If
                    End If
                End If
            Next i
        End If
    End With

    Debug.Print "...took " & GetTickCount() - t & " ms"

End Sub

Private Sub theIdleTimer_Pulse()
Static b As Boolean

    If g_ConfigGet("away_when_fullscreen") = "1" Then
        ' /* track foreground app state */
        b = uIsFullScreenMode()
        If b <> g_IsPresence(SN_PF_DND_FULLSCREEN_APP) Then
            ' /* full screen app state changed */
            If b Then
                g_SetPresence SN_PF_DND_FULLSCREEN_APP

            Else
                g_ClearPresence SN_PF_DND_FULLSCREEN_APP

            End If

            g_Debug "_theIdleTimer.Pulse(): " & Now() & " fullscreen app: " & g_IsPresence(SN_PF_DND_FULLSCREEN_APP)

        End If
    End If

Dim n As Long

    If g_ConfigGet("away_when_screensaver") = "1" Then
        ' /* track screensaver state */

        If SystemParametersInfo(SPI_GETSCREENSAVERRUNNING, 0, n, 0) <> 0 Then
            If b <> g_IsPresence(SN_PF_AWAY_SCREENSAVER_ACTIVE) Then
                ' /* screensaver state has changed */
                If b Then
                    g_SetPresence SN_PF_AWAY_SCREENSAVER_ACTIVE
        
                Else
                    g_ClearPresence SN_PF_AWAY_SCREENSAVER_ACTIVE
        
                End If
        
                g_Debug "_theIdleTimer.Pulse(): " & Now() & " screensaver: " & g_IsPresence(SN_PF_AWAY_SCREENSAVER_ACTIVE)

            End If
        End If
    End If

    ' /* ignore if no idle timeout set */

    n = g_SafeLong(g_ConfigGet("idle_minutes"))
    If n > 30 Then _
        n = 30              ' // bounds-check

    n = n * 60000           ' // convert to ms
    If n < 1 Then _
        Exit Sub


Dim lii As LASTINPUTINFO

    lii.cbSize = Len(lii)
    If GetLastInputInfo(lii) = False Then _
        Exit Sub

    lii.dwTime = GetTickCount() - lii.dwTime
'    Debug.Print "_theIdleTimer.Pulse(): idle time is now " & CStr(lii.dwTime) & " needs to be " & CStr(n)

    b = (lii.dwTime > n)
    If b <> g_IsPresence(SN_PF_AWAY_USER_IDLE) Then
        ' /* idle state has changed */
        If b Then
            g_SetPresence SN_PF_AWAY_USER_IDLE

        Else
            g_ClearPresence SN_PF_AWAY_USER_IDLE

        End If

        g_Debug "_theIdleTimer.Pulse(): " & Now() & " user idle: " & g_IsPresence(SN_PF_AWAY_USER_IDLE)

    End If

End Sub

Private Sub theReadyTimer_Pulse()

    ' /* tell everyone we're open for business */

    g_Debug "Notifying ready to run..."
    PostMessage HWND_BROADCAST, snSysMsg(), SNARL_BROADCAST_LAUNCHED, ByVal CLng(App.Major)

End Sub

Private Sub Timer1_Timer()
Dim pWindow As CSnarlWindow
Dim pt As POINTAPI
Dim i As Long

    If (g_NotificationRoster Is Nothing) Or (mMenuOpen) Then _
        Exit Sub

    GetCursorPos pt
    i = g_NotificationRoster.HitTest(pt.x, pt.y)

    If i > 0 Then
        Set pWindow = g_NotificationRoster.NotificationAt(i)

        ' /* existing? */

        If Not (mClickThruOver Is Nothing) Then
            If mClickThruOver.Id <> pWindow.Id Then
                ' /* different notification */
                mClickThruOver.MakeFuzzy False
                Set mClickThruOver = Nothing

            End If
        End If

'        Debug.Print pWindow.Window.hWnd & " " & pWindow.NotificationOnlyMode

        If pWindow.IsNonInteractive Then
            pWindow.MakeFuzzy True
            Set mClickThruOver = pWindow

        End If

    Else
        ' /* reset current */
        If Not (mClickThruOver Is Nothing) Then
            mClickThruOver.MakeFuzzy False
            Set mClickThruOver = Nothing

        End If

    End If

End Sub

'Friend Sub bUpdateStylesList()
'Dim pc As BControl
'
'    If Not (mPanel Is Nothing) Then
'        If mPanel.Find("installed_styles", pc) Then _
'            pc.Notify "update_list", Nothing
'
'    End If
'
'End Sub

Friend Sub bNotifyStyleEnginesChanged()

    If ISNULL(mPanel) Then _
        Exit Sub

Dim pc As BControl

    ' /* listbox in [AddOns]->[Style Engines] */
    If mPanel.Find("engine_list", pc) Then _
        pc.Notify "refresh", Nothing

    ' /* also update style lists */
    Me.bNotifyDisplaysChanged

End Sub

Public Sub EnableJSON(ByVal Enabled As Boolean)

    g_Debug "frmAbout.EnableJSON(" & CStr(Enabled) & ")", LEMON_LEVEL_PROC_ENTER

    If Enabled Then
        g_Debug "creating JSON listener..."
        Set mJSONListener = New CSnarlListener
        mJSONListener.Go JSON_DEFAULT_PORT

    Else
        g_Debug "stopping JSON listener..."
        mJSONListener.Quit
        Set mJSONListener = Nothing

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Sub EnableSNP(ByVal Enabled As Boolean)

    g_Debug "frmAbout.EnableSNP(" & CStr(Enabled) & ")", LEMON_LEVEL_PROC_ENTER

    If Enabled Then

        Set mSNPListener = New CSnarlListener
        mSNPListener.Go SNP_DEFAULT_PORT

        Set mGNTPListener = New CSnarlListener
        mGNTPListener.Go GNTP_DEFAULT_PORT

        Set mMelonListener = New CSnarlListener
        mMelonListener.Go MELON_DEFAULT_PORT

        ' /* R2.4: native Growl/UDP support */

        g_Debug "creating Growl UDP socket..."
        Set GrowlUDPSocket = New CSocket
        With GrowlUDPSocket
            .Protocol = sckUDPProtocol
            .Bind 9887

        End With

    Else

        g_Debug "closing Growl UDP socket..."

        If Not (GrowlUDPSocket Is Nothing) Then
            GrowlUDPSocket.CloseSocket
            Set GrowlUDPSocket = Nothing

        End If

        mSNPListener.Quit
        Set mSNPListener = Nothing

        mGNTPListener.Quit
        Set mGNTPListener = Nothing

        mMelonListener.Quit
        Set mMelonListener = Nothing

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function DoExtensionConfig(ByVal Index As Long) As Boolean

    If (mPanel Is Nothing) Then _
        Exit Function

    If IsWindowEnabled(mPanel.hWnd) = 0 Then _
        Exit Function

Dim pExtList As BControl

    If mPanel.Find("lb>extensions", pExtList) Then _
        pExtList.SetValue CStr(Index)

Dim pExt As TExtension

    Set pExt = g_ExtnRoster.ExtensionAt(Index)
    DoExtensionConfig = pExt.DoPrefs(mPanel.hWnd)

End Function

'Public Function DoStyleConfig(ByVal Index As Long) As Boolean
'
'    If (mPanel Is Nothing) Then _
'        Exit Function
'
'    If IsWindowEnabled(mPanel.hWnd) = 0 Then _
'        Exit Function
'
'Dim pStyleList As BControl
'
''    If mPanel.Find("installed_styles", pStyleList) Then _
''        pStyleList.SetValue CStr(Index)
'
'    If mPanel.Find("ftb>style", pStyleList) Then _
'        pStyleList.Changed "1"
'
'    DoStyleConfig = True
'
'End Function

'Public Sub DoAppConfig(ByVal AppName As String, Optional ByVal ClassName As String)
'
'    NewDoPrefs 1
'
'Dim i As Long
'
'    If Not (g_AppRoster Is Nothing) Then
'        i = g_AppRoster.IndexOf(AppName)
'        If i Then
'            ' /* select the application */
'            prefskit_SetValue mPanel, "cb>apps", CStr(i)
'
'            ' /* find the class */
'            i = g_AppRoster.AppAt(i).IndexOf(ClassName)
'            If i = 0 Then
'                ' /* not found/null - select _all */
'                prefskit_SetValue mPanel, "lb>classes", "1"
'
'            Else
'                ' /* select it */
'                prefskit_SetValue mPanel, "lb>classes", CStr(i)
'
'            End If
'
'        Else
'            g_Debug "frmAbout.DoAppConfig(): '" & AppName & "' not found", LEMON_LEVEL_CRITICAL
'
'        End If
'
'    Else
'        g_Debug "frmAbout.DoAppConfig(): app roster not available", LEMON_LEVEL_CRITICAL
'
'    End If
'
'End Sub

Public Sub DoAppConfigBySignature(ByVal Signature As String)

    On Error Resume Next

    If (g_AppRoster Is Nothing) Then _
        Exit Sub

Dim i As Long

    i = Val(g_AppRoster.IndexOfSig(Signature))
    If i = 0 Then
        g_Debug "frmAbout.DoAppConfigBySignature(): '" & Signature & "' not in app roster"
        Exit Sub

    End If

    ' /* show the apps page */

    NewDoPrefs 1

    ' /* select the app */

    prefskit_SetValue mPanel, "cb>apps", CStr(g_AppRoster.IndexOfSig(Signature))

    ' /* do a configure... */

Dim pc As BControl

    If mPanel.Find("ftb>app", pc) Then _
        mPanel.PageAt(1).ControlChanged pc, "1"

End Sub

Friend Sub bReadyToRun()

    Set theReadyTimer = new_BTimer(2000, True)

End Sub

Private Function uIsFullScreenMode() As Boolean
Static hWnd As Long
Static h As Long

    hWnd = uTopLevelFromPoint(1, 1)

'    g_Debug g_ClassName(hWnd) & " " & _
            g_ClassName(uParentFromPoint(g_ScreenWidth() - 1, 1)) & " " & _
            g_ClassName(uParentFromPoint(1, g_ScreenHeight() - 1)) & " " & _
            g_ClassName(uParentFromPoint(g_ScreenWidth() - 1, g_ScreenHeight() - 1))

    If hWnd = uTopLevelFromPoint(uScreenWidth() - 1, 1) Then
        If hWnd = uTopLevelFromPoint(uScreenWidth() - 1, uScreenHeight() - 1) Then
            If hWnd = uTopLevelFromPoint(1, uScreenHeight() - 1) Then
'                g_Debug "uIsFullScreenMode(): four points match: " & g_ClassName(hWnd) & " '" & g_WindowText(hWnd) & "'"
                h = GetWindow(hWnd, GW_HWNDPREV)
                Do While h
                    If uIsAppWindow(h) Then _
                        Exit Function
'                        g_PrivateNotify , "Fullscreen abandoned", g_ClassName(h) & " " & g_Quote(g_WindowText(h))

                    h = GetWindow(h, GW_HWNDPREV)

                Loop

                If g_IsAlphaBuild Then _
                    g_PrivateNotify , "Fullscreen detected", g_ClassName(hWnd) & " " & g_Quote(g_WindowText(hWnd))
'                g_Debug "uIsFullScreenMode(): no higher app window"

                ' /* filter out Windows7 Win+Tab class and other system gubbins */
                Select Case g_ClassName(hWnd)
                Case "Flip3D", "Progman"
                    Exit Function

                End Select

                uIsFullScreenMode = True

            End If
        End If
    End If

End Function

Private Function uGetTopLevel(ByVal hWnd As Long) As Long
Static h As Long

    uGetTopLevel = hWnd
    h = GetParent(uGetTopLevel)
    Do While h
        uGetTopLevel = h
        h = GetParent(uGetTopLevel)

    Loop

End Function

Private Function uTopLevelFromPoint(ByVal x As Long, ByVal y As Long) As Long
Static h As Long

    h = WindowFromPoint(x, y)
    If IsWindow(h) <> 0 Then _
        uTopLevelFromPoint = uGetTopLevel(h)

End Function

Private Function uScreenWidth(Optional ByVal VirtualScreen As Boolean = False) As Long

    uScreenWidth = GetSystemMetrics(SM_CXSCREEN)

End Function

Private Function uScreenHeight(Optional ByVal VirtualScreen As Boolean = False) As Long

    uScreenHeight = GetSystemMetrics(SM_CYSCREEN)

End Function

Private Function uIsAppWindow(ByVal hWnd As Long) As Boolean

    If hWnd = 0 Then _
        Exit Function

Static lExStyle As Long
Static Style As Long

    ' /* more reliable version (although it can pick up the 'wrong' window in cases
    '    where there's a choice (e.g. VB IDE and Platform SDK) - code taken from
    '    here: http://shell.franken.de/~sky/explorer-doc/taskbar_8cpp-source.html */

    ' /* modified 7-Sep-09 to also include dialog windows where the owner is an
    '    app window.  To filter these out, just exclude any hWnd where GW_OWNER is
    '    not zero */

    Style = GetWindowLong(hWnd, GWL_STYLE)
    lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)

    If ((Style And WS_VISIBLE) = 0) Or (IsIconic(hWnd)) Then _
        Exit Function

    If (lExStyle And WS_EX_APPWINDOW) Then
        uIsAppWindow = True
        Exit Function

    End If

    If (lExStyle And WS_EX_TOOLWINDOW) = 0 Then
        If (GetParent(hWnd) = 0) And (GetWindow(hWnd, GW_OWNER) = 0) Then
            uIsAppWindow = True

        Else
            uIsAppWindow = uIsAppWindow(GetWindow(hWnd, GW_OWNER))

        End If
    End If

End Function

Friend Sub bSetNotificationHotkey(ByVal Register As Boolean)

    g_Debug "frmAbout.bSetNotificationHotkey()", LEMON_LEVEL_PROC_ENTER

    If Register Then
        ' /* R2.4.2: registers Win+Esc and Win+Ctrl+Esc.  Win+Esc will close the most recent notification; Win+Ctrl+Esc closes all */
        g_Debug "registering hotkeys..."
        mKeyClose = register_system_key(Me.hWnd, vbKeyEscape, B_SYSTEM_KEY_WINDOWS)
        If mKeyClose = 0 Then _
            g_Debug "couldn't register Win+Esc system key", LEMON_LEVEL_WARNING

        mKeyCloseAll = register_system_key(Me.hWnd, vbKeyEscape, B_SYSTEM_KEY_WINDOWS Or B_SYSTEM_KEY_CONTROL)
        If mKeyCloseAll = 0 Then _
            g_Debug "couldn't register Win+Ctrl+Esc system key", LEMON_LEVEL_WARNING

    Else
        ' /* R2.4.2: unregisters Win+Esc and Win+Ctrl+Esc */

        g_Debug "unregistering hotkeys..."

        If mKeyClose Then _
            unregister_system_key Me.hWnd, mKeyClose

        If mKeyCloseAll Then _
            unregister_system_key Me.hWnd, mKeyCloseAll

        mKeyClose = 0
        mKeyCloseAll = 0

    End If

    g_Debug "done"

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Friend Sub bForwardersChanged()
Dim pc As BControl

    If Not (mPanel Is Nothing) Then
        If mPanel.Find("net_forward_list", pc) Then _
            pc.Notify "refresh", Nothing

    End If

End Sub

Friend Sub bSubsChanged()
Dim pc As BControl

    If Not (mPanel Is Nothing) Then
        If mPanel.Find("net_subscriber_list", pc) Then _
            pc.Notify "refresh", Nothing

    End If

End Sub

Friend Sub bUpdateHistoryList()

    If (mPanel Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    If mPanel.Find("history_list", pc) Then _
        uUpdateList g_NotificationRoster.History, pc

End Sub

Friend Sub bSetTrayIcon()

    On Error Resume Next

    If (mTrayIcon Is Nothing) Then _
        Exit Sub

Dim hIcon As Long
Dim sz As String

     ' /* tooltip */

    sz = "Snarl"
    If g_NotificationRoster.HaveMissedNotifications Then _
        sz = sz & " - " & CStr(g_NotificationRoster.RealMissedCount) & " missed notification" & IIf(g_NotificationRoster.RealMissedCount = 1, "", "s")

Dim n As Long

    n = SN_II_NORMAL

    If Not g_IsRunning Then
        ' /* takes precedence */
        n = SN_II_STOPPED
        sz = "Snarl (stopped)"

    ElseIf g_NotificationRoster.HaveMissedNotifications Then
        n = SN_II_MISSED

    ElseIf g_IsDND() Then
        n = SN_II_BUSY

    ElseIf g_IsAway() Then
        n = SN_II_AWAY

    End If

    g_Debug "bSetTrayIcon(): id=" & CStr(n)

    hIcon = LoadImage(App.hInstance, n, IMAGE_ICON, 16, 16, 0)
    If hIcon = 0 Then _
        hIcon = Me.Icon.Handle

    mTrayIcon.Update "tray_icon", hIcon, sz

End Sub

Private Function uMakeSafe(ByVal str As String, ByVal maxLen As Integer) As String

    str = Replace$(str, "|", ":")
    str = Replace$(str, vbCrLf, " ")
    uMakeSafe = g_FormattedMidStr(str, maxLen)

End Function

Public Function PanelhWnd() As Long

    If Not (mPanel Is Nothing) Then _
        PanelhWnd = mPanel.hWnd

End Function

Friend Sub bUpdateMissedList()

    If (mPanel Is Nothing) Then _
        Exit Sub

Dim pItem As TNotification
Dim pc As BControl
Dim i As Long
Dim j As Long

    If mPanel.Find("missed_list", pc) Then
        uUpdateList g_NotificationRoster.MissedList, pc

        With g_NotificationRoster.MissedList
            For i = .CountItems To 1 Step -1
                prefskit_SetItem pc, i, "checked", 1&
                j = j + 1
                Set pItem = .TagAt(i)
                prefskit_SetItem pc, j, "greyscale", IIf(pItem.WasReplayed, 1&, 0&)

            Next i
        End With
    End If


'    If (mPanel Is Nothing) Or (g_NotificationRoster Is Nothing) Then
'        g_Debug "frmAbout.bUpdateMissedList(): something bad happened", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    End If
'
'Dim pc As BControl
'
'    If Not mPanel.Find("missed_list", pc) Then
'        g_Debug "frmAbout.bUpdateMissedList(): something bad happened", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    End If
'
'Dim pMissed As BTagList
'
'    Set pMissed = g_NotificationRoster.MissedList()
'
'Dim pItem As TNotification
'Dim iCurrent As Long
'Dim szt As String
'Dim sz As String
'Dim i As Long
'
'    ' /* store the current selected item */
'
'    iCurrent = Val(pc.GetValue())
'
'    With pMissed
'        ' /* build the content string */
'        .Rewind
'        Do While .GetNextTag(pItem) = B_OK
'
'            ' /* prefix title (if there is one) with the app name */
'            szt = pItem.Info.ClassObj.App.Name
'            If pItem.Info.Title <> "" Then _
'                szt = szt & ": " & uMakeSafe(pItem.Info.Title, 70)
'
'            szt = szt & " (" & g_When(pItem.Info.DateStamp) & ")" & "#?" & _
'                        CStr(pItem.Info.Token) & "#?" & _
'                        uMakeSafe(pItem.Info.Text, 80)
'
'            sz = sz & szt & "|"
'
'        Loop
'
'        pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)
'
'        ' /* set the icons */
'
'        With pMissed
'            For i = 1 To .CountItems
'                Set pItem = .TagAt(i)
'                pc.DoExCmd B_EXTENDED_COMMANDS.B_SET_ITEM, prefskit_CreateImageMessage(i, load_image_obj(g_TranslateIconPath(pItem.Info.IconPath, ""))) ',  pItem.WasSeen)
'
'            Next i
'
'        End With
'
'    End With
'
'    pc.SetValue CStr(iCurrent + 1)

End Sub

Friend Sub bShowMissedPanel()

    Me.NewDoPrefs HISTORY_PAGE
'    mMarkMissedOnClose = True

Dim pc As BControl
Dim rc As RECT

    If mPanel.Find("history_tabs", pc) Then
        ' /* there's a bug in the Prefs Kit tab control which
        '    prevents 'pc.SetValue "2"' from working, so we
        '    employ a little haxie to fake it */
        GetClientRect pc.hWnd, rc
        SendMessage pc.hWnd, WM_LBUTTONDOWN, 0, ByVal MAKELONG(Fix(rc.Right / 2) + 2, rc.Top + 1)
        Sleep 1
        SendMessage pc.hWnd, WM_LBUTTONUP, 0, ByVal MAKELONG(Fix(rc.Right / 2) + 2, rc.Top + 1)

    End If

End Sub

Private Sub uUpdateList(ByRef NotificationList As BTagList, ByRef ListControl As BControl)
Dim pn As TNotification
Dim iCurrent As Long
Dim szText As String
Dim sz As String
Dim i As Long
Dim j As Long

    With NotificationList
        If .CountItems Then

            iCurrent = Val(ListControl.GetValue())

            For i = .CountItems To 1 Step -1
                Set pn = .TagAt(i)

                ' /* prefix title (if there is one) with the app name */
                sz = uMakeSafe(pn.AppNameAndTitle & " (" & g_When(pn.Info.DateStamp) & ")", 70)

                szText = szText & sz & _
                                "#?" & CStr(pn.Info.Token) & _
                                "#?" & uMakeSafe(pn.Info.Text, 80) & "|"

            Next i

            ' /* set content */
            ListControl.SetText g_SafeLeftStr(szText, Len(szText) - 1)

            ' /* set icons */
            For i = .CountItems To 1 Step -1
                j = j + 1
                Set pn = .TagAt(i)
'                prefskit_SetItem ListControl, j, "image-file", g_TranslateIconPath(pn.Info.IconPath, "")
                prefskit_SetItemObject ListControl, j, "image-object", pn.Icon

            Next i

            ListControl.SetValue CStr(iCurrent + 1)

        Else
            ' /* empty list */
            ListControl.SetText ""

        End If

    End With

End Sub

Private Sub uFileDropped(ByVal Path As String, ByRef Text As String)

    Select Case g_GetExtension(Path, True)
    Case "webforward"
        If g_CopyToAppData(Path, "styles\webforward") Then _
            Text = Text & g_Quote(g_RemoveExtension(g_FilenameFromPath(Path))) & " webforwarder installed" & vbCrLf

    Case "rsz"
        g_InstallRSZ Path, True

    End Select

End Sub

Private Sub uDoFileDrop(ByVal hDrop As Long)
Dim c As String

    c = DragQueryFile(hDrop, &HFFFFFFFF, 0&, 0)
    If c = 0 Then _
        Exit Sub

Dim szText As String
Dim sz As String
Dim i As Long

    For i = 0 To c - 1
        sz = String$(2049, 0)
        DragQueryFile hDrop, i, sz, LenB(sz)
        uFileDropped g_TrimStr(sz), szText

    Next i

'    If szText <> "" Then
'        g_PrivateNotify , "Installation complete", szText, , "!system-yes"
'
'    Else
'        g_PrivateNotify , "Installation failed", "There was a problem installing the selected file" & IIf(c > 1, "s", ""), , "!system-no"
'
'    End If

    g_StyleRoster.Restart

End Sub

Private Function uMissedNotificationsSubmenu() As OMMenu
Dim pt As TNotification
Dim pm As OMMenu
Dim j As Integer

    Set pm = New OMMenu

    With g_NotificationRoster.MissedList
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            If Not pt.WasReplayed Then
                j = j + 1
                pm.AddItem pm.CreateItem("!missed" & CStr(pt.Info.Token), g_Quote(g_FormattedMidStr(Replace$(pt.Info.Text, vbCrLf, " "), 48)) & " (" & pt.Info.ClassObj.App.NameEx & ")")
                If j = 10 Then _
                    Exit Do

            End If
        Loop

    End With

    If j > 0 Then
        pm.AddSeparator
        pm.AddItem pm.CreateItem("missed", "All Missed Notifications...")

    End If

    Set uMissedNotificationsSubmenu = pm

End Function

Friend Sub bSetPrevewStyle(ByVal Style As String)

    mTestStyleAndScheme = Style

End Sub

Friend Sub bNotifyDisplaysChanged()
Dim pc As BControl

    If NOTNULL(mPanel) Then
        ' /* list in [AddOns]->[Displays] */
        If mPanel.Find("display_styles_list", pc) Then _
            pc.Notify "refresh", Nothing

        ' /* combo box in [Display]->[Appearance] */
        If mPanel.Find("default_style_list", pc) Then _
            pc.Notify "refresh", Nothing

    End If

End Sub

Private Sub theClassPanel_Done()

    Set theClassPanel = Nothing

End Sub

Public Sub ShowClassConfigPanel(ByVal hWndOwner As Long, ByRef App As TApp, ByVal Class As String)

    If NOTNULL(theClassPanel) Then _
        theClassPanel.Quit

    Set theClassPanel = New TConfigureClassPanel
    theClassPanel.Go hWndOwner, App
    theClassPanel.SelectClass Class

End Sub
