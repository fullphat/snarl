VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LogSnoop Debug Log"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_ICON = 1

Dim theTrayIcon As BNotifyIcon
Dim mPanel As BPrefsPanel
Dim WithEvents Snarl As Snarl
Attribute Snarl.VB_VarHelpID = -1

'Private Const WAIT_INFINITE = -1&

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long

Dim mhWndAppHelper As Long
Dim mhWndSysHelper As Long
Dim mhWndSecHelper As Long

Implements KPrefsPage
Implements KPrefsPanel
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim pmi As OMMenuItem

    Select Case uMsg

    Case WM_USER
        ' /* tray icon */
        Select Case lParam
        Case WM_LBUTTONDBLCLK
            uDoPrefs

        Case WM_RBUTTONUP
            With New OMMenu
                .AddItem .CreateItem("prefs", "Preferences...")
                .AddSeparator
                .AddItem .CreateItem("quit", "Quit")

                Set pmi = .Track(hwnd)
                If Not (pmi Is Nothing) Then
                    Select Case pmi.Name
                    Case "prefs"
                        uDoPrefs

                    Case "quit"
                        Unload Me

                    End Select

                End If

            End With

        End Select


    Case WM_CLOSE
        Me.Add "WM_CLOSE"
        Unload Me


    Case WM_SYSCOMMAND
        If (LoWord(wParam) And &HFFF0&) = SC_MINIMIZE Then
            g_ShowWindow hwnd, False
            ReturnValue = 0
            BWndProcSink_WndProc = True

        End If


    Case snAppMsg

        Select Case wParam
        Case SNARLAPP_DO_PREFS
            uDoPrefs

        Case SNARLAPP_DO_ABOUT
            snDoRequest "notify?app-sig=" & App.ProductName & _
                        "&title=LogSnoop " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments & " (Build " & CStr(App.Revision) & ")" & _
                        "&text=" & App.LegalCopyright & vbCrLf & vbCrLf & "Notifies about applications comings and goings.  As far as LogSnoop is concerned, an application is anything that puts an entry in the Windows Taskbar." & _
                        "&icon=" & g_MakePath(App.Path) & IIf(g_IsIDE(), "", "bin\") & "icon.png"

        End Select


    End Select

End Function

Private Sub Form_Load()
Dim sz As String

    On Error Resume Next

    If App.PrevInstance Then
        ' /* we're already running... */
        sz = Me.Caption
        Me.Caption = ""

        ' /* if -quit specified, tell the other instance to quit */
        If InStr(Command, "-quit") <> 0 Then _
            PostMessage FindWindow("ThunderRT6FormDC", sz), WM_CLOSE, 0, ByVal 0&

        ' /* unload either way */
        Unload Me
        Exit Sub

    End If

    Me.Add "starting..."
    Me.Add "  " & App.Title & " " & App.Major & "." & App.Minor & " Build " & App.Revision

    Set Snarl = get_snarl()
    window_subclass Me.hwnd, Me

    Set theTrayIcon = New BNotifyIcon
    With theTrayIcon
        .SetTo Me.hwnd, WM_USER
        If g_IsIDE() Then
            .Add "icon", Me.Icon.Handle, "LogSnoop"

        Else
            .Add "icon", LoadImage(App.hInstance, 1&, IMAGE_ICON, 16, 16, 0), "LogSnoop"

        End If

    End With

    If Not g_IsIDE() Then _
        Me.Hide

    ' /* register with Snarl, if it's around */

    If is_snarl_running() Then _
        Snarl_SnarlLaunched

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Snarl_SnarlQuit

    window_subclass Me.hwnd, Nothing

    If Not (mPanel Is Nothing) Then _
        mPanel.Quit

    If Not (theTrayIcon Is Nothing) Then
        theTrayIcon.Remove "icon"
        Set theTrayIcon = Nothing

    End If

    snarl_unregister App.ProductName

End Sub

Private Sub Snarl_SnarlLaunched()

    Me.Add "[snarl launched]"

    ' /* register */

Dim hr As Long

    hr = snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", , Me.hwnd, , SNARLAPP_HAS_ABOUT Or SNARLAPP_HAS_PREFS Or SNARLAPP_IS_WINDOWLESS)
    If hr > 0 Then
        Add "snarl token: " & CStr(hr)

'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & APP_APPEARED & "&name=Application appeared"
'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & APP_DISAPPEARED & "&name=Application disappeared"

    Else
        Add "couldn't register with Snarl (" & CStr(Abs(hr)) & ")"

    End If

    ' /* launch the helpers - one for each log file */

    uRunHelper "Application", mhWndAppHelper
    uRunHelper "System", mhWndSysHelper
'    uRunHelper "Security", mhWndSecHelper

End Sub

Private Sub Snarl_SnarlQuit()

    Me.Add "[snarl quit]"

    ' /* quit the various helpers */

    If mhWndAppHelper Then _
        PostMessage mhWndAppHelper, WM_CLOSE, 0, ByVal 0&

    If mhWndSysHelper Then _
        PostMessage mhWndSysHelper, WM_CLOSE, 0, ByVal 0&

    If mhWndSecHelper Then _
        PostMessage mhWndSecHelper, WM_CLOSE, 0, ByVal 0&

End Sub

Public Sub Add(ByVal Text As String)

    List1.AddItem Text
    List1.ListIndex = List1.ListCount - 1
    g_Debug Text

End Sub

Private Sub uDoPrefs()
Dim pPage As BPrefsPage
Dim pCtl As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "LogSnoop Preferences"
            .SetWidth 400

            Set pPage = new_BPrefsPage("", Nothing, Me)

            With pPage
                .SetMargin 32
'                .Add new_BPrefsControl("banner", "", "Folder Watch List")
'
'                Set pm = New CTempMsg
'                pm.Add "item-height", 38
''                pm.Add "checkboxes", 1&
'                Set pCtl = new_BPrefsControl("listbox", "watch_list", , , "1", pm)
'                pCtl.SizeTo 0, 160
'                .Add pCtl
'
'                Set pCtl = new_BPrefsControl("fancyplusminus", "add_remove", "")
'                .Add pCtl
'
''                .Add new_BPrefsSeparator
'
'
''                .Add new_BPrefsControl("fancytoolbar", "feed_toolbar", "Show Headline|Show Summary|Refresh|Feed Information", , , , False)
''
''                .Add new_BPrefsControl("fancytoggle2", "UseDefaultCallback", "Clicking the notification opens the item?", , IIf(gConfig.UseDefaultCallback, "1", "0"))
'
''                .Add new_BPrefsControl("banner", "", "Status Changes")
''                .Add new_BPrefsControl("label", "", "Include changes from the following groups:")
'''                .Add new_BPrefsControl("fancytoggle2", "UseDefaultCallback", "Clicking the notification opens the item?", , IIf(gConfig.UseDefaultCallback, "1", "0"))
'
''                .Add new_BPrefsControl("label", "", "LogSnoop will alert you to incoming IM conversations and phone calls, as well as contact status changes.  Due to limitations of the Communicator API, only certain information is available.")
''                .Add new_BPrefsControl("label", "snarl_state", IIf(mHasSnarl, "Snarl is running", "Snarl is not running"))
''
''                .Add new_BPrefsControl("banner", "", "Options")
'
                .Add new_BPrefsControl("banner", "", "Debug")
                .Add new_BPrefsControl("fancybutton2", "ShowHideDebug", "Show/Hide Debug Window")

                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "quit_app", "Quit LogSnoop")
                .Add new_BPrefsControl("label", "", "LogSnoop " & CStr(App.Major) & "." & CStr(App.Minor) & " (Build " & CStr(App.Revision) & ") " & App.LegalCopyright, , , , False)
'                .Add new_BPrefsControl("label", "", "http://www.fullphat.net", , , , False)

            End With

            .AddPage pPage
            .Go
            g_SetWindowIconToAppResourceIcon .hwnd
            SetForegroundWindow .hwnd

'            uUpdateFeedList

        End With
    End If

    g_ShowWindow mPanel.hwnd, True, True

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
'Dim pfw As TFolderWatch
'Dim i As Long
'
'    Select Case Control.GetName
'
'    Case "add_remove"
'        If Value = "+" Then
''            If g_IsPressed(VK_CONTROL) Then
''                theAddFeedPanel_AddFeed Clipboard.GetText()
''
''            Else
'                Set theAddWatchPanel = New TAddWatchPanel
'                theAddWatchPanel.Go mPanel.hWnd
'
''            End If
'
'        ElseIf (Value = "-") Then
'            i = Val(prefskit_GetValue(mPanel, "watch_list"))
'            Set pfw = mFolders.TagAt(i)
'            If (pfw Is Nothing) Then _
'                Exit Sub
'
'            mFolders.Remove i
'            uWriteConfig
'            uUpdateList
'
'            snDoRequest "remclass?app-sig=" & App.ProductName & "&id=" & pfw.Guid
'
'            prefskit_SetValue mPanel, "watch_list", CStr(i)
'
'        End If
'
''    Case "UseDefaultCallback"
''        gConfig.UseDefaultCallback = (Value = "1")
''        uUpdateConfig
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    If Control.GetName = "quit_app" Then
        Unload Me

    ElseIf Control.GetName = "ShowHideDebug" Then
        Me.Visible = Not Me.Visible

    End If

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

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()

    uUpdateList

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub uUpdateList()

    If (mPanel Is Nothing) Then _
        Exit Sub

'Dim pfw As TFolderWatch
'Dim sz As String
'
'    With mFolders
'        .Rewind
'        Do While .GetNextTag(pfw) = B_OK
'            sz = sz & g_FormattedMidStr(pfw.Path, 55) & "#?0#?" & pfw.FlagsAsString() & "|"
'
'        Loop
'
'    End With
'
'    sz = g_SafeLeftStr(sz, Len(sz) - 1)
'    prefskit_SafeSetText mPanel, "watch_list", sz

End Sub

Private Sub uWriteConfig()
'Dim pfw As TFolderWatch
'Dim pcf As CConfFile3
'Dim pcs As CConfSection

'    Set pcf = New CConfFile3
'    pcf.SetFile g_MakePath(App.Path) & "appsnoop.conf"
'
'    With mFolders
'        .Rewind
'        Do While .GetNextTag(pfw) = B_OK
'            Set pcs = New CConfSection
'            With pcs
'                .SetName "watch"
'                .Add "guid", pfw.Guid
'                .Add "path", pfw.Path
'                .Add "flags", pfw.Flags
'                .Add "recurse", pfw.RecurseAsString
'
'            End With
'
'            pcf.Add pcs
'
'        Loop
'
'    End With
'
'    pcf.Save

End Sub

Private Function uRunHelper(ByVal LogFile As String, ByRef hWndHelper As Long) As Boolean
Dim lPid As Long

    On Error Resume Next

    Err.Clear

    ' /* launch the helper */

    lPid = Shell(Chr$(34) & g_MakePath(App.Path) & "logsnoophelper.exe" & Chr$(34) & " " & LogFile)
    If Err.Number Then
        Debug.Print "uRunHelper(): launch failed: " & Err.Description
        Exit Function

    End If

    Debug.Print "uRunHelper(): pid=" & CStr(lPid)

Dim h As Long

    h = OpenProcess(SYNCHRONIZE, -1, lPid)
    WaitForInputIdle h, 1000
    CloseHandle h

    ghWnd = 0
    EnumWindows AddressOf EnumWindowsProc, lPid

    ' /* if the window was found, ghWnd will be non-zero */

    hWndHelper = ghWnd
    Debug.Print "helper window=" & g_HexStr(hWndHelper)
    uRunHelper = (hWndHelper <> 0)

End Function



