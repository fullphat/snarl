VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDebugPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

    With mPage
        .SetMargin 96
        .Add new_BPrefsControl("banner", "", "Debugging")
        .Add new_BPrefsControl("fancybutton2", "go_lemon", "Open debug log")
'        .Add new_BPrefsControl("label", "", "The log file can be useful for debugging purposes.")

        .Add new_BPrefsControl("fancybutton2", "go_garbage", "Garbage collection", , , , g_IsWinXPOrBetter())

'        .Add new_BPrefsControl("separator", "")
        .Add new_BPrefsControl("banner", "", "Configuration")
        .Add new_BPrefsControl("fancybutton2", "open_config", "Open config folder")
        .Add new_BPrefsControl("label", "", "Opens the current config folder in Explorer so the various configuration files can be edited manually.")

'        .Add new_BPrefsControl("fancybutton2", "cycle_config", "Reload Config File")
'        .Add new_BPrefsControl("label", "", "Reloads the current configuration file.")

'        .Add new_BPrefsControl("separator", "")
        .Add new_BPrefsControl("banner", "", "Diagnostics")
        .Add new_BPrefsControl("fancybutton2", "test", "Test notification")
        .Add new_BPrefsControl("label", "", "Sends a special test message to the Snarl engine which should result in a notification appearing.  This message is sent using the same mechanism a 3rd party application would use and therefore should prove (or otherwise) that the Snarl notification engine is running correctly.")

'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("fancybutton2", "restart_style_roster", "Restart Style Roster")

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"
End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)
'Dim sz As String
Dim hWnd As Long

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()

    Case "go_lemon"
        ShellExecute 0, "open", "notepad.exe", l3LogPath(), vbNullString, SW_SHOW

    Case "open_config"
        ShellExecute 0, "open", g_GetPath(g_SettingsPath()), vbNullString, vbNullString, SW_SHOW

    Case "cycle_config"
        g_ConfigInit

    Case "go_garbage"
        If g_IsWinXPOrBetter() Then _
            CoFreeUnusedLibrariesEx 0, 0

    Case "test"
        ' /* mimic how we would do it from an external app... */
        hWnd = FindWindow("w>Snarl", "Snarl")
        If IsWindow(hWnd) <> 0 Then _
            SendMessage hWnd, WM_SNARLTEST, 0, ByVal 0&

'    Case "restart_style_roster"
'        If Not (g_StyleRoster Is Nothing) Then
'            melonLibClose g_StyleRoster
'            MsgBox "Click OK when you're ready for the Style Roster to start up", vbInformation Or vbOKOnly, App.Title
'            melonLibOpen g_StyleRoster
'
'        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()

End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)

End Sub

Private Function KPrefsPage_hWnd() As Long

End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)

End Sub








