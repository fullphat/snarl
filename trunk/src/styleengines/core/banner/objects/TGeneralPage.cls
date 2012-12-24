VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TGeneralPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()

    With mPage
        .Add new_BPrefsControl("fancybutton2", "go_lemon", "Launch Lemon")
        .Add new_BPrefsControl("label", "", "Lemon is an external log/debug console created by k23 productions.  Snarl sends information to it which can be useful for debugging purposes.")

        .Add new_BPrefsControl("separator", "")
        .Add new_BPrefsControl("fancybutton2", "open_config", "Open Config File")
        .Add new_BPrefsControl("label", "", "Opens the current config file in notedpad so it can be edited manually.")

        .Add new_BPrefsControl("fancybutton2", "cycle_config", "Reload Config File")
        .Add new_BPrefsControl("label", "", "Reloads the current configuration file.")

        .Add new_BPrefsControl("separator", "")
        .Add new_BPrefsControl("fancybutton2", "test", "Test Notification")
        .Add new_BPrefsControl("label", "", "Sends an special test message to the Snarl engine which should result in a notification appearing.  This message is sent using the same mechanism a 3rd party application would use and therefore should prove (or otherwise) that the Snarl notification engine is running correctly.")

'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("fancybutton2", "restart_style_roster", "Restart Style Roster")

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

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
