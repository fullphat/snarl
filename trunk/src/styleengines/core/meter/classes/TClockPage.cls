VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TClockPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()



End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)

End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
Dim pGap As BControl

    With Page
        .Add new_BPrefsControl("label", "", "Placeholder page.")

        Set pGap = new_BPrefsControl("label", "", " ")
        pGap.SizeTo -1, 80
        .Add pGap

    End With

End Sub

Private Sub KPrefsPage_Destroy()

End Sub

Private Sub KPrefsPage_Detached()

End Sub

Private Function KPrefsPage_hWnd() As Long

End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)

End Sub
