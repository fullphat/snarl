VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyleDefaultsPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Closed()

Dim mhWndOwner As Long
Dim mPanel As BPrefsPanel

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    If Control.GetName() = "toolbar" Then
        Select Case Val(Value)
        Case 1
            g_DoSchemePreview2 "Default", "", False, -1
            g_DoSchemePreview2 "Default", "", True, -1

        Case 3
            Debug.Print "(*)"

        End Select
    End If

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
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

    If mhWndOwner <> 0 Then _
        EnableWindow mhWndOwner, -1

    RaiseEvent Closed

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Function Go(ByVal hWndPanel As Long, Optional ByVal IsNC As Boolean) As Long
Dim pp As BPrefsPage
Dim pm As CTempMsg
Dim pc As BControl

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle IIf(IsNC, "Notification Centre", "Default Style Settings")
        .SetWidth 430
        .SetWindow 1

        If IsNC Then
            ' /* page */
            Set pp = new_BPrefsPage("", , Me, 480)
            With pp
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "height", 290
                Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
                BTabStrip_AddPage pc, "General", new_BPrefsPage("nc-general", , New TNetSubPage)
                .Add pc
            End With
            .AddPage pp
            .Go

            SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(mPanel.hWnd, GWL_STYLE) And (Not WS_MINIMIZEBOX)
            g_SetWindowIconToAppResourceIcon .hWnd

            If hWndPanel <> 0 Then
                mhWndOwner = g_GetTopLevel(hWndPanel)
                SetWindowLong .hWnd, GWL_HWNDPARENT, mhWndOwner
                EnableWindow mhWndOwner, 0
                g_MoveWindow .hWnd, , , , , True, mhWndOwner

            End If

            g_ShowWindow .hWnd, True, True

        Else
            ' /* page */
            Set pp = new_BPrefsPage("", , Me, 480)
            With pp
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "height", 240
                Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
                BTabStrip_AddPage pc, "General", new_BPrefsPage("defs-general", , New TNetSubPage)
                BTabStrip_AddPage pc, "Title", new_BPrefsPage("defs-title", , New TNetSubPage)
                BTabStrip_AddPage pc, "Text", new_BPrefsPage("defs-text", , New TNetSubPage)
'                BTabStrip_AddPage pc, "Advanced", new_BPrefsPage("defs-adv", , New TNetSubPage)
                .Add pc
                .Add new_BPrefsControl("fancytoolbar", "toolbar", "Test Settings||Restore Defaults")
            End With
            .AddPage pp
            .Go
            g_SetWindowIconToAppResourceIcon .hWnd
            Go = .hWnd

        End If

    End With

End Function


