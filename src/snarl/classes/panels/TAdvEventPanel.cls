VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAdvEventPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mEvent As TAlert
Dim mhWndOwner As Long

Public Event Closed()
Public Event Saved()

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Select Case Control.GetName()
    Case "redact-when"
        mEvent.SetValue Control.GetName, uRedactionFromIndex(Val(Value))

    Case Else
        Debug.Print "TAdvEventPanel.ControlChanged(): setting " & g_Quote(Control.GetName) & " to " & g_Quote(Value)
        mEvent.SetValue Control.GetName, Value

    End Select

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

    EnableWindow mhWndOwner, -1
    g_ShowWindow mhWndOwner, True, True
    Set mEvent = Nothing
    RaiseEvent Closed

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long, ByRef aEvent As TAlert)
Dim pp As BPrefsPage
Dim b As Boolean

    mhWndOwner = hWndPanel
    Set mEvent = aEvent

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle "Advanced Event Settings"
        .SetWidth 420
        .SetWindow 1

        ' /* page */
        Set pp = new_BPrefsPage("Content", load_image_obj(g_MakePath(App.Path) & "etc\icons\icon-event-content.png"), Me)
        With pp
            .SetMargin 90
            .SetWidth 420           ' // bug in prefs kit makes the default page width 400 pixels
            .Add new_BPrefsControl("label", "", "The following can be used to override the content provided by the application.")
            .Add new_BPrefsControl("fancyfilepicker", "custom_icon", aEvent.GetValue("custom_icon"), "Icon:")
            .Add new_BPrefsControl("fancyfilepicker", "custom_sound", aEvent.GetValue("custom_sound"), "Sound:")
            .Add new_BPrefsControl("fancyedit", "custom_ack", aEvent.GetValue("custom_ack"), "Callback:")
            .Add new_BPrefsControl("fancyedit", "custom_actions", aEvent.GetValue("custom_actions"), "Actions:")
            .Add new_BPrefsControl("label", "", "Separate multiple actions with semicolons, for example " & g_Quote("open=!open;close=!close"))

        End With
        .AddPage pp

        ' /* page */
        Set pp = new_BPrefsPage("Presence", load_image_obj(g_MakePath(App.Path) & "etc\icons\icon-event-presence.png"), Me)
        With pp
            .SetMargin 90
            .SetWidth 420           ' // bug in prefs kit makes the default page width 400 pixels
            .Add new_BPrefsControl("banner", "", "When I'm At My Computer...")
            .Add new_BPrefsControl("fancycycle", "override-normal", "Do default|Log as missed|Display sticky|Discard|Display normally|Make high priority", , aEvent.GetValue("override-normal"))
            .Add new_BPrefsControl("banner", "", "When I'm Away...")
            .Add new_BPrefsControl("fancycycle", "override-away", "Do default|Log as missed|Display sticky|Discard|Display normally|Make high priority", , aEvent.GetValue("override-away"))
            .Add new_BPrefsControl("banner", "", "When I'm Busy...")
            .Add new_BPrefsControl("fancycycle", "override-busy", "Do default|Log as missed|Display sticky|Discard|Display normally|Make high priority", , aEvent.GetValue("override-busy"))

        End With
        .AddPage pp

        ' /* page */
        Set pp = new_BPrefsPage("Redaction", load_image_obj(g_MakePath(App.Path) & "etc\icons\icon-event-redact.png"), Me)
        With pp
            .SetMargin 90
            .SetWidth 420           ' // bug in prefs kit makes the default page width 400 pixels
            .Add new_BPrefsControl("fancycycle", "redact-when", "Never|Personal or higher|Private or higher|Confidential or higher", "When:", uRedactionIndex(aEvent.GetValue("redact-when")))
            .Add new_BPrefsControl("fancyedit", "redact-with", aEvent.GetValue("redact-with"), "With:")
            .Add new_BPrefsControl("label", "", "If a notification's sensitivity rating is the same or higher than the value selected, its body text will be replaced with the text above.")
            .Add new_BPrefsControl("fancytoggle2", "redact-always", "Redact on-screen content as well?", , aEvent.GetValue("redact-always"))

        End With
        .AddPage pp
        .Go
        g_SetWindowIconToAppResourceIcon .hWnd

Dim dw As Long

        dw = GetWindowLong(.hWnd, GWL_STYLE)
        dw = dw And (Not WS_MINIMIZEBOX)
        SetWindowLong .hWnd, GWL_STYLE, dw
        SetWindowLong .hWnd, GWL_HWNDPARENT, hWndPanel
        EnableWindow hWndPanel, 0

Dim rcOwner As RECT
Dim rcThis As RECT

        GetWindowRect hWndPanel, rcOwner
        GetWindowRect .hWnd, rcThis
        g_RectNormalise rcThis
        g_MoveWindow .hWnd, _
                     rcOwner.Left + Fix(((rcOwner.Right - rcOwner.Left) - rcThis.Right) / 2), _
                     rcOwner.Top + Fix(((rcOwner.Bottom - rcOwner.Top) - rcThis.Bottom) / 2)

        g_ShowWindow .hWnd, True, True

    End With

End Sub

Public Function uRedactionIndex(ByVal Redaction As Long) As String

    uRedactionIndex = CStr((Redaction \ 16) + 1)

End Function

Public Function uRedactionFromIndex(ByVal Index As Long) As String

    Index = Index - 1
    uRedactionFromIndex = CStr(Index * 16)

End Function
