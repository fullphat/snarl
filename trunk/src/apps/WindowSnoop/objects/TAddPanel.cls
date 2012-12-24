VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAddPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Done(ByRef NewItem As TRule)

Dim mPanel As BPrefsPanel
Dim mhWndPanel As Long

Dim mNewItem As TRule

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim szTitle As String
Dim szClass As String

    Debug.Print "--invoked: " & Control.GetName & "--"

    Select Case Control.GetName()
    Case "submit"

        szTitle = prefskit_SafeGetText(mPanel, "title")
        szClass = prefskit_SafeGetText(mPanel, "class")

        mNewItem.SetTo "", szTitle, szClass
        RaiseEvent Done(mNewItem)

        mPanel.Quit


    Case "url"
        ShellExecute mPanel.hWnd, "open", "http://msdn.microsoft.com/en-us/library/ms633518(v=VS.85).aspx", vbNullString, vbNullString, SW_SHOW

    End Select

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

    EnableWindow mhWndPanel, -1
    g_ShowWindow mhWndPanel, True, True

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long)
Dim pp As BPrefsPage
Dim pc As BControl

    Set mNewItem = New TRule

    mhWndPanel = hWndPanel
    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle "New Rule"
        .SetWidth 400

        Set pp = new_BPrefsPage("", , Me)

        With pp
            .SetMargin 96

            .Add new_BPrefsControl("banner", "", "Criteria")
'            .Add new_BPrefsControl("fancyfolderpicker", "path")
            .Add new_BPrefsControl("fancyedit", "title", "*", "Title:")
            .Add new_BPrefsControl("fancyedit", "class", "*", "Class:")

            .Add new_BPrefsControl("label", "", "Criteria matching is done using a limited form of regular expression matching (specifically the Visual Basic LIKE operator).  Click the button below to read more about the format used and the functionality available.")
            .Add new_BPrefsControl("fancybutton2", "url", "Open web page")

            .Add new_BPrefsSeparator
            Set pc = new_BPrefsControl("fancybutton2", "submit", "Ok")
            pc.SizeTo 80, 0             ' // a zero height here means "let control decide"
            .Add pc

        End With

        .AddPage pp

        .Go
        g_SetWindowIconToAppResourceIcon .hWnd


Dim dw As Long

        dw = GetWindowLong(.hWnd, GWL_STYLE)
        dw = dw And (Not WS_MINIMIZEBOX)
        SetWindowLong .hWnd, GWL_STYLE, dw

Dim rcConfig As RECT
Dim rc As RECT

        SetWindowLong .hWnd, GWL_HWNDPARENT, hWndPanel
        EnableWindow hWndPanel, 0

        GetWindowRect hWndPanel, rc
        GetWindowRect .hWnd, rcConfig
        g_RectNormalise rcConfig

        g_MoveWindow .hWnd, _
                     rc.Left + Fix(((rc.Right - rc.Left) - rcConfig.Right) / 2), _
                     rc.Top + Fix(((rc.Bottom - rc.Top) - rcConfig.Bottom) / 2)

        g_ShowWindow .hWnd, True, True

    End With

End Sub

