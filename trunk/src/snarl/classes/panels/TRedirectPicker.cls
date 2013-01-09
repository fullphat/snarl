VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRedirectPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mhWndOwner As Long
Dim mStyle As String

Dim mSelectedStyle As TStyle

Public Event Closed()
Public Event Saved(ByVal Name As String, ByVal Flags As SN_REDIRECTION_FLAGS)

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim dw As SN_REDIRECTION_FLAGS
Dim sz As String

    Select Case Control.GetName()
    Case "style_and_scheme"
        sz = uDoNewRedirectMenu(mPanel.hWnd, BW_Frame(Control.hWnd).BottomLeft)
        If sz <> "" Then
            mStyle = sz
            Control.SetText style_MakeFriendly(sz)
            prefskit_SafeEnable mPanel, "save_button", True
            
            If g_StyleRoster.Find(style_GetStyleName(sz), mSelectedStyle) Then _
                prefskit_SafeEnable mPanel, "configure", mSelectedStyle.IsConfigurable

        End If

    Case "configure"
        If NOTNULL(mSelectedStyle) Then _
            mSelectedStyle.ShowPrefs 0

    Case "save_button"
        If mStyle <> "" Then
            If prefskit_GetValue(mPanel, "when_active") = "1" Then _
                dw = dw Or SN_RF_WHEN_ACTIVE

            If prefskit_GetValue(mPanel, "when_away") = "1" Then _
                dw = dw Or SN_RF_WHEN_AWAY

            If prefskit_GetValue(mPanel, "when_busy") = "1" Then _
                dw = dw Or SN_RF_WHEN_BUSY

            RaiseEvent Saved(mStyle, dw)
            mPanel.Quit

        End If

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

    EnableWindow mhWndOwner, -1
    g_ShowWindow mhWndOwner, True, True
    RaiseEvent Closed

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long, Optional ByVal Style As String, Optional ByVal Flags As SN_REDIRECTION_FLAGS)
Dim pp As BPrefsPage
Dim b As Boolean

'    MsgBox hWndPanel

    mhWndOwner = hWndPanel
    mStyle = Style

    If Style <> "" Then
        If g_StyleRoster.Find(style_GetStyleName(Style), mSelectedStyle) Then _
            b = mSelectedStyle.IsConfigurable

    End If

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle IIf(Style <> "", "Edit Redirect", "New Redirect")
        .SetWidth 160
        .SetWindow 1

        ' /* page */

        Set pp = new_BPrefsPage("", , Me)
        With pp
            .SetMargin 24
            .SetWidth 290           ' // bug in prefs kit makes the default page width 400 pixels
            .Add new_BPrefsControl("banner", "", "Redirector")
            .Add new_BPrefsControl("fancybutton2", "style_and_scheme", IIf(Style <> "", style_MakeFriendly(Style), "<Pick>"), , , , (Style = ""))
'            .Add new_BPrefsControl("fancybutton2", "configure", "Configure...", , , , b)

            .Add new_BPrefsControl("banner", "", "Invoke This Redirect...")
            .Add new_BPrefsControl("fancytoggle2", "when_active", "When I'm at my computer", "", IIf(Flags And SN_RF_WHEN_ACTIVE, "1", "0"))
            .Add new_BPrefsControl("fancytoggle2", "when_away", "When I'm away", "", IIf(Flags And SN_RF_WHEN_AWAY, "1", "0"))
            .Add new_BPrefsControl("fancytoggle2", "when_busy", "When I'm busy", "", IIf(Flags And SN_RF_WHEN_BUSY, "1", "0"))

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "save_button", "Ok", , , , (Style <> ""))

        End With

        .AddPage pp

        ' /* page */

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

Private Function uDoNewRedirectMenu(ByVal hWnd As Long, ByRef Pos As BPoint) As String
Dim pmSub As BMenu
Dim pi As MImage
Dim ps As TStyle
Dim pm As BMenu

    Set pm = New BMenu

    With g_StyleRoster
        .Rewind
        Do While .GetNextStyle(ps)
            If ps.IsRedirect Then
                Set pi = new_ImageFromFile(ps.SafeIconPath)
                Set pmSub = uGetSchemes2(ps, pi)
                pm.AddItem new_BMenuItem("", ps.Name, pi, , , , , pmSub)

            End If
        Loop

    End With

    If pm.CountItems = 0 Then _
        pm.AddItem new_BMenuItem("", "None", , False)

    If pm.Track(hWnd, Pos) Then _
        uDoNewRedirectMenu = pm.SelectedItem.Name


'Dim pmi As OMMenuItem
'Dim ps As TStyle
'Dim pm As OMMenu
'Dim pmSub As OMMenu
'
'    Set pm = New OMMenu
'
'    With g_StyleRoster
'        .Rewind
'        Do While .GetNextStyle(ps)
'            If ps.IsRedirect Then
'                Set pmSub = uGetSchemes(ps)
'                pm.AddItem pm.CreateItem("", ps.Name, , , , , , pmSub)
'
'            End If
'        Loop
'
'    End With
'
'    If pm.CountItems = 0 Then _
'        pm.AddItem pm.CreateItem("", "None", , False)
'
'    Set pmi = pm.Track(hWnd, Pos)
'    If NOTNULL(pmi) Then _
'        uDoNewRedirectMenu = pmi.Name

End Function
'
'Private Function uGetSchemes(ByRef Style As TStyle) As OMMenu
'Dim i As Long
'
'    Set uGetSchemes = New OMMenu
'
'    With Style
'        If .CountSchemes Then
'            For i = 1 To .CountSchemes
'                uGetSchemes.AddItem uGetSchemes.CreateItem(Style.Name & "/" & .SafeSchemeAt(i), .SchemeAt(i))
'
'            Next i
'
'        Else
'            uGetSchemes.AddItem uGetSchemes.CreateItem("", "No schemes")
'
'        End If
'    End With
'
'End Function

Private Function uGetSchemes2(ByRef Style As TStyle, ByRef Icon As MImage) As BMenu
Dim i As Long

    Set uGetSchemes2 = New BMenu

    With Style
        If .CountSchemes Then
            For i = 1 To .CountSchemes
                uGetSchemes2.AddItem new_BMenuItem(Style.Name & "/" & .SafeSchemeAt(i), .SchemeAt(i), Icon)

            Next i

        Else
            uGetSchemes2.AddItem new_BMenuItem("", "No schemes", , False)

        End If
    End With

End Function

