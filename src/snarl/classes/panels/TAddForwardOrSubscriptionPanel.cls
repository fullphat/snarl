VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAddForwardOrSubscriptionPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mhWndOwner As Long
Dim mIsForward As Boolean

Public Event Closed()
Public Event Saved(ByVal Name As String, ByVal Hosts As String, ByVal Protocol As String, ByVal ItemType As String, ByVal Password As String)

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Select Case Control.GetName()
'    Case "host"
'        prefskit_SafeEnable mPanel, "test_button", (Value <> "")
'        prefskit_SafeEnable mPanel, "save_button", (Value <> "")
'
'
'    Case "url", "desc"
'        If mType2 = SN_ST_COMET Then _
'            prefskit_SafeEnable mPanel, "save_button", (prefskit_SafeGetText(mPanel, "desc") <> "") And (prefskit_SafeGetText(mPanel, "url") <> "")

    End Select

    prefskit_SafeEnable mPanel, "save_button", ((prefskit_SafeGetText(mPanel, "name") <> "") And (prefskit_SafeGetText(mPanel, "hosts") <> ""))

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim sz As String

    Select Case Control.GetName()
'    Case "test_button"
'        If mPanel.Find("host", pc) Then
'            Control.SetEnabled False
'            Set theSocket = New CSocket
'            If theSocket.Connect(pc.GetText(), IIf((mType2 And SN_ST_TYPE_MASK) = SN_ST_GNTP, GNTP_DEFAULT_PORT, SNP_DEFAULT_PORT)) Then
'                Debug.Print "connecting..."
'                Set theTimer = new_BTimer(5000, True, True)
'
'            Else
'                MsgBox "The host name or IP address is invalid", vbExclamation Or vbOKOnly, App.Title
'                Control.SetEnabled True
'
'            End If
'        End If
'
''    Case "pick_apps"
''        Set theAppPicker = New TAppPickerPanel
''        theAppPicker.Go mPanel.hWnd, mType, ""

    Case "save_button"
        Select Case Val(prefskit_GetValue(mPanel, "protocol"))
        Case 1
            sz = "SNP"

        Case 2
            sz = "GNTP"

        End Select

        RaiseEvent Saved(prefskit_SafeGetText(mPanel, "name"), prefskit_SafeGetText(mPanel, "hosts"), sz, IIf(mIsForward, "forwarder", "subscription"), prefskit_SafeGetText(mPanel, "password"))
        mPanel.Quit

'        If mType2 = SN_ST_COMET Then
'            ' /* web listener */
'            sz = prefskit_SafeGetText(mPanel, "url")
'            If (prefskit_SafeGetText(mPanel, "desc") = "") Or (sz = "") Or (Not g_IsURL(sz)) Then
'                MsgBox "Must provide a name and URL", vbExclamation Or vbOKOnly, App.Title
'
'            ElseIf g_SubsRoster.AddWebListener(prefskit_SafeGetText(mPanel, "desc"), sz) Then
'                RaiseEvent Refresh
'                mPanel.Quit
'
'            End If
'
'        Else
'
'            sz = prefskit_SafeGetText(mPanel, "host")
'            If sz = "" Then _
'                Exit Sub
'
'            If (Not gDebugMode) And (g_IsLocalAddress(sz)) Then
'                MsgBox "Cannot forward notifications to local computer.", vbExclamation Or vbOKOnly, App.Title
'
'            Else
'                If g_SubsRoster.Add(mType2, prefskit_SafeGetText(mPanel, "desc"), sz, "", szErr) Then
'                    frmAbout.bForwardersChanged
'                    mPanel.Quit
'
'                Else
'                    MsgBox szErr, vbOKOnly Or vbExclamation, g_WindowText(mPanel.hWnd)
'
'                End If
'            End If
'        End If

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
    g_WindowToFront mhWndOwner, True
    RaiseEvent Closed

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndOwner As Long, ByRef Entry As ConfigSection, ByVal IsForward As Boolean)
Dim pp As BPrefsPage
Dim sz As String

    mhWndOwner = g_GetTopLevel(hWndOwner)
    mIsForward = IsForward

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetWindow 1

        If IsForward Then
            sz = " Forwarder"

        Else
            sz = " Subscription"

        End If

        If ISNULL(Entry) Then
            .SetTitle "Create" & sz

        Else
            .SetTitle "Edit" & sz

        End If

        .SetWidth 430
        sz = ""

        Set pp = new_BPrefsPage("", , Me)
        With pp
            .SetMargin 90
            If NOTNULL(Entry) Then _
                sz = Entry.GetValueWithDefault("name", "")
            .Add new_BPrefsControl("fancyedit", "name", sz, "Name:")
            .Add new_BPrefsControl("label", "", "Used to describe the " & IIf(mIsForward, "forwarder", "subscription") & " (""Home network"", ""Office"", etc).")

            sz = ""
            If NOTNULL(Entry) Then _
                sz = Entry.GetValueWithDefault("hosts", "")

            If mIsForward Then
                .Add new_BPrefsControl("fancyedit", "hosts", sz, "Hosts:")
                .Add new_BPrefsControl("label", "", "IP addresses of the computers to forward to.  Separate multiple entries with semicolons.")

            Else
                ' /* subscription */
                .Add new_BPrefsControl("fancyedit", "hosts", sz, "Host:")
                .Add new_BPrefsControl("label", "", "IP address of the computer to subscribe to.")

                sz = ""
                If NOTNULL(Entry) Then _
                    sz = Entry.GetValueWithDefault("password", "")

                .Add new_BPrefsControl("fancyedit", "password", sz, "Password:")

            End If

            sz = "1"
            If NOTNULL(Entry) Then
                Select Case LCase$(Entry.GetValueWithDefault("protocol", ""))
                Case "gntp"
                    sz = "2"
            
                End Select
            End If
            .Add new_BPrefsControl("fancycycle", "protocol", "SNP|GNTP", "Protocol:", sz)

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "save_button", "Save", , , , False)

        End With
        .AddPage pp

        .Go
        g_SetWindowIconToAppResourceIcon .hWnd

Dim dw As Long

        dw = GetWindowLong(.hWnd, GWL_STYLE)
        dw = dw And (Not WS_MINIMIZEBOX)
        SetWindowLong .hWnd, GWL_STYLE, dw
        SetWindowLong .hWnd, GWL_HWNDPARENT, mhWndOwner
        EnableWindow mhWndOwner, 0

        g_MoveWindow .hWnd, , , , , True, mhWndOwner
        g_ShowWindow .hWnd, True, True

    End With

End Sub

