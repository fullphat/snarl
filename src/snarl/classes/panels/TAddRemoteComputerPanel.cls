VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAddRemoteComputerPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mhWndOwner As Long
Dim mType2 As SN_SUBSCRIBER_TYPE

Public Event Closed()
Public Event Refresh()

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1
'Dim WithEvents theAppPicker As TAppPickerPanel

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Select Case Control.GetName()
    Case "host"
        prefskit_SafeEnable mPanel, "test_button", (Value <> "")
        prefskit_SafeEnable mPanel, "save_button", (Value <> "")


    Case "url", "desc"
        If mType2 = SN_ST_COMET Then _
            prefskit_SafeEnable mPanel, "save_button", (prefskit_SafeGetText(mPanel, "desc") <> "") And (prefskit_SafeGetText(mPanel, "url") <> "")

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pc As BControl
Dim szErr As String
Dim sz As String

    Select Case Control.GetName()
    Case "test_button"
        If mPanel.Find("host", pc) Then
            Control.SetEnabled False
            Set theSocket = New CSocket
            If theSocket.Connect(pc.GetText(), IIf((mType2 And SN_ST_TYPE_MASK) = SN_ST_GNTP, GNTP_DEFAULT_PORT, SNP_DEFAULT_PORT)) Then
                Debug.Print "connecting..."
                Set theTimer = new_BTimer(5000, True, True)

            Else
                MsgBox "The host name or IP address is invalid", vbExclamation Or vbOKOnly, App.Title
                Control.SetEnabled True

            End If
        End If

'    Case "pick_apps"
'        Set theAppPicker = New TAppPickerPanel
'        theAppPicker.Go mPanel.hWnd, mType, ""

    Case "save_button"

        If mType2 = SN_ST_COMET Then
            ' /* web listener */
            sz = prefskit_SafeGetText(mPanel, "url")
            If (prefskit_SafeGetText(mPanel, "desc") = "") Or (sz = "") Or (Not g_IsURL(sz)) Then
                MsgBox "Must provide a name and URL", vbExclamation Or vbOKOnly, App.Title

            ElseIf g_SubsRoster.AddWebListener(prefskit_SafeGetText(mPanel, "desc"), sz) Then
                RaiseEvent Refresh
                mPanel.Quit

            End If

        Else

            sz = prefskit_SafeGetText(mPanel, "host")
            If sz = "" Then _
                Exit Sub

            If (Not gDebugMode) And (g_IsLocalAddress(sz)) Then
                MsgBox "Cannot forward notifications to local computer.", vbExclamation Or vbOKOnly, App.Title

            Else
                If g_SubsRoster.Add(mType2, prefskit_SafeGetText(mPanel, "desc"), sz, "", szErr) Then
                    frmAbout.bForwardersChanged
                    mPanel.Quit

                Else
                    MsgBox szErr, vbOKOnly Or vbExclamation, g_WindowText(mPanel.hWnd)

                End If
            End If
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

    Set theTimer = Nothing
    EnableWindow mhWndOwner, -1
    g_ShowWindow mhWndOwner, True, True
    RaiseEvent Closed

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long, ByVal PanelType As SN_SUBSCRIBER_TYPE)
Dim szTitle As String
Dim sza As String
Dim pp As BPrefsPage

    mhWndOwner = hWndPanel
    mType2 = PanelType

'        .AddItem .CreateItem("frwd", "Forward to remote computer...")
'        .AddItem .CreateItem("subs", "Subscribe to remote SNP server...")
'        .AddItem .CreateItem("subg", "Subscribe to remote GNTP server...")

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        Select Case PanelType
        Case SN_ST_SNP3_FORWARD
            szTitle = "Forward to Remote Computer"
            sza = "forward notifications to"

        Case SN_ST_SNP3_SUBSCRIPTION
            szTitle = "Subscribe to Remote Server"
            sza = "subscribe to"

'        Case "subg"
'            szTitle = "GNTP Subscription"
'            sza = "subscribe to."

        Case SN_ST_COMET
            szTitle = "Web Listener"

        End Select

        .SetTitle szTitle
        .SetWidth 460

        ' /* page */

        Set pp = new_BPrefsPage("", , Me)

        With pp
            .SetMargin 90

            If mType2 = SN_ST_COMET Then
            
                .Add new_BPrefsControl("fancyedit", "desc", , "Description:")
                .Add new_BPrefsControl("fancyedit", "url", , "URL:")
'                .Add new_BPrefsControl("label", "", "")

            Else

'            .Add new_BPrefsControl("label", "", sza)
'            .Add new_BPrefsSeparator

                .Add new_BPrefsBanner("", "Destination", "")
                .Add new_BPrefsControl("fancyedit", "host", , "Host:")
                .Add new_BPrefsControl("label", "", "Enter the IP address of the computer to " & sza)
                .Add new_BPrefsControl("fancybutton2", "test_button", "Check", , , , False)
    
                .Add new_BPrefsControl("fancyedit", "desc", , "Description:")
                .Add new_BPrefsControl("label", "", "This can be used to describe the connection (e.g. ""Tom's Computer"", ""Office Server"", etc.)")
    
'    '            .Add new_BPrefsBanner("", "Applications", "")
'    '            .Add new_BPrefsControl("fancybutton2", "pick_apps", "Select Applications...")
'
'                .Add new_BPrefsControl("fancyedit", "apps", , "Apps:")
'                .Add new_BPrefsControl("label", "", "Enter the signatures of the applications you wish to " & _
'                                       IIf(mType2 = SN_ST_SNP3_FORWARD, "forward from", "subscribe to") & _
'                                       ".  Multiple entries should be separated with a semi-colon.  Refer to the user guide for more information.")
'
'    '            If PanelType = "frwd" Then
'    '                .Add new_BPrefsControl("label", "", "Enter the signatures of the applications you wish to forward (separate multiple entries with "";"").  To forward from all applications, leave this field empty.")
'    '
'    '            Else
'    '                .Add new_BPrefsControl("label", "", "Enter the signatures of the applications you wish to subscribe (separating multiple entries with "";"") to.  To subscribe to all applications, leave this field empty.")
'    '
'    '            End If

                .Add new_BPrefsControl("fancycycle", "protocol_type", "SNP 3.0|GNTP", "Protocol:", , , False)

            End If

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "save_button", "Save", , , , False)

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

'Private Sub theAppPicker_Closed(ByVal Selection As String)
'
'    MsgBox "Selected: " & Selection
'    Set theAppPicker = Nothing
'
'End Sub

Private Sub theSocket_OnConnect()

    Debug.Print "++ connected ++"

    Set theTimer = Nothing
    MsgBox "Remote computer connected successfully", vbInformation Or vbOKOnly, App.Title
    prefskit_SafeEnable mPanel, "test_button", True

End Sub

Private Sub theSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    Debug.Print "error: " & Description

End Sub

Private Sub theTimer_Pulse()

    MsgBox "Could not contact remote computer", vbExclamation Or vbOKOnly, App.Title
    Set theTimer = Nothing
    prefskit_SafeEnable mPanel, "test_button", True

End Sub
