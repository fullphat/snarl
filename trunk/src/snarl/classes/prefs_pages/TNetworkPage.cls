VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNetworkPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theAddPanel As TAddRemoteComputerPanel
Attribute theAddPanel.VB_VarHelpID = -1
Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

'    uRefreshSubscribers

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

    With mPage

        .SetMargin 24
        .Add new_BPrefsControl("banner", "", "Receiving")

'        Set pm = New CTempMsg
'        pm.Add "text", "Listen for incoming Growl or Snarl notifications?"
'        pm.Add "align", 1
'        .Add new_BPrefsControl("fancytoggle", "listen_for_snarl", , "", g_ConfigGet("listen_for_snarl"), pm)
'
'        pm.Replace "text", "Listen for incoming JSON notifications?"
'        .Add new_BPrefsControl("fancytoggle", "listen_for_json", , , g_ConfigGet("listen_for_json"), pm)
'
'        .Add new_BPrefsControl("label", "", "Snarl will listen on ports 9887 and 23053 for incoming notifications, and port 9889 for incoming JSON notifications.  You may need to adjust your firewall settings in order to receive these.")
'
'        .Add new_BPrefsControl("fancytoggle2", "allow_subs", "Allow remote computers to subscribe to notifications?", "", g_ConfigGet("allow_subs"))

'        Set pm = New CTempMsg
'        pm.Add "plus_minus", 1&
'        pm.Add "checkboxes", 1&
'        Set pc = new_BPrefsControl("listbox", "listener_list", , , , pm)
''        pc.SizeTo 0, 170
'        .Add pc
'
'        .Add new_BPrefsControl("fancyplusminus", "add_remove_listener")

'        .Add new_BPrefsControl("separator", "")
''        pm.Replace "text", "Forward notifications?"
''        .Add new_BPrefsControl("fancytoggle", "ft>forward", , "Sending:", "0", pm, False)

'        If gDebugMode Then
'            .Add new_BPrefsControl("label", "", "Snarl will forward notifications to the following computers:")
'            Set pm = New CTempMsg
'            pm.Add "plus_minus", 1&
'            pm.Add "item-height", 36&
'            pm.Add "checkboxes", 1&
'            Set pc = new_BPrefsControl("listbox", "subscriber_list", , , , pm)
'            pc.SizeTo 0, 170
'            .Add pc
'
'            .Add new_BPrefsControl("fancyplusminus", "fpm>forward")

'        End If

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
'Dim pc As BControl
'Dim sz As String
'Dim i As Long
'
'    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"
'
'    Select Case Control.GetName()
'    Case "ft>forward"
'        If Control.Page.Find("subscriber_list", pc) Then _
'            pc.SetEnabled (Value = "1")
'
'        If Control.Page.Find("fpm>forward", pc) Then _
'            pc.SetEnabled (Value = "1")
'
'    Case "listen_for_json", "listen_for_snarl"
'        g_ConfigSet Control.GetName, Value
'        If Control.GetName = "listen_for_json" Then
'            frmAbout.EnableJSON (Value = "1")
'
'        Else
'            frmAbout.EnableSNP (Value = "1")
'
'        End If
'
'    Case "fpm>forward"
'        If Value = "+" Then
'            uDoSubscriberMenu Control
'
'        Else
'            ' /* remove selected entry */
'            If mPage.Find("subscriber_list", pc) Then
'                i = Val(pc.GetValue())
'                If i > 0 Then
'                    sz = prefskit_GetItem(pc, "data", i)
'                    If g_SubsRoster.RemoveByGuid(sz) Then _
'                        uRefreshSubscribers
'
'                End If
'            End If
'        End If
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
'Dim ps As TSubscriber
'Dim i As Long
'Dim v As Long
'
'    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"
'
'    Select Case Control.GetName()
'    Case "subscriber_list"
'        If Notification = "update_list" Then
''            Control.SetText g_GetRemoteComputers
'
'            MsgBox "!! subscriber list !!"
'
'        ElseIf Notification = "refresh" Then
'            ' /* refresh subscriber list */
'            uRefreshSubscribers
'
'        ElseIf Notification = "checked" Then
'            i = Val(prefskit_GetValue(Control.Page.Panel, Control.GetName))
'            v = Val(prefskit_GetItem(Control, "checked", i))
'            Debug.Print CStr(i) & " --> " & CStr(v)
'            Set ps = g_SubsRoster.List.TagAt(i)
'            If Not (ps Is Nothing) Then _
'                ps.SetEnabled (v = 1)
'
'        End If
'
'    End Select

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
'    mPage.SetMargin 0

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

'Private Sub uRefreshSubscribers()
'
'    If Not gDebugMode Then _
'        Exit Sub
'
'Dim pc As BControl
'
'    If Not mPage.Find("subscriber_list", pc) Then _
'        Exit Sub
'
'Dim ps As TSubscriber
'Dim sz As String
'
'    With g_SubsRoster.List
'        .Rewind
'
'        Do While .GetNextTag(ps) = B_OK
'            sz = sz & ps.Description & "#?" & ps.Guid & "#?" & ps.Detail & "|"
'
'        Loop
'
'    End With
'
'    sz = g_SafeLeftStr(sz, Len(sz) - 1)
'    pc.SetText sz
'
'    If sz = "" Then _
'        Exit Sub
'
'    ' /* set the icons */
'
'Dim pm As CTempMsg
'Dim i As Long
'
'    Set pm = New CTempMsg
'
'    With g_SubsRoster.List
'        .Rewind
'        Do While .GetNextTag(ps) = B_OK
'            i = i + 1
'            Select Case ps.SubscriptionType
'            Case SN_ST_SNP3_FORWARD
'                sz = g_MakePath(App.Path) & "etc\icons\fwd-snp.png"
'
'            Case SN_ST_SNP3_SUBSCRIPTION
'                sz = g_MakePath(App.Path) & "etc\icons\sub-snp.png"
'
'            Case Else
'                sz = g_MakePath(App.Path) & "etc\icons\no_icon.png"
'
'            End Select
'
''            If ps.SubscriptionType = SN_ST_SNP3_SUBSCRIBER Then
''            ElseIf ps.SubscriptionType = SN_ST_GNTP Then
''                sz = g_MakePath(App.Path) & "etc\icons\growl_app.png"
''
''            Else
'
''            End If
'
'            pm.Replace "index", i
'            pm.Replace "image-file", sz
'            pm.Replace "checked", IIf(ps.IsEnabled, 1&, 0&)
'            pc.DoExCmd B_SET_ITEM, pm
'
'        Loop
'
'    End With
'
'End Sub
'
'Private Sub uDoSubscriberMenu(ByRef Control As BControl)
'Dim pi As OMMenuItem
'Dim rc As RECT
'
'    GetWindowRect Control.hWnd, rc
'
'    With New OMMenu
'        .AddItem .CreateItem(CStr(SN_ST_SNP3_FORWARD), "Forward to remote computer...")
'        .AddItem .CreateItem(CStr(SN_ST_SNP3_SUBSCRIPTION), "Subscribe to remote SNP server...")
''        .AddItem .CreateItem("subg", "Subscribe to remote GNTP server...")
'
'        Set pi = .Track(mPage.Panel.hWnd, new_BPoint(rc.Left, rc.Bottom))
'
'    End With
'
'    If (pi Is Nothing) Then _
'        Exit Sub
'
'    Select Case pi.Name
'    Case "subs"
'
'    Case "subg"
'
'    Case "frwd"
'
'    End Select
'
'    Set theAddPanel = New TAddRemoteComputerPanel
'    theAddPanel.Go Control.Page.Panel.hWnd, Val(pi.Name)
'
'End Sub
'
'Private Sub theAddPanel_Closed()
'
'    Set theAddPanel = Nothing
'
'End Sub
