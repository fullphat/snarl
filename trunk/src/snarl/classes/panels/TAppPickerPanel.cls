VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAppPickerPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mhWndOwner As Long
Dim mType As String

Public Event Closed(ByVal Selection As String)

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

'    Select Case Control.GetName()
'    Case "host"
'        Debug.Print Value
'        prefskit_SafeEnable mPanel, "test_button", (Value <> "")
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pc As BControl

'    Select Case Control.GetName()
'    Case "test_button"
'        If mPanel.Find("host", pc) Then
'            Control.SetEnabled False
'            Set theSocket = New CSocket
'            If theSocket.Connect(pc.GetText(), IIf(mType = "subs", SNP_DEFAULT_PORT, GNTP_DEFAULT_PORT)) Then
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
'    Case "fb2>done"
''        If mPage.Find("ia>addr", pc) Then _
''            g_AddRemoteComputer pc.GetText()
''
''        mPanel.Quit
'
'    End Select

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
Dim pc As BControl
Dim sz As String
Dim c As Long
Dim i As Long

    If mPanel.Find("apps", pc) Then
        With pc
            c = prefskit_GetItemCount(pc)
            If c Then
                For i = 1 To c
                    Debug.Print prefskit_GetItem(pc, "subtext", i)
                    sz = sz & prefskit_GetItem(pc, "subtext", i) & ";"

                Next i
            End If

        End With
    End If

    EnableWindow mhWndOwner, -1
    g_ShowWindow mhWndOwner, True, True
    RaiseEvent Closed(g_SafeLeftStr(sz, Len(sz) - 1))

End Sub

Private Sub KPrefsPanel_Ready()
Dim pc As BControl

    If mPanel.Find("apps", pc) Then _
        uSetAppsList pc

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long, ByVal PanelType As String, ByVal CurrentSelection As String)

    mhWndOwner = hWndPanel
    mType = PanelType

'        .AddItem .CreateItem("frwd", "Forward to remote computer...")
'        .AddItem .CreateItem("subs", "Subscribe to remote SNP server...")
'        .AddItem .CreateItem("subg", "Subscribe to remote GNTP server...")

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle "Select Applications"
        .SetWidth 400

        ' /* page */

Dim pp As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

        Set pp = new_BPrefsPage("", , Me)

        With pp
            .SetMargin 24

            .Add new_BPrefsBanner("", "Applications", "")

            Set pm = New CTempMsg
            pm.Add "checkboxes", 1&
            pm.Add "item-height", 28&
            Set pc = new_BPrefsControl("listbox", "apps", , , , pm)
            pc.SizeTo 0, 190
            .Add pc

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancytoolbar", "", "Select All|Clear All|Invert Selection")

        End With

        .AddPage pp

        ' /* page */

        .Go
        g_SetWindowIconToAppResourceIcon .hWnd


Dim rcOwner As RECT
Dim rcThis As RECT
Dim dw As Long

        dw = GetWindowLong(.hWnd, GWL_STYLE)
        dw = dw And (Not WS_MINIMIZEBOX)
        SetWindowLong .hWnd, GWL_STYLE, dw

        SetWindowLong .hWnd, GWL_HWNDPARENT, hWndPanel
        EnableWindow hWndPanel, 0

        GetWindowRect hWndPanel, rcOwner
        GetWindowRect .hWnd, rcThis
        g_RectNormalise rcThis

        g_MoveWindow .hWnd, _
                     rcOwner.Left + Fix(((rcOwner.Right - rcOwner.Left) - rcThis.Right) / 2), _
                     rcOwner.Top + Fix(((rcOwner.Bottom - rcOwner.Top) - rcThis.Bottom) / 2)

        g_ShowWindow .hWnd, True, True

    End With

End Sub

Private Sub uSetAppsList(ByRef List As BControl)

    If (g_AppRoster Is Nothing) Then _
        Exit Sub

Dim sz As String
Dim i As Long

    With g_AppRoster
        If .CountApps Then
            For i = 1 To .CountApps
                With .AppAt(i)
                    sz = sz & .Name & "#?0#?" & .Signature & "|"
 
                End With
            Next i
        End If

    End With

    If sz = "" Then _
        sz = "None|"

    List.SetText g_SafeLeftStr(sz, Len(sz) - 1)

    With g_AppRoster
        If .CountApps Then
            For i = 1 To .CountApps
                With .AppAt(i)
                    prefskit_SetItem List, i, "image-file", .Icon
 
                End With
            Next i
        End If

    End With

End Sub

'Private Function uGetAppsList() As String
'Dim ps As ConfigSection
'Dim pf As ConfigFile
'Dim sz As String
'Dim i As Long
'
'    With New CFolderContent2
'        If .SetTo(g_SettingsPath()) Then
'            .Rewind
'            Do While .GetNextFile(sz)
'                If g_GetExtension(sz, True) = "class" Then
'                    Set pf = New ConfigFile
'                    pf.File = sz
'                    If pf.Load Then
'                        i = pf.FindSection("_cfg")
'                        If i Then
'                            Set ps = pf.SectionAt(i)
'                            If ps.GetValueWithDefault("name", "") <> "" Then _
'                                uGetAppsList = uGetAppsList & ps.GetValueWithDefault("name", "") & "|"
'
'                        End If
'                    End If
'                End If
'            Loop
'
'            If uGetAppsList = "" Then _
'                uGetAppsList = "None|"
'
'            uGetAppsList = g_SafeLeftStr(uGetAppsList, Len(uGetAppsList) - 1)
'
'        Else
'            Debug.Print g_SettingsPath() & " missing"
'
'        End If
'
'    End With
'
'End Function
