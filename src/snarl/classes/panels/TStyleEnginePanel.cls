VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyleEnginePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mhWndPanel As Long

Dim mCurrentEngine As TStyleEngine

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

    uUpdateList

End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pc As BControl
Dim sz As String
Dim i As Long

    Select Case Control.GetName()

    Case "engine_list"
        Set mCurrentEngine = g_StyleRoster.EngineAt(Val(Value))
        If Not (mCurrentEngine Is Nothing) Then
            ' /* configure the toolbar */
            If Control.Page.Find("style_engine_toolbar", pc) Then
                prefskit_SetItem pc, 1, "enabled", IIf(mCurrentEngine.State = SN_SE_LOADED, 0&, 1&)     ' // "Start"
                prefskit_SetItem pc, 2, "enabled", IIf((mCurrentEngine.State = SN_SE_LOADED) And (Not mCurrentEngine.IsSystem), 1&, 0&)     ' // "Stop"
                prefskit_SetItem pc, 3, "enabled", IIf(mCurrentEngine.State = SN_SE_LOADED, 1&, 0&)     ' // "Details"
                prefskit_SetItem pc, 4, "enabled", mCurrentEngine.IsConfigurable                        ' // "Configure..."

                prefskit_SetItem pc, 6, "enabled", IIf((mCurrentEngine.State = SN_SE_LOADED) And (Not mCurrentEngine.IsSystem), 1&, 0&)     ' // "Show in Explorer"

            End If

            ' /* set the description label */
            prefskit_SafeSetText Control.Page.Panel, "engine_desc", mCurrentEngine.Description

        End If


    Case "style_engine_toolbar"
        If (mCurrentEngine Is Nothing) Then _
            Exit Sub

        Select Case Val(Value)
        Case 1
            ' /* start */
            If g_StyleRoster.Load(mCurrentEngine.Name, True, True) Then _
                uUpdateList

        Case 2
            ' /* stop */
            sz = mCurrentEngine.Name
            Set mCurrentEngine = Nothing
            If g_StyleRoster.Unload(sz, True) Then _
                uUpdateList

        Case 3
            ' /* details */
            mCurrentEngine.DoDetailsNotification

        Case 4
            ' /* configure... */
            mCurrentEngine.Configure

        Case 6
            ' /* show in Explorer */
            ShellExecute 0, "open", mCurrentEngine.Engine.Path, vbNullString, vbNullString, SW_SHOWNORMAL

        End Select

'    Case "lb>engines"
'        mSelectedEngine = ""
'
'        With g_StyleRoster
'            If .EngineAt(Val(Value), pEngineInfo) Then
'
'                mSelectedEngine = pEngineInfo.Name
'
'                If mPanel.Find("lbl>desc", pc) Then
'                    If (pEngineInfo.Obj Is Nothing) Then
'                        pc.SetText "Not loaded"
'
'                    Else
'                        pc.SetText pEngineInfo.Description
'
'                    End If
'
'                End If
'
''                If mPage.Find("lbl>path", pc) Then _
'                    pc.SetText pEngineInfo.Path
'
'                If mPanel.Find("fb2>loadunload", pc) Then
'                    If (pEngineInfo.Obj Is Nothing) Then
'                        pc.SetText "Load"
'
'                    Else
'                        pc.SetText "Unload"
'
'                    End If
'
'                    pc.SetEnabled ((pEngineInfo.Flags And &H80000000) = 0)
'
'                End If
'
'            Else
'
'            End If
'
'        End With
'
'
'    Case "ftb>engine"
'
'        ' /* retrieve info for the currently selected engine */
'
'        If Not mPanel.Find("lb>engines", pc) Then _
'            Exit Sub
'
'        If Not g_StyleRoster.EngineAt(Val(pc.GetValue), pEngineInfo) Then _
'            Exit Sub
'
'        Select Case Val(Value)
'        Case 1
'
''        Case 3
''            If g_StyleRoster.StylesForEngine(mSelectedEngine, sz) Then
''                If MsgBox("Uninstalling this style engine will remove the following styles:" & vbCrLf & vbCrLf & sz & vbCrLf & vbCrLf & "Are you sure you want to do this?", _
''                        vbQuestion Or vbYesNo, App.Title) = vbYes Then
''
''                End If
''
''            End If
'
'        Case 3
'
'Dim szEngine As String
'
'            szEngine = InputBox("Style Engine to install?", App.Title)
'            If szEngine = "" Then _
'                Exit Sub
'
'            If g_GetExtension(szEngine) = "" Then _
'                szEngine = szEngine & ".styleengine"
'
'            If g_StyleRoster.Load(szEngine, sz, True) Then
'                g_PrivateNotify "", "", szEngine & " installed ok", -1, g_MakePath(App.Path) & "etc\icons\style_engine.png"
'                uUpdateList
'
'                If g_GetUserFolderPath(sz) Then
'                    sz = g_MakePath(sz) & "styles\" & szEngine
'                    If Not g_Exists(sz) Then _
'                        uCreateBlankFile sz
'
'                End If
'
'            Else
'                MsgBox sz, vbCritical Or vbOKOnly, "Error Installing Style Engine"
'
'            End If
'
'        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName()
    Case "engine_list"
        ' /* configure... */
        mCurrentEngine.Configure

    End Select

'    Case "fb2>done"
'        mPanel.Quit
'
'    Case "fb2>loadunload"
'        If (g_StyleRoster Is Nothing) Then _
'            Exit Sub
'
'        Select Case Control.GetText
'        Case "Load"
'            If g_StyleRoster.Load(mSelectedEngine, sz, True) Then
'                MsgBox "Style Engine '" & mSelectedEngine & "' loaded", vbInformation Or vbOKOnly, App.Title
'                uUpdateList
'
'            Else
'                MsgBox "Style Engine '" & mSelectedEngine & "' failed to load" & vbCrLf & vbCrLf & sz, vbInformation Or vbOKOnly, App.Title
'
'            End If
'
'        Case "Unload"
'
'            If g_StyleRoster.StylesForEngine(mSelectedEngine, sz) Then
'                If MsgBox("Uninstalling this style engine will remove the following styles:" & vbCrLf & vbCrLf & _
'                          sz & vbCrLf & "Are you sure you want to do this?", _
'                          vbQuestion Or vbYesNo, App.Title) = vbYes Then
'
'                    If g_StyleRoster.Unload(mSelectedEngine, sz) Then
'                        MsgBox "Style Engine '" & mSelectedEngine & "' unloaded", vbInformation Or vbOKOnly, App.Title
'                        uUpdateList
'
'                    Else
'                        MsgBox "Error unloading Style Engine", vbExclamation Or vbOKOnly, App.Title
'
'                    End If
'
'                End If
'            End If
'
'        End Select
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
Dim pm As CTempMsg

    mhWndPanel = hWndPanel
    Set mPanel = New BPrefsPanel

    With mPanel
        .SetHandler Me
        .SetTitle "Installed Style Engines"
        .SetWidth 500
        .SetWindow 1                  ' // SPECIAL HACK: tells the BPrefsPanel not to auto-display during Go()

        Set pp = new_BPrefsPage("", , Me)

        With pp
            .SetMargin 0

            Set pm = New CTempMsg
            pm.Add "item-height", 36&
            Set pc = new_BPrefsControl("listbox", "engine_list", , , , pm)
            pc.SizeTo 0, 172
            .Add pc
        
            Set pc = new_BPrefsControl("label", "engine_desc")
            pc.SizeTo 0, 26
            .Add pc
        
            .Add new_BPrefsControl("fancytoolbar", "style_engine_toolbar", "Start|Stop|Details|Configure...||Show in Explorer")
'            .Add new_BPrefsControl("fancybutton2", "fb2>loadunload", "Unload")
            .Add new_BPrefsControl("fancybutton2", "fb2>install", "Install new Style Engine...")
    
        End With

        .AddPage pp
        .Go
        g_SetWindowIconToAppResourceIcon .hWnd

    End With

Dim rcConfig As RECT
Dim rc As RECT

    SetWindowLong mPanel.hWnd, GWL_HWNDPARENT, hWndPanel
    EnableWindow hWndPanel, 0

    GetWindowRect hWndPanel, rc
    GetWindowRect mPanel.hWnd, rcConfig
    g_RectNormalise rcConfig

    g_MoveWindow mPanel.hWnd, _
                 rc.Left + Fix(((rc.Right - rc.Left) - rcConfig.Right) / 2), _
                 rc.Top + Fix(((rc.Bottom - rc.Top) - rcConfig.Bottom) / 2)

    g_ShowWindow mPanel.hWnd, True, True

End Sub

Private Function uUpdateList() As String
Dim pc As BControl

    If Not (mPanel.Find("engine_list", pc)) Then _
        Exit Function

Dim pse As TStyleEngine
Dim sz As String
Dim n As Long
Dim i As Long

    n = MAX(Val(pc.GetValue), 1)

    If Not (g_StyleRoster Is Nothing) Then
        If g_StyleRoster.CountEngines > 0 Then

            With g_StyleRoster.EngineList
                .Rewind
                Do While .GetNextTag(pse) = B_OK
                    sz = sz & pse.InfoString & "#?0#?"
                    Select Case pse.State
                    Case SN_SE_LOADED
                        sz = sz & "Loaded (" & pse.StyleCountString & ")"

                    Case SN_SE_NOT_LOADED
                        sz = sz & "Not loaded"

                    Case SN_SE_FAILED
                        sz = sz & "Failed: " & pse.Error

                    End Select
                    
                    sz = sz & "|"

                Loop

            End With

            pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

            ' /* set the icons */

            With g_StyleRoster.EngineList
                .Rewind
                Do While .GetNextTag(pse) = B_OK
                    i = i + 1
                    prefskit_SetItem pc, i, "image-file", g_MakePath(App.Path) & "etc\icons\" & IIf(pse.State = SN_SE_LOADED, "good", "bad") & ".png"

                Loop

            End With

        Else
            pc.SetText "<None>"
            n = 1

        End If

    Else
        pc.SetText "<Style Roster not running>"
        n = 1

    End If

    pc.SetValue CStr(n)

End Function

Private Sub uCreateBlankFile(ByVal Path As String)

    On Error Resume Next

Dim n As Integer

    n = FreeFile()
    Open Path For Output As #n
'    Print #n, ""
    Close #n

End Sub

