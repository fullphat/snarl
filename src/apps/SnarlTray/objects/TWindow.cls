VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2

Private Declare Function SetTrayHook Lib "geotrayspy.dll" (ByVal szServerName As String) As Boolean
Private Declare Function UnsetTrayHook Lib "geotrayspy.dll" () As Boolean
Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long

Dim mhWnd As Long
Dim mToken As Long

Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    On Error Resume Next

    Select Case uMsg
    Case WM_CREATE
        Form1.Add "WM_CREATE"
        mhWnd = hWnd
        slInitApp hWnd, App.ProductName, "SnarlTray", g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "icon.png"
        uStart


    Case WM_DESTROY
        Form1.Add "WM_DESTROY"
        UnsetTrayHook
        If mToken Then _
            sn41UnregisterApp mToken


    Case WM_CLOSE
        PostQuitMessage 0


    Case WM_COPYDATA
        uCopyData hWnd, wParam, lParam


    Case sn41Broadcast()
        Select Case wParam
        Case SNARL_LAUNCHED
            Form1.Add "Snarl launched"
            uRegisterWithSnarl

        Case SNARL_QUIT
            Form1.Add "Snarl quit"
            mToken = 0                      ' // no point in unregistering...
            UnsetTrayHook

        End Select


    Case sn41AppMsg()
        Select Case wParam
'        Case SNARL41_APP_PREFS
'            Form1.Add "_APP_PREFS"
'            uDoPrefs

        Case SNARL41_APP_ABOUT
            sn41EZNotify mToken, "", _
                         "SnarlTray " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments, _
                         App.LegalCopyright & vbCrLf & "Build " & CStr(App.Revision), , _
                         g_MakePath(App.Path) & "icon.png"

        End Select


    Case slMsg()
        Select Case wParam

        Case SL_ABOUT
            sn41EZNotify mToken, "", _
                         "SnarlTray " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments, _
                         App.LegalCopyright & vbCrLf & "Build " & CStr(App.Revision), , _
                         g_MakePath(App.Path) & "icon.png"

        Case SL_PREFS
        Case SL_SWITCH_TO

        Case SL_QUIT
            PostQuitMessage 0

        Case SL_SHOW_ALL
        Case SL_HIDE_ALL
        Case SL_COUNT_WINDOWS

        End Select


    Case WM_NOTIFICATION
        Form1.Add "NOTIFICATION: " & g_HexStr(wParam) & " " & CStr(lParam)
'        n = LoWord(wParam)

'        Select Case n
'        Case SNARL_NOTIFICATION_ACK, SNARL_NOTIFICATION_MENU
'            If Not uFind(lParam, pObj) Then _
'                Exit Function
'
'            Form1.Add "item found: " & pObj.Class
'
'            If n = SNARL_NOTIFICATION_ACK Then
'                pObj.Display
'
'            ElseIf n = SNARL_NOTIFICATION_MENU Then
'
'                sn41Hide lParam
'                pObj.UnRead = False
'
'                Select Case HiWord(wParam)
'                Case 1
'                    pObj.Display
'
'                Case 2
'                    pObj.Reply.Display
'
'                Case 3
'                    pObj.ReplyAll.Display
'
'                Case 4
'                    pObj.Forward.Display
'
'                End Select
'
'            End If
'
'
''        snChangeAttribute hr, SNARL_ATTRIBUTE_MENU, "Open#?1||Reply#?2|Reply All#?3|Forward#?4||Delete#?5"
'
'        End Select


    Case WM_TEST
        Form1.Add "TEST: " & wParam

'        If wParam = 0 Then
'            uNotifyMail Nothing
'
'        ElseIf wParam = 1 Then
'            uNotifyMail myItems.Item(1)
'
'        ElseIf wParam = 2 Then
'            uFindFirst olMeetingRequest, pObj
'            Debug.Print VarType(pObj)
'            Err.Clear
'            uNotifyAppointment pObj
'            Debug.Print Err.Description
'
'        End If

'    Case Else
'        form1.add "WM_" & g_HexStr(uMsg)

    End Select

End Function

Private Sub uRegisterWithSnarl()
Dim fOk As Boolean

    On Error Resume Next

    err.Clear
    fOk = SetTrayHook(CLASS_NAME)
    Debug.Print fOk
    Debug.Print err.Number

    If (fOk = False) Or (err.Number <> 0) Then
        Form1.Add "failed to register tray hook (" & err.Description & ")"
        Exit Sub

    End If

'    mSpyMessage = RegisterWindowMessage(GSSPY_DOWNLOAD)

    mToken = sn41RegisterApp(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", mhWnd, WM_NOTIFICATION, SNARL41_APP_IS_WINDOWLESS Or SNARL41_APP_HAS_ABOUT)
    If mToken Then
        Form1.List1.Tag = CStr(mToken)
        Form1.Add "Registered with Snarl (token=" & CStr(mToken) & ")"
        sn41AddClass mToken, "info", "Informational notifications"
        sn41AddClass mToken, "warn", "Warning notifications"
        sn41AddClass mToken, "crit", "Critical notifications"
        sn41AddClass mToken, "none", "Other notifications"

    Else
        Form1.Add "Error registering with Snarl"

    End If

End Sub

Private Sub uStart()

    If sn41IsSnarlRunning() Then
        uRegisterWithSnarl

    Else
        Form1.Add "Snarl not running, waiting..."

    End If

End Sub

Private Sub uCopyData(ByVal hWnd As Long, ByVal wParam As Long, ByVal lParam As Long)
Static cds As COPYDATASTRUCT

    CopyMemory cds, ByVal lParam, Len(cds)
'    g_Debug "TDaemonWindow.uCopyData(): cb=" & CStr(cds.cbData) & " dw=" & CStr(cds.dwData) '& " tail=" & mUsedTail

    ' /* from the systray? */

    If cds.dwData <> 1 Then _
        Exit Sub

    ' /* the COPYDATASTRUCT includes 2 longwords of extra data:
    '       +0 = unknown
    '       +4 = command
    '       +8 = NOTIFYICONDATA struct
    ' */

Static pNotify6 As NOTIFYICONDATA6
Static pNotify5 As NOTIFYICONDATA5
Static pNotify As NOTIFYICONDATA
Static iCmd As Long

    CopyMemory pNotify, ByVal cds.lpData + 8, Len(pNotify)
    CopyMemory iCmd, ByVal cds.lpData + 4, 4

'    Form1.List1.AddItem "uCopyData(): cb=" & pNotify.cbSize & " command=" & iCmd
'    Form1.Log "uCopyData(): cb=" & pNotify.cbSize & " command=" & iCmd & " " & uCmdStr(iCmd)

    If (iCmd <> NIM_ADD) And (iCmd <> NIM_MODIFY) Then
        Form1.List1.AddItem "ignoring non-add or modify..."
        Exit Sub

    End If

Dim fBalloonTip As Boolean
Dim szClass As String
Dim szTitle As String
Dim szText As String
Dim szIcon As String
Dim lIcon As Long

    Select Case pNotify.cbSize
    Case Is >= SIZEOF_NOTIFYICONDATA6
        CopyMemory pNotify6, ByVal cds.lpData + 8, Len(pNotify6)
        With pNotify6
            If (.uFlags And NIF_INFO) Then
                ' /* balloon tip */
                szTitle = g_TrimStr(StrConv(.szInfoTitle, vbFromUnicode))
                szText = g_TrimStr(StrConv(.szInfo, vbFromUnicode))
                szClass = "none"

                If pNotify.cbSize > SIZEOF_NOTIFYICONDATA6 Then _
                    Form1.List1.AddItem "post-V6 struct (cb=" & CStr(pNotify.cbSize) & ")"

                If (.dwInfoFlags And NIIF_ICON_MASK) = NIIF_USER Then
'                    DestroyIcon pEntry.hInfoIcon
                    szIcon = "%" & CStr(CopyIcon(.hIcon))

                Else
                    Select Case (.dwInfoFlags And NIIF_ICON_MASK)
                    Case NIIF_INFO
                        szIcon = "!system-info"
                        szClass = "info"

                    Case NIIF_WARNING
                        szIcon = "!system-warning"
                        szClass = "warn"

                    Case NIIF_ERROR
                        szIcon = "!system-critical"
                        szClass = "crit"
                
                    End Select

                End If

                fBalloonTip = True

            End If

        End With


    Case SIZEOF_NOTIFYICONDATA5
        CopyMemory pNotify5, ByVal cds.lpData + 8, Len(pNotify5)
        With pNotify5
            If (.uFlags And NIF_INFO) Then
                ' /* balloon tip */
                szTitle = g_TrimStr(StrConv(.szInfoTitle, vbFromUnicode))
                szText = g_TrimStr(StrConv(.szInfo, vbFromUnicode))
                szClass = "none"

                If (.dwInfoFlags And NIIF_ICON_MASK) = NIIF_USER Then
'                    DestroyIcon pEntry.hInfoIcon
                    szIcon = "%" & CStr(CopyIcon(.hIcon))

                Else
                    Select Case (.dwInfoFlags And NIIF_ICON_MASK)
                    Case NIIF_INFO
                        szIcon = "!system-info"
                        szClass = "info"

                    Case NIIF_WARNING
                        szIcon = "!system-warning"
                        szClass = "warn"

                    Case NIIF_ERROR
                        szIcon = "!system-critical"
                        szClass = "crit"
                
                    End Select

                End If

                fBalloonTip = True

            End If
        End With

    Case Else


    End Select

Dim szSource As String
Dim lPid As Long
Static i As Long

    If fBalloonTip Then
        Form1.List1.AddItem "Balloon tip: " & szTitle & " / " & szText

        If (szText = "") And (szTitle = "") Then _
            Exit Sub

        ' /* get sending process details */

        If GetWindowThreadProcessId(pNotify.hWnd, lPid) <> 0 Then
            Form1.List1.AddItem "sent from " & g_HexStr(pNotify.hWnd) & " (pid " & CStr(lPid) & ")"
            szSource = uGetSource(lPid)
            Form1.List1.AddItem "source: " & szSource

        End If

        i = sn41EZNotify(mToken, szClass, szTitle, szText & szSource, , szIcon)

    End If

End Sub

Private Function uGetSource(ByVal pid As Long) As String
Dim szPath As String
Dim szDesc As String

    If Not g_ProcessNameAndPath(pid, szPath, szDesc) Then _
        Exit Function

    uGetSource = IIf(szDesc = "", g_FilenameFromPath(szPath), szDesc)

    If uGetSource <> "" Then _
        uGetSource = vbCrLf & vbCrLf & "Source: " & uGetSource

End Function
