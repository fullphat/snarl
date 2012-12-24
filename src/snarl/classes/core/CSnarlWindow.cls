VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSnarlWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theActionsPopup As TAppsPopUpWindow
Attribute theActionsPopup.VB_VarHelpID = -1

Private Const TIMER_MAIN_TICK = 22
Private Const TIMER_ASYNC_FADE = 23
Private Const TIMER_TRACK_POINTER = 24
Private Const TIMER_PULSE = 25

Dim theWindow As NitroWindow
Dim mView As mfxView

Dim m_Alpha As Integer
Dim mOriginalTimeout As Long            ' // in seconds!

Dim mInfo As T_NOTIFICATION_INFO

Dim mCurrent As Long
Dim mTargetAlpha As Long
Dim mAlphaStep As Single

    ' /* 37.33 */
Dim m_Disable As Long                   ' // >0 = timer paused, <=0 timer running
    ' /* 37.80
Dim mVisible As Boolean
    ' /* 39.62 */
Dim mInstance As IStyleInstance
    ' /* 38.9 */
'Dim theBitmap As mfxBitmap
Dim mPointerInside As Boolean
    ' /* 38.132 */
Dim mMenu As OMMenu                     ' // set by calling snChangeAttribute(SNARL_ATTRIBUTE_MENU)
    ' /* R2.2 */
Dim mStyleFlags As S_STYLE_FLAGS
Dim mFuzzy As Boolean
Dim mQuitting As Boolean
Dim mTick As Long                       '// timeout/100
    ' /* R2.3 */
Dim mClickThru As Boolean               '// set during Create() - taken from class settings and then modified by style flags
Dim mStylePath As String
Dim mMenuOpen As Boolean
    ' /* R2.4 */
Dim mCloseRect As BRect
Dim mActionsRect As BRect
Dim mFadeStart As Long
Dim mOriginalAlpha As Long
    ' /* R2.5 */
Dim mScriptCallback As TScriptCallback
    ' /* R2.6 */
Dim mCallbackRect As BRect
Dim mContent As mfxBitmap
Dim mGadgetPressed As String

Implements MWndProcSink

Friend Function Create(ByRef Info As T_NOTIFICATION_INFO, ByRef Instance As IStyleInstance, ByVal StyleFlags As S_STYLE_FLAGS, ByVal StylePath As String) As M_RESULT

    ' /* must have a class */

    If (Info.ClassObj Is Nothing) Then
        Create = M_INVALID_ARGS
        Exit Function

    End If

'    Info.StyleToUse = style_GetSchemeName(Info.StyleToUse)

    LSet mInfo = Info
    mClickThru = mInfo.ClassObj.IsClickThru
    mStylePath = StylePath

'    Set mOpenButton = g_CreateButton(new_BPoint(66, 26))
'    Set mActionsGadget = uCreateButton(new_BPoint(24, 24), True)

    ' /* R2.6 - style can provide customised gadgets */

'    Set mCloseGadget = load_image_obj(g_MakePath(StylePath) & "gadgets\close.png")
'    If Not is_valid_image(mCloseGadget) Then _
'        Set mCloseGadget = bm_Close


'    Set mActionsGadget = load_image_obj(g_MakePath(StylePath) & "gadgets\actions.png")
'    If Not is_valid_image(mActionsGadget) Then _
        Set mActionsGadget = bm_Actions

    If mInfo.AckButtonLabel = "" Then _
        mInfo.AckButtonLabel = g_ButtonLabelFromAck(mInfo.DefaultAck)


    ' /* R2.5 Beta 2 */

Dim psc As TScriptCallback
Dim sz As String

    If (mInfo.ScriptFilename <> "") And (mInfo.ScriptLanguage <> "") Then
        Set psc = New TScriptCallback
        If psc.InitFrom(mInfo.ScriptFilename, mInfo.ScriptLanguage, sz) Then
            Set mScriptCallback = psc

        Else
            g_Debug "CSnarlWindow.Create(): error intialising callback script (" & sz & ")", LEMON_LEVEL_WARNING

        End If

    End If

Dim pbm As mfxBitmap
Dim rc As RECT
Dim i As Long

    m_Alpha = 255 * (g_SafeLong(g_ConfigGet("global_opacity")) / 100)
    mOriginalAlpha = m_Alpha
    mOriginalTimeout = Info.Timeout
    Set mInstance = Instance
    mStyleFlags = StyleFlags

    ' /* R2.2: styles can (now) also request click-through notifications */

    If (StyleFlags And S_STYLE_CLICK_THROUGH) Then _
        mClickThru = True

    ' /* ask our style instance to update its content */

    uUpdateStyleContent Info

Dim dwExStyle As Long

    dwExStyle = WS_EX_TOPMOST
    If mClickThru Then
        dwExStyle = dwExStyle Or WS_EX_TRANSPARENT
        If mOriginalTimeout = 0 Then
            g_Debug "CSnarlWindow.Create(): timeout cannot be zero if click-through", LEMON_LEVEL_WARNING
            mOriginalTimeout = IIf(Val(g_ConfigGet("default_duration")) > 0, Val(g_ConfigGet("default_duration")), 10)

        End If
    End If

    If nto_CreateWindow(theWindow, Me, N_POPUP_WINDOW, , , , dwExStyle) = M_OK Then
        uGetContent
        If NOTNULL(mContent) Then
            mPointerInside = False
            uRedraw 0
            theWindow.MakeVisible
            If (StyleFlags And S_STYLE_PULSE_NEEDED) Then _
                SetTimer theWindow.hWnd, TIMER_PULSE, 100, 0

        Else
            g_Debug "CSnarlWindow.Create(): returned style bitmap was NULL", LEMON_LEVEL_CRITICAL

        End If
    End If

End Function

Public Property Get Window() As NitroWindow

    Set Window = theWindow

End Property

Public Sub Show()
Dim bDone As Boolean
Dim pt As POINTAPI

    If Not (mInstance Is Nothing) Then
        If (mStyleFlags And S_STYLE_CUSTOM_SHOW) Then

            mInstance.Show True
            m_Alpha = 255
            mCurrent = 255

            Do
                bDone = True
                mInstance.AdjustPosition pt.x, pt.y, m_Alpha, bDone

                ' /* if S_STYLE_WILL_RESIZE flag is set, ask the instance for a new bitmap each time */

'                If (mStyleFlags And S_STYLE_WILL_RESIZE) Then
''                    Set theBitmap = uGetContent2()
''                    theView.SizeTo theBitmap.Width + OFFSET_ICON, theBitmap.Height + OFFSET_ICON
'
'                End If

                uRedraw m_Alpha, pt.x, pt.y
                DoEvents

                ' /* keep going until the instance leaves bDone as True */

            Loop While Not bDone

            Me.SetTimeout mOriginalTimeout
            mVisible = True
            Exit Sub

        End If
    End If

    If Not (theWindow Is Nothing) Then
        uSetTargetAlpha m_Alpha
        Me.SetTimeout mOriginalTimeout
'        If m_Timeout > 0 Then _
            theWindow.AddTimer TIMER_MAIN_TICK, m_Timeout * 1000

        mVisible = True

    End If

End Sub

Public Sub Quit()

    If mQuitting Then _
        Exit Sub

    Debug.Print "CSnarlWindow.Quit(): closing..."

    mVisible = False
    If Not (theWindow Is Nothing) Then
        mQuitting = True                                    '// set this notification as no longer valid...
        KillTimer theWindow.hWnd, TIMER_TRACK_POINTER
'        KillTimer theWindow.hWnd, TIMER_PULSE
        theWindow.AddExStyles WS_EX_TRANSPARENT             '// window becomes click-through while fading

        If Not (g_NotificationRoster Is Nothing) Then _
            g_NotificationRoster.Remove mInfo.Token         '// remove from roster so space can be reused

        uSetTargetAlpha 0                                   '// async fade to zero

    End If

End Sub

Public Sub Remove()

    Debug.Print "CSnarlWindow.Remove(): removing window..."

    mAlphaStep = 0

    If Not (theWindow Is Nothing) Then _
        KillTimer theWindow.hWnd, TIMER_ASYNC_FADE

    nto_DestroyWindow theWindow
'    Set theView = Nothing

        ' /* R2.4: if this was a low priority notification, check the queue */

'                        If mInfo.Priority < 0 Then _
            g_NotificationRoster.QueueNext

End Sub

Private Sub Class_Terminate()

    Debug.Print "CSnarlWindow.Terminate"

End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static fPressed As Boolean
Static bPaused As Boolean
Static pt As POINTAPI
Static t As Long

    Select Case uMsg
    Case WM_SETCURSOR
        SetCursor LoadCursor(0, IDC_ARROW)
        If LoWord(lParam) = HTERROR Then
            ' /* window is disabled which means a menu is active */
            Select Case HiWord(lParam)
            Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
                ' /* enable the window but don't process the message */
                EnableWindow hWnd, -1
                ReturnValue = -1
                MWndProcSink_WndProc = True

            End Select

        ElseIf Not mPointerInside Then
            If Not mClickThru Then _
                SetWindowPos hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE

            SetTimer hWnd, TIMER_TRACK_POINTER, 100, 0
            mPointerInside = True
            uRedraw

        End If


    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
        ' /* V38.131: better way of working - button pressed now does a SetCapture() and we only respond
        '    if the mouse button is released over the same item */
        mGadgetPressed = uHitTest(lParam)
        Debug.Print "--> " & mGadgetPressed
        Select Case mGadgetPressed
        Case "action", "open_button"
            uRedraw

        End Select
        fPressed = True
        SetCapture hWnd


    Case WM_LBUTTONUP
        If fPressed Then
            ReleaseCapture
            If mGadgetPressed = uHitTest(lParam) Then
                Select Case mGadgetPressed
                Case "close"
                    g_Debug "CSnarlWindow.WndProc(): close button clicked"
                    uDoClose bPaused

                Case "open_button"
                    g_Debug "CSnarlWindow.WndProc(): callback button clicked"
                    uDoInvoke bPaused

                Case "client"
                    g_Debug "CSnarlWindow.WndProc(): client area invoked"
                    If g_ConfigGet("callback_as_button") = "1" Then
                        ' /* new mode: close */
                        g_Debug "V44 mode: doing close..."
                        uDoClose bPaused

                    Else
                        ' /* old mode: invoke */
                        g_Debug "legacy mode: doing invoke..."
                        uDoInvoke bPaused

                    End If

                Case "actions_button"
                    bPaused = True
                    uDoActionsMenu
                    bPaused = False

                Case "action"
                    bPaused = True
                    uDoActionsMenu
                    bPaused = False

                End Select

            Else
                g_Debug "CSnarlWindow.WndProc(): WM_LBUTTONUP ignored: hit-test mismatch"

            End If
        Else
            g_Debug "CSnarlWindow.WndProc(): WM_LBUTTONUP ignored: outside window"

        End If

        If mGadgetPressed <> "" Then
            mGadgetPressed = ""
            uRedraw

        End If

'    Case WM_LBUTTONDBLCLK
'        ' /* R2.4 beta 4: make sticky if not already */
'        If mOriginalTimeout > 0 Then _
'            Me.SetTimeout 0


    Case WM_RBUTTONUP
        If fPressed Then
            ReleaseCapture
            If mGadgetPressed = uHitTest(lParam) Then
                Select Case mGadgetPressed
                Case "close"
                    g_Debug "CSnarlWindow.WndProc(): close button invoked (right-click)"
                    bPaused = True
                    uDoSysMenu
                    bPaused = False

                Case "client"
                    bPaused = True
                    If Not (mMenu Is Nothing) Then
                        ' /* track menu */
                        If Not uTrackMenu(hWnd) Then
                            bPaused = False
                            Exit Function

                        End If

                    Else
                        If (g_ConfigGet("allow_right_clicks") = "1") Or ((mInfo.IntFlags And SN_NF_API_MASK) < 42) Then
                            ' /* this is deprecated as of R2.4 DR7 */
                            uNotify SNARL_CALLBACK_R_CLICK

                        Else
                            ' /* R2.4.1: pop open the system menu instead */
                            uDoSysMenu

                        End If

                    End If

                    bPaused = False

                End Select

            Else
                g_Debug "CSnarlWindow.WndProc(): WM_RBUTTONUP ignored: hit-test mismatch"

            End If
        Else
            g_Debug "CSnarlWindow.WndProc(): WM_RBUTTONUP ignored: outside window"

        End If

        mGadgetPressed = ""


    Case WM_MBUTTONUP
        ' /* 38.133 - support the middle button */
        If fPressed Then
            ReleaseCapture

            If mGadgetPressed = uHitTest(lParam) Then
                Select Case mGadgetPressed
                Case "close"
    
                Case "client"

                    If (g_ConfigGet("allow_right_clicks") = "1") Or ((mInfo.IntFlags And SN_NF_API_MASK) < 42) Then
                        ' /* this is deprecated as of R2.4 DR7 */

                        bPaused = True
                        uNotify SNARL_CALLBACK_M_CLICK
                        bPaused = False

                    End If

'                    ' /* R2.4: hide when middle button clicked? */
'                    If (mInfo.Flags And SNARL41_NOTIFICATION_AUTO_DISMISS) Then _
'                        Hide

                End Select
            Else
                g_Debug "CSnarlWindow.WndProc(): WM_MBUTTONUP ignored: hit-test mismatch"

            End If
        Else
            g_Debug "CSnarlWindow.WndProc(): WM_MBUTTONUP ignored: outside window"

        End If
        mGadgetPressed = ""


    Case WM_TIMER

'        If (wParam = TIMER_MAIN_TICK) Then _
            Debug.Print (bPaused) & " " & (mPointerInside) & " " & (mFuzzy) & " " & CStr(mTick) & " " & g_SafeLeftStr(mInfo.Text, 20)

        If (wParam = TIMER_MAIN_TICK) And (Not bPaused) And (Not mPointerInside) And (Not mFuzzy) Then
            mTick = mTick - 1
'            Debug.Print CStr(mTick) & ": " & mInfo.Text
            If mTick = 0 Then
                bPaused = True
                g_Debug "CSnarlWindow.WndProc(): TIMED_OUT (token=" & CStr(mInfo.Token) & ")", LEMON_LEVEL_INFO
                uNotify SNARL_CALLBACK_TIMED_OUT
                bPaused = False
                Hide

            End If

        ElseIf wParam = TIMER_ASYNC_FADE Then

'            Debug.Print "target: " & mTargetAlpha & " current: " & mCurrent

            If mAlphaStep < 0 Then
                ' /* fading down */
                If mCurrent <= mTargetAlpha Then
                    ' /* done! */
                    mCurrent = mTargetAlpha
                    KillTimer hWnd, wParam
                    mAlphaStep = 0
                    
                    If mQuitting Then _
                        Remove

                    Exit Function

                End If

            Else
                ' /* fading up */
                If mCurrent >= mTargetAlpha Then
                    ' /* done! */
'                    Debug.Print "*** Fading took: " & GetTickCount() - mFadeStart & " ms ***"
                    mCurrent = mTargetAlpha
                    KillTimer hWnd, wParam
                    apply_view_to_window mView, hWnd, mCurrent        ' // re-apply in case we've overshot
                    mAlphaStep = 0
                    Exit Function

                End If
            End If

'            Debug.Print GetTickCount() - t
'            t = GetTickCount()

            ' /* keep fading */
            mCurrent = mCurrent + mAlphaStep
            apply_view_to_window mView, hWnd, mCurrent

        ElseIf (wParam = TIMER_TRACK_POINTER) And (Not fPressed) And (Not mMenuOpen) Then

            If WindowFromPoint(GET_X_LPARAM(GetMessagePos()), GET_Y_LPARAM(GetMessagePos())) <> hWnd Then
                KillTimer hWnd, wParam
                mPointerInside = False
                uRedraw

            End If
            

'            GetCursorPos pt
'            If Not theWindow.Frame.Contains(nto_NewPoint(pt.x, pt.y)) Then
'                KillTimer hWnd, wParam
'                mPointerInside = False
'                uRedraw
'
'            End If

        ElseIf wParam = TIMER_PULSE Then
            uPulse

        End If


'    Case WM_NCHITTEST
'
'        Debug.Print "NCHITTEST"
'
'        mPointerInside = True
'        uRedraw
'        SetTimer hWnd, TIMER_TRACK_POINTER, 100, 0


    Case WM_CAPTURECHANGED
        fPressed = False

    Case WM_NCACTIVATE
        If wParam = 0 Then _
            EnableWindow hWnd, -1

    Case WM_ENTERMENULOOP, WM_EXITMENULOOP
        mMenuOpen = (uMsg = WM_ENTERMENULOOP)

    End Select

End Function

Private Function uHitTest(ByVal lParam As Long) As String
Dim pp As BPoint
Dim s As Single
Dim n As Long

    ' /* returns:
    '       "close" if the pointer is over the close button
    '       "action" if the pointer is over the action button (and actions are defined)
    '       "client" if the pointer is over the notification
    '       "" if the pointer is outside the notification
    ' */

    Set pp = new_BPointFromInt32(lParam)
    If uWantsDropshadow(n) Then _
        pp.OffsetBy -n, -n

    s = Val(g_ConfigGet("scaling"))
    If s <= 0# Then _
        s = 1#                  ' // must NOT be zero!

    pp.x = pp.x / s
    pp.y = pp.y / s
    
    If BW_Bounds(theWindow.hWnd).Contains(pp) Then
        If pp.IsInRect(mCallbackRect) Then
            uHitTest = "open_button"
    
        ElseIf pp.IsInRect(mActionsRect) Then
            uHitTest = "action"
            
        ElseIf pp.IsInRect(mCloseRect) Then
            uHitTest = "close"
    
        Else
            uHitTest = "client"
    
        End If

    Else
        g_Debug "CSnarlWindow.uHitTest(): not inside window"

    End If

End Function

Public Sub Hide()
Dim bDone As Boolean
Dim pt As POINTAPI

    If Not (mInstance Is Nothing) Then
        If (mStyleFlags And S_STYLE_CUSTOM_HIDE) Then
            mVisible = False
            mQuitting = True

            With theWindow
                .RemoveTimer TIMER_MAIN_TICK
                .RemoveTimer TIMER_TRACK_POINTER
                .AddExStyles WS_EX_TRANSPARENT          '// window becomes click-through while hiding

            End With

            g_NotificationRoster.Remove mInfo.Token     '// remove from roster so space can be reused
            mInstance.Show False

            Do
                bDone = True
                mInstance.AdjustPosition pt.x, pt.y, m_Alpha, bDone

'                If (mStyleFlags And S_STYLE_WILL_RESIZE) Then
''                    Set theBitmap = uGetContent2()
''                    theView.SizeTo theBitmap.Width + OFFSET_ICON, theBitmap.Height + OFFSET_ICON
'
'                End If
                uRedraw m_Alpha, pt.x, pt.y
                DoEvents

            Loop While Not bDone

            Remove
            Exit Sub
        Else
            Debug.Print "CSnarlWindow.Hide(): no custom hide required"

        End If
    Else
        Debug.Print "CSnarlWindow.Hide(): no instance"

    End If

    theWindow.RemoveTimer TIMER_MAIN_TICK
    Quit

End Sub

Public Property Get Id() As Long

    Id = mInfo.Token

End Property

Public Sub uSetTargetAlpha(ByVal Alpha As Long, Optional Fast As Boolean = False)

    Debug.Print "CSnarlWindow.uSetTargetAlpha(): current=" & mCurrent & " requested=" & Alpha

    If (Alpha = 0) And (mCurrent = 0) And (mQuitting) Then
        ' /* R2.4 DR8: handles an unusual case whereby the notification is being
        '    removed before async fading has even begun.  */
        Remove
        Exit Sub

    ElseIf Alpha = mCurrent Then
        ' /* nothing to do */
        Exit Sub

    End If

    ' /* figure out granularity */

Dim s As Single

    s = Abs(mCurrent - Alpha)

'    Debug.Print "alpha delta = " & s

    If Fast Then
        s = (s / 150) * (10 + 6)
    
    Else
        s = (s / 300) * (10 + 6)

    End If

    If Alpha < mCurrent Then _
        s = -s

'    Debug.Print "step size = " & s

    mAlphaStep = CLng(s)

    ' /* kick off the fade process */

    mTargetAlpha = Alpha
    SetTimer theWindow.hWnd, TIMER_ASYNC_FADE, 10, 0
    mFadeStart = GetTickCount()

End Sub

Public Sub Enable()

    m_Disable = m_Disable - 1

End Sub

Public Sub Disable()

    m_Disable = m_Disable + 1

End Sub

Public Function SetTimeout(ByVal Timeout As Long) As M_RESULT

    SetTimeout = M_FAILED
    If (theWindow Is Nothing) Then _
        Exit Function

    SetTimeout = M_INVALID_ARGS
    If (Timeout < 0) Or (Timeout > 65535) Then _
        Exit Function

    SetTimeout = M_OK
    theWindow.RemoveTimer TIMER_MAIN_TICK
    mOriginalTimeout = Timeout

    If Timeout = 0 Then
        uQuickUpdate
        Exit Function

    End If

'    mTick = Timeout * 10                    ' // timer frequency is 100ms...
'    theWindow.AddTimer TIMER_MAIN_TICK, 100

    mTick = Timeout
    theWindow.AddTimer TIMER_MAIN_TICK, 1000

End Function

Public Property Get IsVisible() As Boolean

    IsVisible = mVisible

End Property

Public Function GetView() As mfxView

    Set GetView = mView

End Function

Private Sub uRedraw(Optional ByVal Alpha As Long = -1, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)

    If (ISNULL(mContent)) Or (ISNULL(theWindow)) Then _
        Exit Sub

    If Alpha = -1 Then _
        Alpha = mCurrent

Dim n As Long

    If (x = -1) And (y = -1) Then
'            Set pr = theWindow.Frame
'            x = pr.Left - 10
'            y = pr.Top - 10

    ElseIf uWantsDropshadow(n) Then
        ' /* specific position requested and dropshadow enabled so compensate */
        x = x - n
        y = y - n

    End If

    Set mActionsRect = Nothing
    Set mCloseRect = Nothing
    Set mCallbackRect = Nothing

    ' /* build the view up */

Dim pt As BPoint
Dim pv As mfxView
Dim pb As mfxBitmap

Const EMBLEM_ALPHA = 230
Const EMBLEM_SIZE = 14

    Set mView = New mfxView
    With mView
        ' /* style-provided content */
        .SizeTo mContent.Width, mContent.Height
        Set pb = mContent.Duplicate
'        If (mPointerInside) And (Not mClickThru) Then _
            pb.ChangeBrightness -0.05

        .DrawScaledImage pb

        ' /* set top-right corner */
        Set pt = new_BPoint(.Width - EMBLEM_SIZE - 4, 4)

        ' /* emblems */

        If mOriginalTimeout = 0 Then
            ' /* sticky */
            .DrawScaledImage bm_IsSticky, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), EMBLEM_ALPHA
            pt.OffsetBy -(EMBLEM_SIZE + 1), 0

        End If

        ' /* indicate notification has a callback - only if using legacy mode, it's interactive and there is a callback */

        If (g_ConfigGet("callback_as_button") = "0") And (Not mClickThru) And (mInfo.DefaultAck <> "") Then
            .DrawScaledImage bm_HasActions, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), EMBLEM_ALPHA
            pt.OffsetBy -(EMBLEM_SIZE + 1), 0

        End If

        ' /* forwarded/remote - either/or */

        If (mInfo.IntFlags And SN_NF_FORWARD) Then
            ' /* if the notification was forwarded */
            .DrawScaledImage bm_Forward, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), EMBLEM_ALPHA
            pt.OffsetBy -(EMBLEM_SIZE + 1), 0

        ElseIf (mInfo.IntFlags And SN_NF_REMOTE) Then
            ' /* if the notification was generated by a remote app */
            .DrawScaledImage bm_Remote, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), EMBLEM_ALPHA
            pt.OffsetBy -(EMBLEM_SIZE + 1), 0

        End If

Dim pr As BRect

        If (mPointerInside) And (Not mClickThru) Then
            ' /* HUD */
            If (g_ConfigGet("show_timestamp") = "1") And (.Width > 79) Then _
                uDrawHUDContent

            Set mActionsRect = new_BRect(0, 0, (24 - 1), (24 - 1)).OffsetByCopy(4, 4)
'            Set mActionsRect = new_BRect(0, 0, (24 - 1), (24 - 1)).OffsetByCopy(.Width - (24 + 4), .Height - (24 + 4))

            ' /* draw the close gadget if running in legacy mode */
            If g_ConfigGet("callback_as_button") = "0" Then
                Set mCloseRect = new_BRect(0, 0, (24 - 1), (24 - 1)).OffsetByCopy(4, 4)
                mActionsRect.OffsetBy 0, 24 + 2

                Set pb = bm_Button.Duplicate
                If mGadgetPressed = "close" Then _
                    pb.ChangeBrightness -0.2

                .DrawScaledImage pb, mCloseRect.TopLeft, mCloseRect.Size
                .DrawScaledImage bm_CloseGadget, mCloseRect.TopLeft.OffsetByCopy(Fix((mCloseRect.Width - 16) / 2), Fix((mCloseRect.Height - 16) / 2)), new_BPoint(16, 16)

            End If

            ' /* actions gadget */
            If mInfo.Actions.CountItems > 0 Then
'                Set pb = create_bitmap_from_image(IIf(g_ConfigGet("use_style_icons"), mActionsGadget, bm_Actions))
                Set pb = bm_Button.Duplicate
                If (mGadgetPressed = "action") Or (NOTNULL(theActionsPopup)) Then _
                    pb.ChangeBrightness -0.2

                .DrawScaledImage pb, mActionsRect.TopLeft, mActionsRect.Size
                .DrawScaledImage bm_ActionsGadget, mActionsRect.TopLeft.OffsetByCopy(Fix((mActionsRect.Width - 16) / 2), Fix((mActionsRect.Height - 16) / 2)), new_BPoint(16, 16)

'                If .Height > (mActionsButton.Height + 4 + 4) Then
'                    ' /* gets the big button */
'                    Set pb = mActionsButton.Duplicate
'                    If mGadgetPressed = "action" Then _
'                        pb.ChangeBrightness -0.15
'
'                    .DrawScaledImage pb, new_BPoint(.Width - mActionsButton.Width - 4, 4)
'                    Set mActionsRect = pr.Duplicate
'                    pr.OffsetBy 0, pr.Height + 2
'
'                ElseIf .Height > 47 Then
'                    ' /* below close gadget */
''                    Set mActionsRect = mCloseRect.OffsetByCopy(0, 20 * s)
''                    .DrawScaledImage IIf(g_ConfigGet("use_style_icons"), mActionsGadget, bm_Actions), new_BPoint(0, 20), new_BPoint(20, 20)
'
'                Else
'                    ' /* alongside close gadget */
''                    Set mActionsRect = mCloseRect.OffsetByCopy(20 * s, 0)
''                    .DrawScaledImage IIf(g_ConfigGet("use_style_icons"), mActionsGadget, bm_Actions), new_BPoint(20, 0), new_BPoint(20, 20)
'
'                End If

            Else
                ' /* no actions: no need for the rect */
                Set mActionsRect = Nothing

            End If

            ' /* callback button */
            If (g_ConfigGet("callback_as_button") = "1") And (uCanDoCallback()) Then
                Set pr = new_BRect(0, 0, bm_CallbackButton.Width - 1, bm_CallbackButton.Height - 1)
                pr.OffsetBy .Width - pr.Width - 4, .Height - pr.Height - 4
                
'                If NOTNULL(mActionsRect) Then _
                    pr.OffsetBy -(24 + 2), 0

                Set pb = bm_CallbackButton.Duplicate
                If mGadgetPressed = "open_button" Then _
                    pb.ChangeBrightness -0.2

                .DrawScaledImage pb, pr.TopLeft
                Set mCallbackRect = pr.Duplicate
                pr.InsetBy 4, 0

                .SetFont "Arial", 8 ', True
                .SetHighColour rgba(0, 0, 0)
                .SetLowColour rgba(0, 0, 0, 16)
                .DrawString .GetFormattedText(mInfo.AckButtonLabel, pr.Width), pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER  ' Or MFX_SIMPLE_OUTLINE

            End If

            ' /* buttons */
'            If (mInfo.DefaultAck <> "") And (mInfo.Actions.CountItems > 0) Then
'                ' /* both */
'                If .Height > ((mActionsButton.Height * 2) + 4 + 4 + 4) Then
'                    .DrawScaledImage mOpenButton, new_BPoint(4, 4)      '//   new_BPoint(.Width - mActionsButton.Width - 4, 4)
'                    .DrawScaledImage mActionsButton, new_BPoint(4, 4 + mActionsButton.Height + 4)  '// new_BPoint(.Width - mActionsButton.Width - 4, 4 + mActionsButton.Height + 4)
'                    Set mCallbackRect = pr.Duplicate
'                    pr.OffsetBy 0, pr.Height + 4
'                    Set mActionsButtonRect = pr.Duplicate
'
'                End If
'
'            ElseIf mInfo.DefaultAck <> "" Then
'                ' /* open only */
'                If .Height > (mActionsButton.Height + 4 + 4) Then
'                    .DrawScaledImage mOpenButton, new_BPoint(4, 4)      '//   new_BPoint(.Width - mActionsButton.Width - 4, 4)
'                    Set mCallbackRect = pr.Duplicate
'
'                End If
'
'            ElseIf mInfo.Actions.CountItems > 0 Then
'                ' /* actions only */
'                If .Height > (mActionsButton.Height + 4 + 4) Then
'                    .DrawScaledImage mActionsButton, new_BPoint(4, 4)      '//   new_BPoint(.Width - mActionsButton.Width - 4, 4)
'                    Set mActionsButtonRect = pr.Duplicate
'
'                End If
'            End If

        End If

        ' /* final content */

        Set pb = .ConvertToBitmap()

    End With

    ' /* scale */

Dim s As Single

    s = Val(g_ConfigGet("scaling"))
    If s <= 0# Then _
        s = 1#                  ' // must NOT be zero!

    With mView
        .SizeTo ROUNDUP(pb.Width * s), ROUNDUP(pb.Height * s)
        .Clear
        .DrawScaledImage pb, , new_BPoint(.Width, .Height)


'        .SetHighColour rgba(255, 255, 255)
'        .SetLowColour rgba(0, 128, 255)
'        .DrawString CStr(mPointerInside), .Bounds, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER Or MFX_SIMPLE_OUTLINE

    End With

    ' /* draw with or without dropshadow */

    If uWantsDropshadow(n) Then
        ' /* has a drop shadow */
'        If (mPointerInside) And (Not mClickThru) Then
'            Set pb = create_dropshadow(mView, , n * 2, 1)
'
'        Else
            Set pb = create_dropshadow(mView, , n * 2, Fix((Val(g_ConfigGet("dropshadow_strength")) / 100) * 255))
'
'        End If
        ' /* re-create using new image */
        With mView
            .SizeTo pb.Width, pb.Height
            .Clear
            .DrawScaledImage pb

        End With
    End If

    apply_view_to_window mView, theWindow.hWnd, Alpha, x, y

End Sub

Friend Sub bSetAlpha(ByVal Alpha As Integer)

    If Alpha < 0 Then
        Alpha = 0

    ElseIf Alpha > 255 Then
        Alpha = 255

    End If

    mCurrent = Alpha
    uRedraw

End Sub

Private Sub uDoSysMenu()
Dim pmi As OMMenuItem
Dim pm As OMMenu

    With New OMMenu
        ' /* R2.2 */
        .AddItem .CreateItem("stky", "Make sticky", , (mOriginalTimeout > 0))
        .AddItem .CreateItem(" ", "Details", , , , , , uDetailsMenu())          ' // R2.3
        .AddSeparator
        .AddItem .CreateItem("acfg", mInfo.ClassObj.App.Name & " settings...", , mInfo.ClassObj.App.HasConfig)
        .AddItem .CreateItem("ncfg", "Notification settings...")
'        .AddItem .CreateItem("$" & mInfo.ClassObj.App.Name & "#?" & mInfo.ClassObj.Name, "Notification settings...")
        .AddSeparator
        Set pm = uGetRedirectorsMenu()
        .AddItem .CreateItem(" ", "Redirect using", , (pm.CountItems > 0), , , , pm)
        Set pm = uGetForwardersMenu()
        .AddItem .CreateItem(" ", "Forward to", , (pm.CountItems > 0), , , , pm)
        .AddItem .CreateItem("wlog", "Write content to log")
        .AddItem .CreateItem("copy", "Copy content to Clipboard")
        .AddSeparator
        .AddItem .CreateItem("blck", "Ignore further notifications")
        .AddSeparator
'        .AddSeparator
        .AddItem .CreateItem("capp", "Close all from " & mInfo.ClassObj.App.Name)
        .AddItem .CreateItem("call", "Close all")
        .AddItem .CreateItem("clse", "Close")

        If g_IsPressed(VK_CONTROL) Then
            .AddSeparator
            .AddItem .CreateItem("scrn", "Capture notification")

        End If

        EnableWindow theWindow.hWnd, 0

        Set pmi = .Track(theWindow.hWnd)

    End With

Dim szData As String
Dim sz() As String

    If Not (pmi Is Nothing) Then
        EnableWindow theWindow.hWnd, -1         ' // item was selected so re-enable the window
        
        Select Case pmi.Name
        Case "clse"
            Hide

        Case "blck"
            mInfo.ClassObj.SetEnabled False

        Case "capp"
            g_NotificationRoster.CloseMultiple mInfo.ClassObj.App.Token

        Case "call"
            g_NotificationRoster.CloseMultiple 0

        Case "scrn"
            uCapture

        Case "stky"
            Me.SetTimeout 0

        Case "wlog"
            g_WriteToLog mInfo.Title, mInfo.Text

        Case "copy"
            Clipboard.Clear
            Clipboard.SetText mInfo.Title & vbCrLf & mInfo.Text

        Case "acfg"
            mInfo.ClassObj.App.DoSettings 0

        Case "ncfg"
            frmAbout.ShowClassConfigPanel 0, mInfo.ClassObj.App, mInfo.ClassObj.Name


        Case Else
            szData = g_SafeRightStr(pmi.Name, Len(pmi.Name) - 1)
            
            Select Case g_SafeLeftStr(pmi.Name, 1)
'            Case "$"
'                ' /* notification settings */
'                sz = Split(szData, "#?")
'                frmAbout.DoAppConfig sz(0), sz(1)

            Case ">"
                ' /* forward */
                g_SubsRoster.QuickForward szData, mInfo
                
            Case "]"
                ' /* redirect */
                g_StyleRoster.RedirectTo szData, mInfo

            End Select
        End Select
    End If

End Sub

Public Function SetAck(ByVal Value As String) As M_RESULT

    mInfo.DefaultAck = Value
    SetAck = M_OK

End Function

Private Sub uCapture()
Dim szPath As String
Dim sz As String
Dim c As Long

    If Not g_GetSystemFolder(CSIDL_DESKTOP, szPath) Then _
        Exit Sub

    szPath = g_MakePath(szPath)

    c = 1
    sz = szPath & "snarl_" & CStr(c) & ".png"
    Do While g_Exists(sz)
        c = c + 1
        If c > 65535 Then _
            Exit Sub

        sz = szPath & "snarl_" & CStr(c) & ".png"

    Loop

    mPointerInside = False
    uRedraw

    CaptureToFile sz

    mPointerInside = True
    uRedraw

End Sub

Public Function CaptureToFile(ByVal Filename As String)

    If Not (mView Is Nothing) Then _
        mView.WriteToFile Filename, "image/png"

End Function

Public Function SetMenu(ByVal Value As String) As M_RESULT

    ' /* 'Value' should be a string formatted as <Label>[#?<value>][|<Label>[#?<value>]]...
    '    <Label> defines the menu item label and can be any text string.  <Value> must be numeric
    '    and in the range of 1 to 65535.  <Value> does not need to be specified, in which
    '    case the item is added but no notification is sent if it is selected.  If <Label> is
    '    blank, a separator item is added (which cannot be selected) */

    ' /* R2.3 onwards: if <Label> is prefixed with '$' the following character is used as
    '    a control character.  The following are currently defined:
    '       'X' a checkmark is displayed against the menu item
    '       '\' the menu item is ghosted and cannot be selected
    '       'x'  a checkmark is displayed against the menu item and the menu item is ghosted and cannot be selected
    ' */

    On Error Resume Next

    Set mMenu = Nothing
    SetMenu = M_OK

    If Value = "" Then _
        Exit Function

Dim sz() As String
Dim i As Long

    sz() = Split(Value, "|")
    Debug.Print "$$%: " & UBound(sz())

    If UBound(sz()) = -1 Then _
        Exit Function

    Set mMenu = New OMMenu

    For i = 0 To UBound(sz())
        uAddToMenu mMenu, sz(i)

    Next i

End Function

Private Sub uAddToMenu(ByRef Menu As OMMenu, ByVal Item As String)
Dim szText As String
Dim szData As String
Dim i As Long

    i = InStr(Item, "#?")
    If i = 0 Then
        szText = Item

    Else
        szText = g_SafeLeftStr(Item, i - 1)
        szData = g_SafeRightStr(Item, Len(Item) - i - 1)

    End If

Dim itemEnabled As Boolean
Dim itemMarked As Boolean

    ' /* pre-set this */

    itemEnabled = True

    If szText = "" Then
        ' /* R2.3 seperator added even if <value> supplied with empty <Label> */
        Menu.AddSeparator
        Exit Sub

    ElseIf g_SafeLeftStr(szText, 1) = "$" Then
        ' /* control character */

        Select Case g_SafeMidStr(szText, 2, 1)
        Case "X"
            ' /* checkmark */
            itemMarked = True

        Case "\"
            ' /* disabled */
            itemEnabled = False

        Case "x"
            ' /* disabled and marked */
            itemMarked = True
            itemEnabled = False

        Case Else
            g_Debug "CSnarlWindow.uAddToMenu(): invalid control character '" & g_SafeMidStr(szText, 2, 1) & "'", LEMON_LEVEL_CRITICAL
            Exit Sub

        End Select

        szText = g_SafeRightStr(szText, Len(szText) - 2)

    End If

Dim dw As Long

    ' /* if <Value> is provided, make sure it's valid */

'    If szData <> "" Then
'        If Not g_IsNumeric(szData) Then
'            g_Debug "CSnarlWindow.uAddToMenu(): can't add '" & Item & "' to menu: data part is not numeric", LEMON_LEVEL_CRITICAL
'            Exit Sub
'
'        End If
'
'        dw = Val(szData)
'        If (dw < 1) Or (dw > 65535) Then
'            g_Debug "CSnarlWindow.uAddToMenu(): can't add '" & Item & "' to menu: data part must be 1-65535", LEMON_LEVEL_CRITICAL
'            Exit Sub
'
'        End If
'
'    End If

    Menu.AddItem mMenu.CreateItem(szData, szText, , itemEnabled, itemMarked)

End Sub

Private Function uTrackMenu(ByVal hWndOwner As Long) As Boolean

    If (mMenu Is Nothing) Then _
        Exit Function

Dim pmi As OMMenuItem

    EnableWindow hWndOwner, 0
    Set pmi = mMenu.Track(hWndOwner)
    If (pmi Is Nothing) Then _
        Exit Function

    EnableWindow hWndOwner, -1          ' // item was selected so re-enable the window

    If g_SafeLeftStr(pmi.Name, 1) = "!" Then
        ' /* R2.4: menu id's can be bang commands, paths, urls, etc. */
        g_ProcessAck pmi.Name

    Else
        ' /* send it as a SNARL_NOTIFICATION_MENU notification */
        uNotify SNARL_CALLBACK_MENU_SELECTED, pmi.Name

    End If

    uTrackMenu = True

End Function

Private Sub uNotify(ByVal Code As Integer, Optional ByVal Data As String)
Dim szResponse As String
Dim szData As String
Dim n As Integer

    If Not (mInfo.Socket Is Nothing) Then
        ' /* if we have a socket, then reply to that */

        If Code <> SNARL_NOTIFY_ACTION Then _
            Code = Code + 270               ' // convert into "proper" V42 notification

        If (mInfo.IntFlags And SN_NF_IS_SNP3) Then
            ' /* R2.4.2 DR3: SNP3 handling */

            Select Case Code
            Case SNARL_NOTIFY_INVOKED
                szData = "WasClicked"

            Case SNARL_NOTIFY_EXPIRED
                szData = "TimedOut"

            Case SNARL_NOTIFY_CLOSED
                szData = "WasClosed"

            Case SNARL_NOTIFY_ACTION
                szData = "ActionSelected"
                szResponse = "event-data: " & g_SafeRightStr(Data, Len(Data) - 1) & vbCrLf

            Case Else
                g_Debug "CSnarlWindow.uNotify(): SNP3 does not support event code " & CStr(Code), LEMON_LEVEL_INFO
                Exit Sub

            End Select

            g_SNP3SendCallback mInfo.Socket, Code, szData, szResponse, uGetSNP3Info()


        ElseIf (mInfo.IntFlags And SN_NF_IS_GNTP) Then
            ' /* R2.4.2: GNTP handling */

            Select Case Code
            Case SNARL_NOTIFY_INVOKED
                szData = "CLICKED"

            Case SNARL_NOTIFY_EXPIRED
                szData = "TIMEDOUT"

            Case SNARL_NOTIFY_CLOSED
                szData = "CLOSED"

            Case Else
                g_Debug "CSnarlWindow.uNotify(): GNTP does not support notification " & CStr(Code)
                Exit Sub

            End Select

            If gntp_CreateCallbackResponse(szData, mInfo.ClassObj.App.Name, mInfo.OriginalContent, szResponse) Then
                g_Debug "CSnarlWindow.uNotify(): sending GNTP callback response to " & mInfo.Socket.RemoteHostIP & ":" & CStr(mInfo.Socket.RemotePort)
                mInfo.Socket.SendData szResponse

            Else
                g_Debug "CSnarlWindow.uNotify(): failed to create GNTP response", LEMON_LEVEL_CRITICAL

            End If

            ' /* close the socket now, as per the GNTP specification */

            mInfo.Socket.CloseSocket

        Else
            ' /* SNP2 */

            ' /* most callbacks are deprecated as of V42 - INVOKED, EXPIRED and ACTION are the only
            '    ones which aren't, so we only do these for SNP */

            Select Case Code
            Case SNARL_NOTIFY_INVOKED, SNARL_NOTIFY_EXPIRED
                ' /* nothing special to do here */

            Case SNARL_NOTIFY_ACTION, SNARL_NOTIFY_MENU
                ' /* send the data verbatim */
                szData = "/" & Data

            Case Else
                g_Debug "CSnarlWindow.uNotify(): not sending deprecated command " & CStr(Code)
                Exit Sub

            End Select

            g_Debug "CSnarlWindow.uNotify(): dest=" & mInfo.Socket.RemoteHostIP & ":" & CStr(mInfo.Socket.RemotePort)

            mInfo.Socket.SendData "SNP/" & SNP_VERSION & "/" & CStr(Code) & _
                                  "/" & uSNP2CallbackResponse(Code) & _
                                  "/" & IIf(mInfo.CustomUID = "", CStr(mInfo.Token), mInfo.CustomUID) & szData & _
                                  vbCrLf

        End If

    ElseIf Not (mScriptCallback Is Nothing) Then
        ' /* R2.5 Beta 2 */

        Select Case Code
        Case SNARL_NOTIFY_INVOKED, SNARL_CALLBACK_INVOKED
            mScriptCallback.DoClickedCallback mInfo.CustomUID
        
        Case SNARL_NOTIFY_EXPIRED, SNARL_CALLBACK_TIMED_OUT
            mScriptCallback.DoTimedOutCallback mInfo.CustomUID

        Case SNARL_NOTIFY_ACTION, SNARL_NOTIFY_MENU
            ' /* send the data verbatim */
            mScriptCallback.DoActionCallback mInfo.CustomUID, Data

        Case Else
            g_Debug "CSnarlWindow.uNotify(): not sending deprecated command " & CStr(Code)

        End Select

    ElseIf (IsWindow(mInfo.hWndReply) <> 0) And (mInfo.uReplyMsg <> 0) Then
        ' /* otherwise, if we have a valid window and message, reply to that */

        n = LoWord(g_SafeLong(g_SafeRightStr(Data, Len(Data) - 1)))
        g_Debug "CSnarlWindow.uNotify(): sending " & CStr(Code) & " to 0x" & g_HexStr(mInfo.hWndReply) & _
                " using 0x" & g_HexStr(mInfo.uReplyMsg, 4) & " token=" & CStr(mInfo.Token) & " data=" & CStr(n), LEMON_LEVEL_INFO

        PostMessage mInfo.hWndReply, mInfo.uReplyMsg, MAKELONG(Code, n), ByVal mInfo.Token

    Else
        g_Debug "CSnarlWindow.uNotify(): no reply socket, script or window"

    End If

End Sub

Private Function uSNP2CallbackResponse(ByVal Id As Long) As String
Dim sz As String

    Select Case Id

    Case SNARL_NOTIFY_INVOKED                   '// notification was left-clicked
        sz = "INVOKED"

    Case SNARL_NOTIFY_ACTION                    '// user picked an action from the pop-up
        sz = "ACTION"

    Case SNARL_NOTIFY_EXPIRED
        sz = "EXPIRED"

    Case Else
        sz = "(undefined)"

    End Select

    uSNP2CallbackResponse = sz

End Function

Private Sub uQuickUpdate()

    uGetContent
    uRedraw

End Sub

' /*---------------------------------------------------------------------------------------
'   Update() -- intelligent update
'
'   This function updates the existing notification's content, then asks the style instance
'   for an updated position (the style instance can resize the content during the update
'   request) and re-positions it on screen at the requested location and size.
'
' ---------------------------------------------------------------------------------------*/

Friend Function Update(ByRef Info As T_NOTIFICATION_INFO) As Boolean

    ' /* ask our style instance to update */

    uUpdateStyleContent Info
    uGetContent

    ' /* ask the instance if it wants to change the notification's position */

Dim pt As POINTAPI

    pt.x = theWindow.Frame.Left
    pt.y = theWindow.Frame.Top
    Me.AdjustPosition pt.x, pt.y
    g_MoveWindow theWindow.hWnd, pt.x, pt.y

    ' /* reset timeout */

    SetTimeout mOriginalTimeout

    ' /* redraw */

    uRedraw
    Update = True

End Function

Public Function IsNonInteractive() As Boolean

    IsNonInteractive = mClickThru

End Function

Public Sub MakeFuzzy(ByVal Enabled As Boolean)

    If (mFuzzy = Enabled) Or (Not mVisible) Then _
        Exit Sub

Static bmOld As mfxBitmap

    mFuzzy = Enabled

    If mFuzzy Then
        Set bmOld = mContent.Duplicate
        mContent.MakeGreyscale
        uRedraw

        ' /* R2.4: bring window to front - required? */

        SetWindowPos theWindow.hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
        uSetTargetAlpha Min(64, Fix(mOriginalAlpha / 2)), True

    Else

        Set mContent = bmOld.Duplicate
        Set bmOld = Nothing
        uRedraw

        Debug.Print "CSnarlWindow.MakeFuzzy(): restoring to alpha " & mOriginalAlpha
        uSetTargetAlpha mOriginalAlpha, True

    End If

End Sub

Private Sub uGetContent()

    On Error Resume Next

    If Not (mInstance Is Nothing) Then _
        Set mContent = create_bitmap_from_image(mInstance.GetContent())

    If (mContent Is Nothing) Then
        ' /* style didn't provide content so we switch to the built-in one */
        g_Debug "CSnarlWindow.uGetContent(): selected style didn't provide usable content", LEMON_LEVEL_WARNING
        Set mInstance = New TInternalStyle
        mStyleFlags = S_STYLE_MULTI_INSTANCE
        uUpdateInstance
        Set mContent = create_bitmap_from_image(mInstance.GetContent())

    End If

End Sub

Private Sub uUpdateStyleContent(ByRef Info As T_NOTIFICATION_INFO)

    On Error GoTo er

    If (mInstance Is Nothing) Then _
        Set mInstance = New TInternalStyle


    With Info

        ' /* translate special markers */

'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)

        ' /* take into consideration any "do not change" markers */

        If .Title = Chr$(255) Then _
            .Title = mInfo.Title

        If .Text = Chr$(255) Then _
            .Text = mInfo.Text

        If .IconPath = Chr$(255) Then _
            .IconPath = mInfo.IconPath

        ' /* this actually only contains the scheme name */

        If .SchemeName = Chr$(255) Then _
            .SchemeName = mInfo.SchemeName

    End With

    ' /* store the new values */

    With mInfo
        .Title = Info.Title
        .Text = Info.Text
        .IconPath = Info.IconPath
        .SchemeName = Info.SchemeName
        .LastUpdated = Now()            ' // update this as well
        .OriginalContent = Info.OriginalContent

    End With

    Debug.Print "updatestylecontent: " & mInfo.Text

    uUpdateInstance
    Exit Sub

er:
    g_Debug "CSnarlWindow.uUpdateStyleContent(): " & err.Description, LEMON_LEVEL_CRITICAL

End Sub

Private Sub uUpdateInstance()

    If (mInstance Is Nothing) Then _
        Exit Sub

    ' /* translate into a notification_info struct that the style type lib can understand */

Dim pInfo As notification_info

    With pInfo
'        .Title = Replace$(Info.Title, Chr$(255), "")
'        .Text = Replace$(Info.Text, Chr$(255), "")
'        .Icon = Replace$(Info.IconPath, Chr$(255), "")
'        .Scheme = Replace$(LCase$(Info.SchemeName), Chr$(255), "")

        If (mStyleFlags And S_STYLE_V42_CONTENT) Then
            .Title = mInfo.ClassObj.App.Name
            .Text = mInfo.OriginalContent

        Else
            .Title = mInfo.Title
            .Text = mInfo.Text

        End If

        .Icon = g_TranslateIconPath(mInfo.IconPath, mStylePath)
        .Scheme = LCase$(mInfo.SchemeName)

        If mInfo.Priority > 0 Then _
            .Flags = .Flags Or S_NOTIFICATION_IS_PRIORITY

    End With

    mInstance.UpdateContent pInfo

End Sub

Public Sub AdjustPosition(ByRef x As Long, ByRef y As Long)
Dim n As Long

    If uWantsDropshadow(n) Then
        x = x + n
        y = y + n

    End If

    mInstance.AdjustPosition x, y, 255, True

    ' /* compensate for dropshadow */

    If uWantsDropshadow(n) Then
        x = x - n
        y = y - n

    End If

End Sub

Private Function uWantsDropshadow(Optional ByRef ShadowSize As Long) As Boolean

    ' /* styles can veto the system drop shadow */

    uWantsDropshadow = ((mStyleFlags And S_STYLE_NO_DROPSHADOW) = 0)
    If uWantsDropshadow Then _
        ShadowSize = g_SafeLong(g_ConfigGet("dropshadow_size"))

End Function

Private Sub uPulse()

    If Not (mInstance Is Nothing) Then
        If mInstance.Pulse() Then _
            uQuickUpdate

    End If

End Sub

Public Function StylePath() As String

    StylePath = mStylePath

End Function

Public Function Frame() As BRect

    If Not (theWindow Is Nothing) Then _
        Set Frame = new_BRectFromRect(theWindow.Frame)

End Function

Private Function uDetailsMenu() As OMMenu

    Set uDetailsMenu = New OMMenu
    With uDetailsMenu
        .AddItem .CreateItem("$" & mInfo.ClassObj.App.Name & "#?" & mInfo.ClassObj.Name, "Generated by: " & mInfo.ClassObj.App.NameEx & " ('" & mInfo.ClassObj.Description & "' class)")
        .AddSeparator
        .AddItem .CreateItem("", "Received: " & Format$(mInfo.DateStamp, "d mmm yyyy") & " at " & Format$(mInfo.DateStamp, "ttttt"))
        .AddItem .CreateItem("", "API Version: " & CStr(mInfo.APIVersion))

        If ((mInfo.IntFlags And SN_NF_REMOTE)) And (Not (mInfo.Socket Is Nothing)) Then _
            .AddItem .CreateItem("", "Sender: " & mInfo.Socket.RemoteHostIP)

    End With

End Function

Public Function Title() As String

    Title = mInfo.Title

End Function

Public Function Text() As String

    Text = mInfo.Text

End Function

Public Sub QueueCountChanged()

    uRedraw

End Sub

Public Function IsMergeCandidate(ByRef ComparisonClass As TAlert, ByVal ComparisonTitle As String) As Boolean

    If (ComparisonClass Is Nothing) Or ((mInfo.IntFlags And SN_NF_MERGE) = 0) Then _
        Exit Function

    If (ComparisonClass.App.Signature = mInfo.ClassObj.App.Signature) And (ComparisonClass.Name = mInfo.ClassObj.Name) Then _
        IsMergeCandidate = (ComparisonTitle = mInfo.Title) Or (ComparisonTitle = "")

End Function

Public Sub RethinkActions()

    If (mInfo.Actions.CountItems = 1) Or (mInfo.Actions.CountItems = 0) Then _
        uQuickUpdate

End Sub

Public Function HasActions() As Boolean

    HasActions = (mInfo.Actions.CountItems > 0)

End Function

Private Sub uDoActionsMenu(Optional ByRef aRect As BRect)
Dim pr As BRect
Dim n As Long

    If mInfo.Actions.CountItems = 0 Then
        g_Debug "CSnarlWindow.uDoActionsMenu(): no actions defined", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    uWantsDropshadow n
    Set pr = mActionsRect.Duplicate
    pr.OffsetBy BW_Frame(theWindow.hWnd).Left + n, BW_Frame(theWindow.hWnd).Top + n

'    EnableWindow theWindow.hWnd, 0

    mMenuOpen = True

    Set theActionsPopup = New TAppsPopUpWindow
    theActionsPopup.Create 24

Dim pn As BTagItem

    With mInfo.Actions
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            theActionsPopup.AddItem pn.Name, pn.Value

        Loop
        
    End With

    theActionsPopup.Show pr


'Dim pmi As OMMenuItem
''Dim n As Long
'Dim i As Long
'
'    With New OMMenu
'        For i = 1 To mInfo.Actions.CountItems
'            .AddItem .CreateItem(CStr(i), mInfo.Actions.TagAt(i).Name)
'            If i < mInfo.Actions.CountItems Then _
'                .AddSeparator
'
'        Next i
'
'        ' /* get dropshadow compensation */
''        uWantsDropshadow n
'        Set pmi = .Track(theWindow.hWnd) '//, BW_Frame(theWindow.hWnd).TopLeft.OffsetByCopy(aRect.Left + n, aRect.Bottom + n))
'
'    End With
'
'    If (pmi Is Nothing) Then _
'        Exit Sub
'
'    ' /* get selected command */
'
'Dim szCmd As String
'
'    i = Val(pmi.Name)
'    szCmd = mInfo.Actions.TagAt(i).Value
'
'    If g_SafeLeftStr(szCmd, 1) = "@" Then
'        ' /* dynamic callback */
'        uNotify SNARL_NOTIFY_ACTION, szCmd
'
'    Else
'        ' /* static callback - process but don't do a callback */
'        g_ProcessAck szCmd
'
'    End If
'
'    EnableWindow theWindow.hWnd, -1
'
''    If (mInfo.Flags And SNARL41_NOTIFICATION_AUTO_DISMISS) Then _
'        Me.Hide
'
End Sub

Private Function uGetSNP3Info() As String
Dim ppd As BPackedData
Dim szn As String
Dim szv As String
Dim sz As String

    sz = "notification-uid: " & mInfo.CustomUID & vbCrLf

    Set ppd = New BPackedData
    With ppd
        .SetTo mInfo.OriginalContent
        .Rewind
        Do While .GetNextItem(szn, szv)
            If (LCase$(g_SafeLeftStr(szn, 5)) = "data-") Or _
               (LCase$(g_SafeLeftStr(szn, 6)) = "label-") Or _
               (LCase$(g_SafeLeftStr(szn, 6)) = "value-") Then _
                sz = sz & szn & ": " & szv & vbCrLf

        Loop

    End With

    uGetSNP3Info = sz

End Function

Private Function uGetForwardersMenu() As OMMenu
Dim ps As ConfigSection
Dim pMenu As OMMenu

    Set pMenu = New OMMenu

    With g_SubsRoster.Config
        .Rewind
        Do While .GetNextSection(ps)
            If ps.GetValueWithDefault("type", "") = "forwarder" Then _
                pMenu.AddItem pMenu.CreateItem(">" & ps.Name, ps.GetValueWithDefault("name", ""))

        Loop

    End With

    Set uGetForwardersMenu = pMenu

End Function

Private Function uGetRedirectorsMenu() As OMMenu
Dim pSchemesMenu As OMMenu
Dim pMenu As OMMenu
Dim ps As TStyle
Dim i As Long

    Set pMenu = New OMMenu

    With g_StyleRoster
        .Rewind
        Do While .GetNextStyle(ps)
            If ps.IsRedirect Then
                Set pSchemesMenu = New OMMenu
                For i = 1 To ps.CountSchemes
                    pSchemesMenu.AddItem pSchemesMenu.CreateItem("]" & ps.Name & "/" & ps.SafeSchemeAt(i), ps.SchemeAt(i))

                Next i
                pMenu.AddItem pMenu.CreateItem("", ps.Name, , , , , , pSchemesMenu)

            End If
        Loop
    End With

    Set uGetRedirectorsMenu = pMenu

End Function

Private Sub uDrawHUDContent()
Dim szHUD As String
Dim pr As BRect
Dim cy As Long

    With mView
'        .SetFont "Tahoma", 8, True
        .SetFont "Arial", 8, True
        .TextMode = MFX_TEXT_ANTIALIAS
        .SetHighColour rgba(255, 255, 255, 230)
        .SetLowColour rgba(0, 0, 0, 160)

        ' /* timestamp overlay */
        szHUD = g_When(mInfo.DateStamp)
        Set pr = new_BRect(0, 0, .StringWidth(szHUD) + 3, .StringHeight("A"))
        If g_ConfigGet("callback_as_button") = "1" Then
            pr.OffsetBy 4, .Height - pr.Height - 4

        Else
            pr.OffsetBy .Width - pr.Width, .Height - pr.Height

        End If

        cy = pr.Height
        .FillRoundRect pr, 3, 3, MFX_SOLID_LOW
        .DrawString szHUD, pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

        ' /* R2.4b4: app name overlay */
        szHUD = mInfo.ClassObj.App.NameEx()
        Set pr = new_BRect(0, 0, .StringWidth(szHUD) + 3, .StringHeight("A"))
        If g_ConfigGet("callback_as_button") = "1" Then
            pr.OffsetBy 4, .Height - pr.Height - cy - 4

        Else
            pr.OffsetBy .Width - pr.Width, .Height - cy - pr.Height

        End If
        .FillRoundRect pr, 3, 3, MFX_SOLID_LOW

        ' /* R2.4.2 DR3: if insecure app (or insecure notification), draw app name in different colour */

        If (mInfo.ClassObj.App.IsSecure) Or (mInfo.IntFlags And SN_NF_SECURE) Then
            .SetHighColour rgba(255, 255, 255)

        Else
            .SetHighColour rgba(255, 190, 110)

        End If

        .DrawString szHUD, pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

    End With

End Sub

''Dim n As Long
'
'        ' /* can only ever have one low priority notification visible at any one time
'        '    (obviously this excludes visible non-windowed styles) */
'
''        If mInfo.Priority < 0 Then _
'            n = g_NotificationRoster.QueueCount
'
'        ' /* if we have queued notifications, draw the number in a badge */
'
''        If n > 0 Then
''            .SetFont "Tahoma", 8, True
''            .TextMode = MFX_TEXT_ANTIALIAS
''            .EnableSmoothing True
''
''            ' /* figure out the badge size and position */
''
''            Set pr = new_BRect(0, 0, MAX(.StringWidth(CStr(n)), .StringHeight("A")), .StringHeight("A"))
''            pr.ExpandBy 8, 8
''            pr.OffsetBy .Width - pr.Width - 1, 1
''
''            If uWantsDropshadow() Then _
''                pr.OffsetBy -Val(g_ConfigGet("dropshadow_size")), Val(g_ConfigGet("dropshadow_size"))
''
''            ' /* background */
''
''            .SetHighColour rgba(0, 0, 64, 190)
''            .SetLowColour rgba(0, 0, 64, 100)
''            .FillRoundRect pr, RX, RX, MFX_VERT_GRADIENT
''
''            ' /* label */
''
''            .SetHighColour rgba(255, 255, 255)
''            .DrawString CStr(n), pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER
''
''            ' /* border */
''
''            .SetHighColour rgba(255, 255, 255)
''            .StrokeRoundRect pr.InsetByCopy(1, 1), RX, RX, 2
''
''            ' /* edge and inset */
''
''            .SetHighColour rgba(0, 0, 0, 150)
''            .StrokeRoundRect pr, RX, RX, 1
''            .StrokeRoundRect pr.InsetByCopy(3, 3), RX, RX, 1
''
''        End If
'
''        Set pb = .ConvertToBitmap()

Private Function uCanDoCallback() As Boolean

    uCanDoCallback = (mInfo.DefaultAck <> "") ' Or ((mInfo.hWndReply <> 0) And (mInfo.uReplyMsg <> 0))

End Function

Private Sub theActionsPopup_Closed()

    mMenuOpen = False
    Set theActionsPopup = Nothing
    uRedraw

End Sub

Private Sub theActionsPopup_Selected(ByVal Signature As String)

    ' /* get selected command */

    If g_SafeLeftStr(Signature, 1) = "@" Then
        ' /* dynamic callback */
        uNotify SNARL_NOTIFY_ACTION, Signature

    Else
        ' /* static callback - process but don't do a callback */
        g_ProcessAck Signature

    End If

End Sub

Private Sub uDoClose(ByRef PauseFlag As Boolean)

    If g_IsPressed(VK_CONTROL) Then
        ' /* CTRL held down... */
        g_NotificationRoster.CloseMultiple mInfo.ClassObj.App.Token

    ElseIf g_IsPressed(VK_SHIFT) Then
        ' /* SHIFT held down: close all */
        g_NotificationRoster.CloseMultiple 0

    Else
        PauseFlag = True                  ' // don't count down while we're doing this...
        bNotifyWasClosed
        PauseFlag = False
        Hide

    End If

End Sub

Friend Sub bNotifyWasClosed()

    uNotify SNARL_CALLBACK_CLOSED

End Sub

Private Sub uDoInvoke(ByRef PauseFlag As Boolean)

    PauseFlag = True

    If g_SafeLeftStr(mInfo.DefaultAck, 1) = "@" Then
        ' /* dynamic callback */
        uNotify SNARL_CALLBACK_INVOKED, mInfo.DefaultAck

    ElseIf mInfo.DefaultAck <> "" Then
        ' /* static callback - process but don't do a callback */
        g_ProcessAck mInfo.DefaultAck

    Else
        ' /* legacy */
        uNotify SNARL_CALLBACK_INVOKED

    End If

'    If mInfo.DefaultAck <> "" Then
'        g_ProcessAck mInfo.DefaultAck
'
'    Else
'        uNotify SNARL_NOTIFY_INVOKED, mInfo.DefaultAck
'
'    End If

    PauseFlag = False
    Hide

End Sub

'                    g_Debug "CSnarlWindow.WndProc(): close button invoked - not sending ACK"
'                    If (g_IsPressed(VK_LCONTROL)) Or (g_IsPressed(VK_RCONTROL)) Then
'                        ' /* CTRL held down... */
'                        If (g_IsPressed(VK_LSHIFT)) Or (g_IsPressed(VK_RSHIFT)) Then
'                            ' /* SHIFT held down: close all */
'                            g_NotificationRoster.CloseMultiple 0
'
'                        ElseIf mInfo.hWndReply <> 0 Then
'                            ' /* just from the provided window */
'                            g_NotificationRoster.CloseMultiple mInfo.hWndReply
'
'                        Else
'                            ' /* no reply window so just close ourselves */
'                            Hide
'
'                        End If
'
'                    Else
'                        ' /* just close ourself */
'                        bPaused = True                  ' // don't count down while we're doing this...
'                        uNotify SNARL_CALLBACK_CLOSED
'                        bPaused = False
'                        Hide
'
'                    End If

