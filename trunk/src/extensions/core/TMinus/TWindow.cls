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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TWindow.cls
    '/
    '/  Description:    Handles messages directed to the extension's window
    '/
    '/  Created:        19-May-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* Notes:
    '
    '
    ' */

Dim mhWnd As Long
Dim mTrayIcon As BNotifyIcon
Dim mPaused As Boolean

Dim mOriginal As Long               ' // in seconds
Dim mTick As Long                   ' // in 10ths of a second

Dim mDue As String

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_ICON = 1

Public Event CountdownStarted(ByVal Duration As String, ByVal DueDate As Date)
Public Event CountdownComplete(ByVal Duration As String)

Implements MWndProcSink

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    Select Case uMsg
    Case WM_NCCREATE
        mhWnd = hWnd

        ' /* create tray icon */
        Set mTrayIcon = New BNotifyIcon
        With mTrayIcon
            .SetTo hWnd, &H4FF
            .Add App.Title, LoadImage(App.hInstance, 1&, IMAGE_ICON, 16, 16, 0), App.Title

        End With

    Case WM_NCDESTROY
        ' /* zap tray icon */
        mTrayIcon.Remove App.Title
        Set mTrayIcon = Nothing

    Case &H4FF
        ' /* from tray icon */

        Select Case lParam
        Case WM_RBUTTONUP
            uDoTrayMenu

        End Select

    Case WM_TIMER
        If wParam = 1 Then
            If Not mPaused Then
                mTick = mTick - 1
                If mTick = 0 Then
                    uReset
                    RaiseEvent CountdownComplete(uSecondsToText(mOriginal \ 10))

                Else
                    mTrayIcon.Update App.Title, , "Running: " & uSecondsToText(mTick \ 10, True) & " to go"

                End If

            End If

        End If

    Case WM_USER + 4

        Select Case wParam
        Case 0
            ' /* set a countdown running - lParam contains number of seconds */
            g_Debug "WM_USER+4: start countdown of " & CStr(lParam) & " second(s)"
            uSetCountdown lParam
            MWndProcSink_WndProc = True

        Case 1
            ' /* subtract now from specified time and set a countdown of that many seconds */
            g_Debug "WM_USER+4: start countdown to 0x" & g_HexStr(lParam)
            ReturnValue = uTimeOfDay(lParam)
            MWndProcSink_WndProc = True

        Case Else
            g_Debug "WM_USER+4: invalid wParam"

        End Select

    End Select

End Function

Private Function uTimeOfDay(ByVal TimeOfDay As Long) As TMINUS_STATUS_CODES
Dim dt As Date

    ' /* return zero for success or error number which can be passed back as an exit code */

    On Error Resume Next

    g_Debug "uTimeOfDay()", LEMON_LEVEL_PROC_ENTER
    
    Err.Clear
    
    dt = DateValue(Now)
    dt = DateAdd("h", HIBYTE(LoWord(TimeOfDay)), dt)
    dt = DateAdd("n", LOBYTE(LoWord(TimeOfDay)), dt)

    g_Debug "requested time is " & CStr(dt)

    If Err.Number <> 0 Then
        uTimeOfDay = TMINUS_BAD_DATE
        g_Debug Err.Description & " converting to a date", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

Dim n As Long

    Err.Clear
    n = DateDiff("s", Now, dt)

    If Err.Number <> 0 Then
        uTimeOfDay = TMINUS_BAD_SECONDS
        g_Debug Err.Description & " converting to seconds", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    g_Debug "number of seconds is " & n

    If n < 0 Then
        uTimeOfDay = TMINUS_DATE_IN_PAST
        g_Debug "requested time is in the past", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    uTimeOfDay = uSetCountdown(n)
    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub uDoTrayMenu()
Dim pmi As OMMenuItem

    SetForegroundWindow mhWnd
    
    With New OMMenu
        .AddItem .CreateItem("", "Quick Countdown", , , , , , uCCMenu())
        .AddSeparator
        .AddItem .CreateItem("", IIf(mDue = "", "No countdown running", "Due: " & mDue), , False)
        .AddItem .CreateItem("pause", IIf(mPaused, "Resume", "Pause"))
        .AddSeparator
        .AddItem .CreateItem("reset", "Reset")
        .AddSeparator
        .AddItem .CreateItem("about", "About " & App.Title)

        Set pmi = .Track(mhWnd)

    End With

    PostMessage mhWnd, WM_NULL, 0, ByVal 0&

    If Not (pmi Is Nothing) Then
        Select Case pmi.Name
        Case "about"
            mTrayIcon.ShowInfo App.Title, App.Title & " by full phat products", "This is a Snarl extension.  You can enable and disable it from within the Snarl Preferences panel.", B_NOTIFY_ICON_INFO

        Case "pause"
            mPaused = Not mPaused
            If mPaused Then _
                mTrayIcon.Update App.Title, , "Paused: " & uSecondsToText(mTick \ 10, True) & " to go"
        
        
        Case "reset"
            uSetCountdown mOriginal \ 10

        Case Else
            ' /* must be a quick countdown */
            uSetCountdown Val(pmi.Name)

        End Select

    End If

End Sub

Private Function uCCMenu() As OMMenu
Dim dw As Long
Dim dt As Date

Dim pErr As Boolean

    Set uCCMenu = New OMMenu
    With uCCMenu

        .AddItem .CreateItem("30", "30 seconds")
        .AddItem .CreateItem("60", "1 minute")
        .AddItem .CreateItem("300", "5 minutes")
        .AddItem .CreateItem("600", "10 minutes")
        .AddItem .CreateItem("1800", "30 minutes")

        .AddSeparator

        dt = DateAdd("h", 1, Now)
        dt = TimeSerial(Hour(dt), 0, 0) + DateValue(Now)
        dw = g_SafeLong(DateDiff("s", Now, dt), pErr)

        If (Not pErr) And (dw > 0) Then
            .AddItem .CreateItem(CStr(dw), Format$(dt, "short time") & " (" & CStr(dw \ 60) & " " & g_Pluralize("minute", (dw \ 60)) & ")")

        End If

        dt = TimeSerial(Hour(Now()), 30, 0) + DateValue(Now)
        dw = g_SafeLong(DateDiff("s", Now, dt), pErr)

        If (Not pErr) And (dw > 0) Then
            .AddItem .CreateItem(CStr(dw), Format$(dt, "short time") & " (" & CStr(dw \ 60) & " " & g_Pluralize("minute", (dw \ 60)) & ")")

        End If

'        .AddItem .CreateItem(CStr(dw), CStr(dt) & " (" & CStr(dw) & " minutes")

    End With

End Function

Private Function uSetCountdown(ByVal Seconds As Long) As TMINUS_STATUS_CODES

    g_Debug "uSetCountdown()", LEMON_LEVEL_PROC_ENTER

'    Seconds = (Seconds \ 30) * 30

    ' /* pre-set return value */

    uSetCountdown = TMINUS_SUCCESS

    ' /* 1.1 - do nothing if zero timeout */

    If Seconds = 0 Then
        g_Debug "resetting"
        uReset
        g_Debug "", LEMON_LEVEL_PROC_EXIT
        Exit Function

'    ElseIf Seconds > 86399 Then
'        g_Debug "seconds cannot be greater than 24 hours", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
'        uSetCountdown = TMINUS_BAD_SECONDS
'        Exit Function

    End If

    mOriginal = (Seconds * 10)
    mTick = mOriginal

    mDue = CStr(DateAdd("s", Seconds, Now))

    RaiseEvent CountdownStarted(uSecondsToText(Seconds, True), DateAdd("s", Seconds, Now))
    SetTimer mhWnd, 1, 100, 0

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function uSecondsToText(ByVal Seconds As Long, Optional ByVal Plural As Boolean = True) As String
Dim i As Long

'    g_Debug "uSecondsToText()", LEMON_LEVEL_PROC_ENTER

    If Seconds < 60 Then
        ' /* seconds */
        uSecondsToText = CStr(Seconds) & " second"
        If Plural Then _
            uSecondsToText = uSecondsToText & IIf(Seconds = 1, "", "s")

    ElseIf Seconds < 3600 Then
        ' /* minutes */
        i = Seconds \ 60
        uSecondsToText = CStr(i) & " minute"
        If Plural Then _
            uSecondsToText = uSecondsToText & IIf(i = 1, "", "s")

        i = (Seconds Mod 60)
        If i > 0 Then
            uSecondsToText = uSecondsToText & " " & CStr(i) & " second"
            If Plural Then _
                uSecondsToText = uSecondsToText & IIf(i = 1, "", "s")

        End If

    Else
        ' /* hours */
        i = Seconds \ 3600
        uSecondsToText = CStr(i) & " hour"
        If Plural Then _
            uSecondsToText = uSecondsToText & IIf(i = 1, "", "s")

        i = (Seconds Mod 3600) \ 60
        uSecondsToText = uSecondsToText & " " & CStr(i) & " minute"
        If Plural Then _
            uSecondsToText = uSecondsToText & IIf(i = 1, "", "s")

'        i = (Seconds Mod 60)
'        If i > 0 Then
'            uSecondsToText = uSecondsToText & " " & CStr(i) & " second"
'            If Plural Then _
'                uSecondsToText = uSecondsToText & IIf(i = 1, "", "s")
'
'        End If

    End If

'    g_Debug CStr(Seconds) & " == " & uSecondsToText
'    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub uReset()

    KillTimer mhWnd, 1
    mTrayIcon.Update App.Title, , App.Title
    mDue = ""

End Sub
