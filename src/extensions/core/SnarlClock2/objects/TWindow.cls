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

Public Event DoHourlyReminder()
Public Event TimerDone()

Dim mTimer As Date
Dim mTimerEnabled As Boolean

Implements MWndProcSink

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static dwId As Long
'Static prev As Long
Dim i As Long

    Select Case uMsg
    Case WM_NCCREATE
        SetTimer hWnd, 1, 990, 0

    Case WM_NCDESTROY
        KillTimer hWnd, 1

    Case WM_TIMER
        If wParam <> 1 Then _
            Exit Function

        ' /* hourly reminder */
        If (Minute(Now()) = 0) And (Second(Now()) = 0) Then _
            RaiseEvent DoHourlyReminder

'        If mTimerEnabled = True Then
'            ' /* doing a countdown */
'
'            mTimer = DateAdd("s", -1, mTimer)
'
'                If (Hour(mTimer) = 0) And (Minute(mTimer) = 0) Then
'                    Select Case Second(mTimer)
'                    Case 0
'                        ' /* we're done */
'                        If snIsMessageVisible(dwId) Then _
'                            snHideMessage dwId              ' // hide 10 second countdown notification
'
'                        RaiseEvent TimerDone
'                        mTimerEnabled = False
'
'                    Case 10
'                        If gConfig.timer_10sec_warning Then
'                            ' /* show the T-10 countdown notification */
'                            dwId = snShowMessageEx(CLASS_COUNTDOWN_COMPLETE, _
'                                                   Second(mTimer) & " second" & IIf(Second(mTimer) = 1, "", "s") & " to go", _
'                                                   "", _
'                                                   0, _
'                                                   IIf(gConfig.timer_icon = "", g_MakePath(App.Path) & "timer.png", gConfig.timer_icon))
'
'                        End If
'
'                    Case Is < 10
'                        If gConfig.timer_10sec_warning Then _
'                            snUpdateMessage dwId, CStr(Second(mTimer)) & " seconds to go", ""
'
'                    End Select
'
'                End If
'
'                g_UpdateTimerControls mTimer
'
''            End If
'
'        End If



'    Case WM_START_COUNTDOWN
'        ' /* lParam = minutes */
'
'        g_Debug "WM_START_COUNTDOWN: fEnabled=" & CStr(wParam) & " minutes=" & CStr(lParam)
'
'        If wParam = 1 Then
'            ' /* start */
'
'            If snIsMessageVisible(dwId) Then _
'                snHideMessage dwId
'
'            dwId = 0
'            mTimerEnabled = True
'
'        Else
'            ' /* stop */
'            mTimerEnabled = False
'
'        End If
'
'    Case WM_SET_COUNTDOWN
'        mTimer = TimeSerial(HiWord(lParam), LoWord(lParam), 0)
'        g_Debug "WM_SET_COUNTDOWN: " & CStr(mTimer)

    End Select

End Function

