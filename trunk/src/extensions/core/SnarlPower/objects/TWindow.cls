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
    '/  Created:        31-Mar-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* Notes:
    '
    '   We need a native window here simply because VB forms don't seem to receive
    '   posted broadcast messages for some reason.
    '
    ' */

Private Const WM_CREATE = &H1
Private Const WM_TIMER = &H113
Private Const WM_POWERBROADCAST = &H218
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Dim mhWnd As Long

Dim mCurrentPower As B_POWER_STATUS
Dim mDoneLowPowerWarning As Boolean
Dim mLowPowerThreshold As Long

Public Event OnPowerStatusChanged(ByVal Connected As Boolean)
Public Event OnLowPowerWarning()

Implements MWndProcSink

Private Sub Class_Initialize()

    EZRegisterClass App.ProductName
    mhWnd = EZAddWindow(App.ProductName, Me)

End Sub

Private Sub Class_Terminate()

    EZRemoveWindow mhWnd
    EZUnregisterClass App.ProductName

End Sub

Private Function MWndProcSink_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static n As B_POWER_STATUS
Static m As Long

    Select Case uMsg

    Case WM_CREATE
        mCurrentPower = get_power_status()


    Case WM_POWERBROADCAST
        n = get_power_status()
        If n <> mCurrentPower Then
            If n = B_POWER_STATE_CONNECTED Then
                RaiseEvent OnPowerStatusChanged(True)

            ElseIf n = B_POWER_STATE_DISCONNECTED Then
                RaiseEvent OnPowerStatusChanged(False)

            End If

            mCurrentPower = n

        End If


'    Case WM_USER + 1
'        mLowPowerThreshold = wParam
'        mDoneLowPowerWarning = False
'        SetTimer hWnd, 1, 100, 0


    Case WM_TIMER
        m = get_battery_charge_percent()
        If mDoneLowPowerWarning Then
            If m > mLowPowerThreshold Then _
                mDoneLowPowerWarning = False

        Else
            If m <= mLowPowerThreshold Then
                RaiseEvent OnLowPowerWarning
                mDoneLowPowerWarning = True

            End If

        End If

    End Select

End Function

Public Function IsValid() As Boolean

    IsValid = (mhWnd <> 0)

End Function

Public Sub SetLowPowerThreshold(ByVal Value As Long)

    g_Debug "TWindow.SetLowPowerThreshold(): " & CStr(Value), LEMON_LEVEL_PROC

    mLowPowerThreshold = Value
    mDoneLowPowerWarning = False
    SetTimer mhWnd, 1, 100, 0

End Sub
