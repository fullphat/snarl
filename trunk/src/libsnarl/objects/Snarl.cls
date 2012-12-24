VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Snarl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME = "libmsnarl_snooper"
Dim mhWnd As Long

Public Event SnarlLaunched()
Public Event SnarlQuit()

Implements BWndProcSink

'Public Function SimpleNotify(Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal IconPath As String) As Long
'
'    SimpleNotify = sn41EZNotify(0, "", Title, Text, , IconPath)
'
'End Function

Private Sub Class_Initialize()

    ' /* create our listener window */

    EZRegisterClass CLASS_NAME
    mhWnd = EZAddWindow(CLASS_NAME, Me)
    Debug.Print "snooper window is " & g_HexStr(mhWnd)

End Sub

Private Sub Class_Terminate()

    EZRemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    If uMsg = snBroadcastMsg() Then
        Select Case wParam
        Case SNARL_BROADCAST_LAUNCHED
            RaiseEvent SnarlLaunched

        Case SNARL_BROADCAST_QUIT
            RaiseEvent SnarlQuit

        End Select

    End If

End Function



