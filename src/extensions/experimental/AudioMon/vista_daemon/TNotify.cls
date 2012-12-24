VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_MOUSEMOVE = &H200

Implements IAudioEndpointVolumeCallback

Private Sub IAudioEndpointVolumeCallback_OnNotify(pNotify As MMDeviceAPILib.AUDIO_VOLUME_NOTIFICATION_DATA)

    On Error Resume Next

    Form1.uAdd "*"
    Form1.uAdd pNotify.bMuted
    Form1.uAdd pNotify.fMasterVolume

'    PostMessage ghWndDest, &H501, 0, pNotify.bMuted

'    PostMessage Form1.hwnd, WM_MOUSEMOVE, 0, 453

'    PostMessage ghWndDest, &H500, 1, pNotify.fMasterVolume * 100

    Form1.uAdd "-"

'    Err.Clear

End Sub
