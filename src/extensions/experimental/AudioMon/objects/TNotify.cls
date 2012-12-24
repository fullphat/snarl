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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_AUDIO_LEVEL_CHANGED = "Audio volume changed"
Private Const CLASS_AUDIO_MUTED = "Audio muted"
Private Const CLASS_AUDIO_UNMUTED = "Audio unmuted"

Dim mVolume As Single
Dim mMute As Long

Implements IAudioEndpointVolumeCallback

Private Sub IAudioEndpointVolumeCallback_OnNotify(pNotify As MMDeviceAPILib.AUDIO_VOLUME_NOTIFICATION_DATA)

    Form1.uAdd "+"

'    On Error Resume Next
'
'    Open "c:\log.txt" For Append As #1
'
'    Print #1, "OK!"
'
'    Close #1

'    g_Debug "+++ " ' & pNotify.bMuted & " " & pNotify.fMasterVolume
'
    If pNotify.bMuted <> mMute Then
        mMute = pNotify.bMuted
        PostMessage Val(Form1.Tag), &H500, 0, ByVal mMute

    ElseIf pNotify.fMasterVolume <> mVolume Then
        mVolume = pNotify.fMasterVolume
        PostMessage Val(Form1.Tag), &H500, 1, ByVal pNotify.fMasterVolume * 100

    End If

End Sub
