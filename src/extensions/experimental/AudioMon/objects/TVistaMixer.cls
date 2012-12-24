VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TVistaMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event MuteChanged(ByVal Muted As Boolean)
Public Event VolChanged(ByVal Decibels As Single)

Const CLSCTX_INPROC_SERVER = 1
Const CLSCTX_INPROC_HANDLER = 2
Const CLSCTX_LOCAL_SERVER = 4
Const CLSCTX_REMOTE_SERVER = 16
Const CLSCTX_NO_CODE_DOWNLOAD = 400
Const CLSCTX_NO_FAILURE_LOG = 4000
Const CLSCTX_SERVER = CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER Or CLSCTX_REMOTE_SERVER
Const CLSCTX_ALL = CLSCTX_INPROC_HANDLER Or CLSCTX_SERVER
Const CLSCTX_INPROC = CLSCTX_INPROC_SERVER Or CLSCTX_INPROC_HANDLER

Dim mEndPoint As IAudioEndpointVolume
Dim mVolume As Single
Dim mMute As Long

Dim mNotify As TNotify

Public Function Initialize() As Boolean
Dim pEnum As MMDeviceEnumerator

    On Error Resume Next

    Err.Clear
    Set pEnum = New MMDeviceEnumerator
    If Err.Number Then
        g_Debug "TVistaMixer.Initialize(): failed to create MMDeviceEnumerator (" & Err.Description & ")", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pDev As IMMDevice

    pEnum.GetDefaultAudioEndpoint eRender, eMultimedia, pDev
    If (pDev Is Nothing) Then
        g_Debug "TVistaMixer.Initialize(): no default audio endpoint", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pRef As UUID

    With pRef
        .Data1 = &H5CDF2C82
        .Data2 = &H841E
        .Data3 = &H4546
        .Data4(0) = &H97
        .Data4(1) = &H22
        .Data4(2) = &HC
        .Data4(3) = &HF7
        .Data4(4) = &H40
        .Data4(5) = &H78
        .Data4(6) = &H22
        .Data4(7) = &H9A

    End With

    pDev.Activate pRef, CLSCTX_ALL, 0, mEndPoint
    If (mEndPoint Is Nothing) Then
        g_Debug "TVistaMixer.Initialize(): couldn't get IAudioEndpointVolume interface", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* register */

    Set mNotify = New TNotify

    If mEndPoint.RegisterControlChangeNotify(mNotify) = 0 Then
        Initialize = True

    Else
        g_Debug "TVistaMixer.Initialize(): couldn't register for notifications", LEMON_LEVEL_CRITICAL

    End If

End Function

Public Sub TidyUp()

    If (mEndPoint Is Nothing) Then _
        Exit Sub

    mEndPoint.UnregisterControlChangeNotify mNotify
    Set mEndPoint = Nothing

End Sub




