Attribute VB_Name = "mMain"
Option Explicit

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
Dim mNotify As TNotify

Public ghWndDest As Long

Public Sub Main()

    On Error Resume Next

'    MsgBox Command$ & " " & App.PrevInstance

    If App.PrevInstance Then _
        Exit Sub

    If Command$ = "-quit" Then
        ' /* findwindow "audiomond" -- kill window */
        Exit Sub

    End If

    Form1.Show
    Form1.uAdd Command$


    ghWndDest = Val(Command$)

    Form1.uAdd CStr(ghWndDest)

    If ghWndDest = 0 Then _
        Exit Sub

Dim pEnum As MMDeviceEnumerator

    Set pEnum = New MMDeviceEnumerator
    If (pEnum Is Nothing) Then _
        Exit Sub

    Form1.uAdd "enum ok"

Dim pDev As IMMDevice

    pEnum.GetDefaultAudioEndpoint eRender, eMultimedia, pDev
    If (pDev Is Nothing) Then _
        Exit Sub

    Form1.uAdd "dev ok"

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
    If (mEndPoint Is Nothing) Then _
        Exit Sub

    Form1.uAdd "endpoint ok"

    ' /* register */

    Set mNotify = New TNotify

    If mEndPoint.RegisterControlChangeNotify(mNotify) <> 0 Then _
        Exit Sub

    Form1.uAdd "register ok"

    Form1.Show
    Form1.uAdd "destination is 0x" & Hex$(ghWndDest)

End Sub
