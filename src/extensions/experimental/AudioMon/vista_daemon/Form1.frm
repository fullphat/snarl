VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "audiomond"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   3915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' /*
'
'   Credits:
'
'       Sample code at http://www.codeproject.com/KB/vista/CoreAudio.aspx?msg=2124292
'       Microsoft API documentation at http://msdn.microsoft.com/en-us/library/dd370839(VS.85).aspx
'       Type library syntax from http://www.vbaccelerator.com/home/vb/Code/Libraries/Writing_CDs/IMAPI/article.asp
'       Type library constants from http://www.mvps.org/emorcillo/en/code/vb6/index.shtml
'       String handling code from http://www.xtremevbtalk.com/showthread.php?t=68956
'
' */

'Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Any) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'
'Dim mEnum As MMDeviceEnumerator
'
'Const CLSCTX_INPROC_SERVER = 1
'Const CLSCTX_INPROC_HANDLER = 2
'Const CLSCTX_LOCAL_SERVER = 4
'Const CLSCTX_REMOTE_SERVER = 16
'Const CLSCTX_NO_CODE_DOWNLOAD = 400
'Const CLSCTX_NO_FAILURE_LOG = 4000
'Const CLSCTX_SERVER = CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER Or CLSCTX_REMOTE_SERVER
'Const CLSCTX_ALL = CLSCTX_INPROC_HANDLER Or CLSCTX_SERVER
'Const CLSCTX_INPROC = CLSCTX_INPROC_SERVER Or CLSCTX_INPROC_HANDLER
'
'Dim mDev As IMMDevice
'Dim mEndPoint As IAudioEndpointVolume
'
'Private Sub Command1_Click()
'
'    On Error Resume Next
'
'    Err.Clear
'    Set mEnum = New MMDeviceEnumerator
'    If Err.Number Then
'        uAdd "Failed (" & Err.Description & ")"
'
'    Else
'        uAdd "MMDeviceEnumerator created ok"
'
'    End If
'
'End Sub
'
'Private Sub Command2_Click()
'Dim pDev As IMMDevice
'Dim pList As IMMDeviceCollection
'Dim i  As Long
'Dim dw As Long
'Dim hr As Long
'Dim st As Long
'
'    On Error Resume Next
'
'    If (mEnum Is Nothing) Then
'        MsgBox "Enumerator not created!", vbCritical
'
'    Else
'        hr = mEnum.EnumAudioEndpoints(eAll, DEVICE_STATEMASK_ALL, pList)
'        If hr = 0 Then
'            uAdd "enum succeeded!"
'
'Dim sz As String
'
'            If Not (pList Is Nothing) Then
'                Err.Clear
'                hr = pList.GetCount(dw)
'                If hr = 0 Then
'                    uAdd "IMMDeviceCollection.GetCount() returned: " & CStr(dw)
'
'                    For i = 0 To dw + 1
'                        Set pDev = Nothing
'                        hr = pList.Item(i, pDev)
'                        If hr = 0 Then
'                            uAdd "IMMDeviceCollection.Item(" & CStr(i) & ") ok - item: " & Not (pDev Is Nothing)
'                            st = 0
'                            Err.Clear
'                            pDev.GetState st
'                            uAdd Err.Description
'                            uAdd "IMDevice(" & CStr(i) & ").GetState: " & Hex$(st) & " (" & hr & ")"
'
'                            st = 0
'                            Err.Clear
'                            pDev.GetId st
'                            uAdd Err.Description
'                            uAdd "IMDevice(" & CStr(i) & ").GetID: " & Hex$(st)
'                            uAdd LPWSTRtoBSTR(st)
'
'                        Else
'                            uAdd "IMMDeviceCollection.Item(" & CStr(i) & ") failed: " & Hex$(hr)
'
'                        End If
'
'                    Next i
'
'                Else
'                    uAdd "IMMDeviceCollection.GetCount() failed: " & Hex$(hr)
'
'                End If
'
'            Else
'                uAdd "pList is NULL!"
'
'            End If
'
'        Else
'            uAdd "GetDefaultAudioEntpoint() enum failed: " & Hex$(hr)
'
'        End If
'
'    End If
'
'End Sub

Public Sub uAdd(ByVal Text As String)

    Text1.Text = Text1.Text & Text & vbCrLf
    Text1.SelStart = Len(Text1.Text)

End Sub

'Private Sub Command3_Click()
'Dim pDev As IMMDevice
'
'    On Error Resume Next
'
'    If (mEnum Is Nothing) Then
'        MsgBox "Enumerator not created!", vbCritical
'        Exit Sub
'
'
'    End If
'
'    If FAILED(mEnum.GetDefaultAudioEndpoint(eRender, eMultimedia, pDev)) Then
'        uAdd "GetDefaultAudioEndPoint() failed"
'        Exit Sub
'
'    End If
'
'    If (pDev Is Nothing) Then
'        uAdd "Returned IMMDevice is NULL"
'        Exit Sub
'
'    End If
'
'    uAdd "Got default audio endpoint for eRender/eMultimedia"
'
'Dim pEndPoint As IAudioEndpointVolume
'
'    Err.Clear
'
'Dim hr As Long
'
'Dim pRef As UUID
'
'    With pRef
'        .Data1 = &H5CDF2C82
'        .Data2 = &H841E
'        .Data3 = &H4546
'        .Data4(0) = &H97
'        .Data4(1) = &H22
'        .Data4(2) = &HC
'        .Data4(3) = &HF7
'        .Data4(4) = &H40
'        .Data4(5) = &H78
'        .Data4(6) = &H22
'        .Data4(7) = &H9A
'
'    End With
'
''Const IID_IAudioEndpointVolume = "{5CDF2C82-841E-4546-9722-0CF74078229A}"
'
'    hr = pDev.Activate(pRef, CLSCTX_ALL, 0, pEndPoint)
'    uAdd "Activate(): " & Hex$(hr)
'
'    If (pEndPoint Is Nothing) Then
'        uAdd "Returned IAudioEndpointVolume is NULL"
'        Exit Sub
'
'    End If
'
'    ' /* register */
'
''    hr = pEndPoint.RegisterControlChangeNotify(Me)
''    uAdd "RegisterControlChangeNotify(): " & Hex$(hr)
'
'Dim bMute As Long
'
'    If FAILED(pEndPoint.GetMute(bMute)) Then
'        uAdd "GetMute() failed"
'        Exit Sub
'
'    End If
'
'    uAdd "Current mute status: " & bMute
'
'    hr = pEndPoint.SetMute(IIf(bMute = 0, 1, 0), 0)
'    uAdd "SetMute(): " & Hex$(hr)
'
'
'
''            If Not (pDev Is Nothing) Then
''                Err.Clear
''                hr = pDev.GetId()
''                uAdd "IMMDevice.GetId(): " & Err.Number & " " & Err.Description
''                uAdd "IMMDevice.GetId(): " & hr
'''                uAdd "ppstrId: " & Hex$(dw)
''
''                Err.Clear
''                hr = pDev.GetState()
''                uAdd "IMMDevice.GetState(): " & Err.Number & " " & Err.Description
''                uAdd "IMMDevice.GetState(): " & hr
'''                uAdd "pdwState: " & Hex$(dw)
''
''
''            Else
''                uAdd "IMMDevice is NULL!"
''
''            End If
'
'End Sub
'
'Public Function LPWSTRtoBSTR(ByVal lpwsz As Long) As String
'    ' Input: a valid LPWSTR pointer lpwsz
'    ' Return: a sBSTR with the same character array
'    Dim cChars As Long
'    ' Get number of characters in lpwsz
'    cChars = lstrlenW(lpwsz)
'    ' Initialize string
'    LPWSTRtoBSTR = String$(cChars, 0)
'    ' Copy string
'    Call CopyMemory(ByVal StrPtr(LPWSTRtoBSTR), ByVal lpwsz, cChars * 2)
'
'End Function
'
'Private Function FAILED(ByVal hr As Long) As Boolean
'
'    FAILED = (hr <> 0)
'
'End Function
'
'Private Sub Command4_Click()
'
'    On Error Resume Next
'
'    If (mEnum Is Nothing) Then
'        MsgBox "Enumerator not created!", vbCritical
'        Exit Sub
'
'
'    End If
'
'    If FAILED(mEnum.GetDefaultAudioEndpoint(eRender, eMultimedia, mDev)) Then
'        uAdd "GetDefaultAudioEndPoint() failed"
'        Exit Sub
'
'    End If
'
'    If (mDev Is Nothing) Then
'        uAdd "Returned IMMDevice is NULL"
'        Exit Sub
'
'    End If
'
'    uAdd "Got default audio endpoint for eRender/eMultimedia"
'
'    Err.Clear
'
'Dim hr As Long
'
'Dim pRef As UUID
'
'    With pRef
'        .Data1 = &H5CDF2C82
'        .Data2 = &H841E
'        .Data3 = &H4546
'        .Data4(0) = &H97
'        .Data4(1) = &H22
'        .Data4(2) = &HC
'        .Data4(3) = &HF7
'        .Data4(4) = &H40
'        .Data4(5) = &H78
'        .Data4(6) = &H22
'        .Data4(7) = &H9A
'
'    End With
'
'    hr = mDev.Activate(pRef, CLSCTX_ALL, 0, mEndPoint)
'    uAdd "Activate(): " & Hex$(hr)
'
'    If (mEndPoint Is Nothing) Then
'        uAdd "Returned IAudioEndpointVolume is NULL"
'        Exit Sub
'
'    End If
'
'    ' /* register */
'
'    hr = mEndPoint.RegisterControlChangeNotify(New TNotify)
'    uAdd "RegisterControlChangeNotify(): " & Hex$(hr)
'
'End Sub
'
