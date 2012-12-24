VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TGenericInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TGenericInstance.cls
    '/
    '/  Description:    Represents an instance of an NMA or Toasty notification
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mStyle As TStyle
Dim mUID As Long

Public IsNotifyMyAndroid As Boolean

Dim WithEvents theNewRequest As CHTTPRequest2
Attribute theNewRequest.VB_VarHelpID = -1

Implements IStyleInstance
Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mUID)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = CStr(mUID)

End Function

Public Sub SetStyle(ByRef Style As TStyle)

    Set mStyle = Style

End Sub

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage
End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    If Me.IsNotifyMyAndroid Then
        uGoAndroid NotificationInfo

    Else
        uGoToasty NotificationInfo

    End If

End Sub

Private Sub uGoToasty(ByRef NotificationInfo As notification_info)

    If mStyle.ReadSetting("device_id") = "" Then
        g_Debug "TGenericInstance(): no device id specified", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo NotificationInfo.Text

    gID = gID + 1
    mUID = gID

Dim bData() As Byte

'    uGetIcon2 "c:\a.png", bData

    If ppd.Exists("icon") Then _
        uGetIcon2 ppd.ValueOf("icon"), bData

    Set theNewRequest = New CHTTPRequest2
    With theNewRequest
        .SetRequestType "POST"
        .AddRequestHeader "user-agent", "Snarl-Toasty/1.1"
        .AddRequestHeader "from", "samwise@system.local"
        .AddRequestHeader "accept", "*/*"
        .AddRequestHeader "content-length", CStr(UBound(bData) + 1)
        .SetBodyFromBytes bData()

        g_Debug "TGenericInstance().uGoToasty: sending request..."
        
'        .GetContent "http://127.0.0.1/notify/" & mStyle.ReadSetting("device_id") & _
                    "?title=" & g_URLEncode(ppd.ValueOf("title")) & _
                    "&text=" & g_URLEncode(ppd.ValueOf("text")) & _
                    "&sender=" & g_URLEncode(NotificationInfo.Title)
        
        
        .GetContent "http://api.supertoasty.com/notify/" & mStyle.ReadSetting("device_id") & _
                    "?title=" & g_URLEncode(ppd.ValueOf("title")) & _
                    "&text=" & g_URLEncode(ppd.ValueOf("text")) & _
                    "&sender=" & g_URLEncode(NotificationInfo.Title)

    End With

    g_AddThis Me

End Sub

Private Sub uGoAndroid(ByRef NotificationInfo As notification_info)

    If mStyle.ReadSetting("api_key") = "" Then
        g_Debug "TGenericInstance().uGoAndroid: no api key specified", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim sz() As String

    gID = gID + 1
    mUID = gID

    With NotificationInfo
        sz = Split(.Title, "#?")
        Debug.Print "application: " & sz(0)
        Debug.Print "event: " & sz(1)
        Debug.Print "description: " & NotificationInfo.Text
        Debug.Print "apikey: " & mStyle.ReadSetting("api_key")

    End With

Dim szBody As String

    szBody = "apikey=" & mStyle.ReadSetting("api_key") & _
             "&application=" & g_URLEncode(sz(0)) & _
             "&event=" & g_URLEncode(sz(1)) & _
             "&description=" & g_URLEncode(NotificationInfo.Text)

    Set theNewRequest = New CHTTPRequest2
    With theNewRequest
        .SetRequestType "POST"
        .AddRequestHeader "user-agent", "Snarl-NotifyMyAndroid/1.1"
        .AddRequestHeader "from", "frodo@system.local"
        .AddRequestHeader "accept", "*/*"
        .AddRequestHeader "content-length", CStr(LenB(szBody))
        .AddRequestHeader "content-type", "application/x-www-form-urlencoded"
        .SetBody szBody

        g_Debug "TGenericInstance().uGoAndroid: sending request..."
        .GetContent "https://nma.usk.bz/publicapi/notify"

    End With

    g_AddThis Me

End Sub

Private Sub theNewRequest_Completed(ByVal WasSuccessful As Boolean)

    g_Debug "TGenericInstance.Completed()", LEMON_LEVEL_PROC_ENTER
    g_Debug "request: " & CStr(mUID) & " wassuccessful: " & WasSuccessful

    If WasSuccessful Then
        g_Debug "response: " & Replace$(theNewRequest.ResponseBody, vbCrLf, "¶")
        g_Debug "headers: " & Replace$(theNewRequest.ResponseHeaders, vbCrLf, "¶")

    End If

    g_RemoveThis Me
    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

'Private Function uGetIcon(ByVal Path As String) As String
'Dim i As Integer
'Dim sz As String
'
'    ' /* gets icon data as a string */
'
'    On Error Resume Next
'
'    i = FreeFile()
'    Err.Clear
'    Open Path For Binary Access Read As #i
'    If Err.Number <> 0 Then
'        g_Debug "TGenericInstance.uGetIcon(): error opening '" & Path & "': " & Err.Description
'
'    Else
'        sz = String$(LOF(i), 0)
'        Get #i, , sz
'        uGetIcon = sz
'
'    End If
'
'    Close #i
'
'End Function

Private Function uGetIcon2(ByVal Path As String, ByRef Data() As Byte) As Boolean
Dim i As Integer

    ' /* gets icon data as a string */

    On Error Resume Next

    i = FreeFile()
    Err.Clear
    Open Path For Binary Access Read As #i
    If Err.Number <> 0 Then
        g_Debug "TGenericInstance.uGetIcon2(): error opening '" & Path & "': " & Err.Description, LEMON_LEVEL_CRITICAL

    ElseIf LOF(i) < 65535 Then
        ReDim Data(LOF(i) - 1)
        Get #i, , Data()
        uGetIcon2 = True

    Else
        g_Debug "TGenericInstance.uGetIcon2(): LOF exceeds 65KB limit", LEMON_LEVEL_CRITICAL

    End If

    Close #i

End Function
