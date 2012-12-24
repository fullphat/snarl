VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TToastyInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mStyle As TStyle
Dim mUID As Long

Public IsNotifyMyAndroid As Boolean

Dim WithEvents theRequest As BWebRequest
Attribute theRequest.VB_VarHelpID = -1

Implements IStyleInstance
Implements IMeterStyle
Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mUID)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = CStr(mUID)

End Function

Private Sub IMeterStyle_SetStyle(Style As TStyle)

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

Private Sub theRequest_Completed(ByVal WasSuccessful As Boolean)

    Debug.Print "instance " & CStr(mUID) & " done: " & WasSuccessful
    If WasSuccessful Then _
        Debug.Print "response: " & theRequest.Content

    g_RemoveThis Me

End Sub

Private Sub uGoToasty(ByRef NotificationInfo As notification_info)

    If mStyle.ReadSetting("device_id") = "" Then
        g_Debug "TToastyInstance(): no device id specified", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim sz() As String

    gID = gID + 1
    mUID = gID

    With NotificationInfo
        sz = Split(.Title, "#?")

        Debug.Print "sender: " & sz(0)
        Debug.Print "title: " & sz(1)
        Debug.Print "text: " & NotificationInfo.Text
        Debug.Print "device: " & mStyle.ReadSetting("device_id")

    End With

    Set theRequest = New BWebRequest

    With theRequest
        .GetContent "http://api.supertoasty.com/notify/" & mStyle.ReadSetting("device_id") & _
                    "?title=" & sz(1) & _
                    "&text=" & NotificationInfo.Text & _
                    "&sender=" & sz(0) & _
                    "&image="

    End With

    g_AddThis Me

End Sub

Private Sub uGoAndroid(ByRef NotificationInfo As notification_info)

    If mStyle.ReadSetting("device_id") = "" Then
        g_Debug "TToastyInstance(): no device id specified", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim sz() As String

    gID = gID + 1
    mUID = gID

    With NotificationInfo
        sz = Split(.Title, "#?")

        Debug.Print "sender: " & sz(0)
        Debug.Print "title: " & sz(1)
        Debug.Print "text: " & NotificationInfo.Text
        Debug.Print "device: " & mStyle.ReadSetting("device_id")

    End With

    Set theRequest = New BWebRequest

    With theRequest
        .GetContent "http://api.supertoasty.com/notify/" & mStyle.ReadSetting("device_id") & _
                    "?title=" & sz(1) & _
                    "&text=" & NotificationInfo.Text & _
                    "&sender=" & sz(0) & _
                    "&image="

    End With

    g_AddThis Me

End Sub


