VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TForward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mDestination As String
Dim mProtocol As String
Dim mPort As Long
Dim mGuid As String
Dim mPassword As String
Dim mSalt As String

Dim mData As BTagList
Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1
Dim WithEvents theRetryTimer As BTimer
Attribute theRetryTimer.VB_VarHelpID = -1

Dim mIsSubscription As Boolean              ' // means we should send a {subscribe} on connection
Dim mConnected As Boolean                   ' // subscriptions only: successfully subscribed to remote computer

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mGuid

End Function

Private Function BTagItem_Value() As String
End Function

Friend Function ForwardTo(ByRef Info As T_NOTIFICATION_INFO, ByVal Destination As String, ByVal Protocol As String) As Boolean
Dim hr As Long
Dim sz As String

    mDestination = Destination
    mProtocol = Protocol
    mGuid = g_CreateGUID(True)
    Set mData = new_BTagList()

    Select Case mProtocol
    Case "snp"
        mData.Add new_BTagItem("", snp_CreateForward(Info))
        mPort = SNP_DEFAULT_PORT

    Case "gntp"
        mData.Add new_BTagItem("", uCreateGNTPRegister(Info.ClassObj.App))
        mData.Add new_BTagItem("", uCreateGNTPNotify(Info))
        mPort = GNTP_DEFAULT_PORT

    Case "smb"
        g_Debug "TForward.ForwardTo(): not implemented yet", LEMON_LEVEL_CRITICAL
        Exit Function
'        sz = StrConv("Hello, world!", vbUnicode)
'        hr = NetMessageBufferSend(vbNullString, StrConv(g_GetComputerName(), vbUnicode), StrConv(g_GetComputerName(), vbUnicode), sz, Len(sz))
'        MsgBox g_ApiError(hr)

    Case Else
        g_Debug "TForward.ForwardTo(): unknown protocol " & g_Quote(Protocol), LEMON_LEVEL_CRITICAL
        Exit Function

    End Select

    uConnect
    ForwardTo = True

End Function

Private Function uCreateGNTPRegister(ByRef App As TApp) As String

    ' /* base content */

    uCreateGNTPRegister = "GNTP/1.0 REGISTER NONE" & vbCrLf & _
                          "Application-Name: " & App.Name & vbCrLf & _
                          "Notifications-Count: " & CStr(App.CountAlerts) & vbCrLf & vbCrLf

    ' /* add classes */

Dim i As Long

    With App
        If .CountAlerts Then
            For i = 1 To .CountAlerts
                With .AlertAt(i)
                    uCreateGNTPRegister = uCreateGNTPRegister & "Notification-Name: " & .Name & vbCrLf & _
                                "Notification-Display-Name: " & .Description & vbCrLf & _
                                "Notification-Enabled: " & IIf(.IsEnabled, "True", "False") & vbCrLf & vbCrLf

                End With
            Next i
        End If
    End With

    ' /* add notification content and footer */

    uCreateGNTPRegister = uCreateGNTPRegister & vbCrLf & vbCrLf

End Function

Private Function uCreateGNTPNotify(ByRef Info As T_NOTIFICATION_INFO) As String
'Dim szClass As String

    ' /* base content */

    uCreateGNTPNotify = "GNTP/1.0 NOTIFY NONE" & vbCrLf & _
                        "Application-Name: " & Info.ClassObj.App.Name & vbCrLf

'    If Args.Exists("id") Then
'        szClass = Args.ValueOf("id")
'
'    Else
'        szClass = Args.ValueOf("class")
'
'    End If

    uCreateGNTPNotify = uCreateGNTPNotify & "Notification-Name: " & Info.ClassObj.Name & vbCrLf

    If Info.CustomUID <> "" Then _
        uCreateGNTPNotify = uCreateGNTPNotify & "Notification-ID: " & Info.CustomUID & vbCrLf

    uCreateGNTPNotify = uCreateGNTPNotify & "Notification-Title: " & (Replace$(Info.Title, vbCrLf, vbLf)) & vbCrLf
    uCreateGNTPNotify = uCreateGNTPNotify & "Notification-Text: " & (Replace$(Info.Text, vbCrLf, vbLf)) & vbCrLf

    If Info.Timeout = 0 Then _
        uCreateGNTPNotify = uCreateGNTPNotify & "Notification-Sticky: True" & vbCrLf

    uCreateGNTPNotify = uCreateGNTPNotify & "Notification-Priority: " & CStr(Info.Priority) & vbCrLf

'Notification-Icon: <url> | <uniqueid>
'Optional - The icon to display with the notification.

'Notification-Coalescing-ID: <string>
'Optional - If present, should contain the value of the Notification-ID header of a previously-sent notification. This serves as a hint to the notification system that this notification should replace/update the matching previous notification. The notification system may ignore this hint.

'Notification-Callback-Context: <string>
'Optional - Any data (will be passed back in the callback unmodified)
'
'Notification-Callback-Context-Type: <string>
'Optional, but Required if 'Notification-Callback-Context' is passed - The type of data being passed in Notification-Callback-Context (will be passed back in the callback unmodified). This does not need to be of any pre-defined type, it is only a convenience to the sending application.
'
'Notification-Callback-Target: <string>
'Optional - An alternate target for callbacks from this notification. If passed, the standard behavior of performing the callback over the original socket will be ignored and the callback data will be passed

    ' /* end marker */

    uCreateGNTPNotify = uCreateGNTPNotify & vbCrLf & vbCrLf

End Function

Public Function SubscribeTo(ByVal Name As String, ByVal Destination As String, ByVal Protocol As String, ByVal Guid As String, ByVal Password As String) As Boolean
Dim szHash As String

    mDestination = Destination
    mProtocol = LCase$(Protocol)
    mPassword = Password
    mIsSubscription = True
    mGuid = Guid

    Set mData = new_BTagList()

    Select Case mProtocol
    Case "snp"
        If g_IsAlphaBuild Then _
            g_PrivateNotify "", "[Internal Test Message]", "Subscribing via SNP to " & Destination & "..." & vbCrLf & "GUID: " & mGuid
'        MsgBox "Subscribing via SNP to " & Destination & "..." & vbCrLf & "GUID: " & mGuid

        mData.Add new_BTagItem("", snp_CreateSubscription(g_GetComputerName()))
        mPort = SNP_DEFAULT_PORT

    Case "gntp"
        If g_IsAlphaBuild Then _
            g_PrivateNotify "", "[Internal Test Message]", "Subscribing via GNTP to " & Destination & "..." & vbCrLf & "GUID: " & mGuid
        
        'Subscribed-to Machine's password: foo
        'Subscriber-ID value sent by subscribing machine: 0f8e3530-7a29-11df-93f2-0800200c9a66
        'Resulting password used by subscribed-to machine when forwarding: foo0f8e3530-7a29-11df-93f2-0800200c9a66
        szHash = mPassword '& Guid

        'The password is converted an UTF8 byte array
        'A cyptographically secure salt is generated (should be between 4 and 16 bytes)
        mSalt = HexStrToASCII("61C21AF94141BC4D")

        'The salt bytes are appended to the password bytes to form the key basis
        szHash = szHash & mSalt

        'The key is generated by computing the hash of the key basis using one of the supported hashing algorithms
        szHash = MD5DigestStrToHexStr(szHash)

        'The key hash is produced by computing the hash of the key (using the same hashing algorithm used in step 4)
        szHash = MD5DigestStrToHexStr(HexStrToASCII(szHash))

        'and hex-encoding it to a fixed-length string
        szHash = "MD5:" & szHash & "." & ASCIIToHexStr(mSalt)

'        MsgBox gntp_CreateSubscribeRequest(mGuid, g_GetComputerName(), szHash)

        mData.Add new_BTagItem("", gntp_CreateSubscribeRequest(mGuid, g_GetComputerName(), szHash))
        mPort = GNTP_DEFAULT_PORT

    Case Else
        g_Debug "TForward.ForwardTo(): unknown protocol " & g_Quote(Protocol), LEMON_LEVEL_CRITICAL
        Exit Function

    End Select

    uConnect
    SubscribeTo = True

End Function

Private Sub theRetryTimer_Pulse()

    If g_IsAlphaBuild Then _
        g_PrivateNotify "", "[Internal Test Message]", "Retrying connection to " & mDestination & "..."

    ' /* retry */
    uConnect

End Sub

Private Sub theSocket_OnClose()

    g_Debug "TForward.OnClose(): socket was closed"
    Set theSocket = Nothing

    If mProtocol = "gntp" Then
        If mData.CountItems > 0 Then
            ' /* this looks weird, but this is how GNTP works; each request typically
            '    uses a new socket/connection, so if the previous request succeeded
            '    the socket may still be closed by the receiver.  If we have more data
            '    to send, we reconnect here (which will then cause the send to happen
            '    when OnConnect() is fired) */
            g_Debug "TForward.OnClose(): more data to send - reconnecting..."
            Set theSocket = New CSocket
            theSocket.Connect mDestination, GNTP_DEFAULT_PORT

        Else
            ' /* no more data to send */
            If mIsSubscription Then
                ' /* TO DO: start the keep-alive timer */

            Else
                ' /* we're done forwarding: remove ourselves from the roster */
                g_SubsRoster.RemoveForwardOrSubscription mGuid

            End If
        End If

    ElseIf mProtocol = "snp" Then
        ' /* we're done, for whatever reason */
        g_SubsRoster.RemoveForwardOrSubscription mGuid

    End If

End Sub

Private Sub theSocket_OnConnect()

    g_Debug "TForward.OnConnect(): now connected to " & mDestination, LEMON_LEVEL_INFO
    Set theTimer = Nothing

    If mData.CountItems > 0 Then
        ' /* send and remove head node */
        g_Debug "TForward.OnConnect(): sending request..."
        theSocket.SendData mData.TagAt(1).Value
        mData.Remove 1
        
    End If

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim szResponseType As String
Dim sz As String
Dim i As Long

    theSocket.GetData sz
    Debug.Print "--"
    Debug.Print sz
    Debug.Print "--"

    Select Case mProtocol
    Case "snp"
        If mIsSubscription Then
            ' /* subscription - handle as per standard incoming request */
            If snp3_IsResponse(sz, szResponseType) Then
                ' /* should be the result of the {subscribe} */
                If g_IsAlphaBuild Then _
                    g_PrivateNotify "", "[Internal Test Message]", "Subscription via SNP to " & theSocket.RemoteHostIP & ":" & theSocket.RemotePort & " = " & g_Quote(szResponseType) & vbCrLf & "GUID: " & mGuid
                mConnected = (szResponseType = "OK")

            Else
                ' /* incoming forwarded notification from subscribed-to computer */
                snp3_Translate sz, theSocket, SN_NF_REMOTE Or SN_NF_IS_SNP3 Or SN_NF_FORWARD

            End If

        Else
            ' /* forward */
            If snp3_IsResponse(sz, szResponseType) Then
                If szResponseType <> "OK" Then
                    ' /* failed */
                    g_Debug "TForward.OnDataArrival(): previous SNP request failed - aborting...", LEMON_LEVEL_CRITICAL
                    mData.MakeEmpty
                    theSocket.CloseSocket

                ElseIf mData.CountItems > 0 Then
                    ' /* success and more data so send it and remove it */
                    g_Debug "TForward.OnDataArrival(): sending next SNP request..."
                    theSocket.SendData mData.TagAt(1).Value
                    mData.Remove 1

                Else
                    ' /* success and no more data to send */
                    theSocket.CloseSocket

                End If
            Else
                g_Debug "TForward.OnDataArrival(): spurious SNP request received", LEMON_LEVEL_CRITICAL
                If g_IsAlphaBuild Then _
                    g_PrivateNotify "", "[Internal Test Message]", "Spurious SNP request received from " & theSocket.RemoteHostIP & ":" & theSocket.RemotePort

            End If
        End If


    Case "gntp"
        If mIsSubscription Then
            If gntp_IsResponse(sz, szResponseType) Then
                ' /* should be the result of the SUBSCRIBE */
                If g_IsAlphaBuild Then _
                    g_PrivateNotify "", "[Internal Test Message]", "Subscription via GNTP to " & theSocket.RemoteHostIP & ":" & theSocket.RemotePort & " = " & g_Quote(szResponseType) & vbCrLf & "GUID: " & mGuid
                mConnected = (szResponseType = "-OK")

            Else
                ' /* incoming forwarded notification from subscribed-to computer */
    '                snp3_Translate sz, theSocket, SN_NF_REMOTE Or SN_NF_IS_SNP3 Or SN_NF_FORWARD

            End If

        Else
            ' /* forward - so this should be a response to the REGISTER and NOTIFY packets */
            If InStr(uGetFirstLine(sz, vbLf), "-OK") = 0 Then
                ' /* failed */
                g_Debug "TForward.OnDataArrival(): previous GNTP request failed - aborting...", LEMON_LEVEL_CRITICAL
                mData.MakeEmpty

            End If
        End If

        ' /* for GNTP, we always close the socket (the sender will probably do this anyway)
        '    if there's more data to send, then we'll do a reconnect from the OnSocketClose()
        '    event */
        theSocket.CloseSocket

    Case Else
        g_Debug "TForward.OnDataArrival(): unknown protocol " & g_Quote(mProtocol), LEMON_LEVEL_CRITICAL

    End Select

End Sub

Private Function uGetFirstLine(ByVal str As String, Optional ByVal EndMarker As String = vbCrLf) As String
Dim i As Long

    i = InStr(str, EndMarker)
    If i Then _
        uGetFirstLine = g_SafeLeftStr(str, i - 1)

End Function

Private Sub theTimer_Pulse()

    g_Debug "TForward.Pulse(): timed out connecting to " & g_Quote(mDestination), LEMON_LEVEL_INFO

    If g_IsAlphaBuild Then _
        g_PrivateNotify "", "[Internal Test Message]", "Timed out connecting to " & mDestination & " (" & mProtocol & ")" & vbCrLf & "GUID: " & mGuid

    If NOTNULL(theSocket) Then
        theSocket.CloseSocket
        Set theSocket = Nothing

    End If

    Set theTimer = Nothing

    If mIsSubscription Then
        ' /* if this is a subscription, wait 60 seconds and try again */
        Set theRetryTimer = new_BTimer(60000, True, True)

    Else
        ' /* failed forward, so discard data and remove from roster */
        Set mData = Nothing
        g_SubsRoster.RemoveForwardOrSubscription mGuid

    End If

End Sub

Public Function IsSubscription() As Boolean

    IsSubscription = mIsSubscription

End Function

Public Sub Unsubscribe()

    If (mIsSubscription) And (NOTNULL(theSocket)) Then _
        theSocket.SendData "SNP/3.0" & vbCrLf & "unsubscribe" & vbCrLf & "END" & vbCrLf

End Sub

Private Sub uConnect()

    g_Debug "TForward.uConnect(): attempting to connect to " & mDestination & ":" & CStr(mPort) & "...", LEMON_LEVEL_INFO
    Set theSocket = New CSocket
    theSocket.Connect mDestination, mPort
    Set theTimer = new_BTimer(5000, True, True)

End Sub

Public Sub TidyUp()

    If NOTNULL(theSocket) Then
        theSocket.CloseSocket
        Set theSocket = Nothing

    End If

End Sub

Public Function IsConnected() As Boolean

    IsConnected = mConnected

End Function
