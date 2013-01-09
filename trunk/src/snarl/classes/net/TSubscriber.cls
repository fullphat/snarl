VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSubscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TSubscriber.cls
    '/
    '/  Description:    Represents a *subscribed* computer using GNTP or SNP
    '/
    '/  © 2012 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* more detail... */
    '
    ' For SNP, the subscribing computer should keep the socket open, so all Forward() needs to do
    ' is send the notification content through the existing socket.  For GNTP, a new socket is
    ' required each time, so we construct a TForward object, tell it its destination (the remote
    ' computer) and set it going; it will handle connecting and posting the notification all
    ' by itself.

Dim mAddress As String          ' // remote host
Dim mPort As Long               ' // remote host's port
Dim mLastBump As Date           ' // GNTP: last time a subscription requet from the remote IP was received
Dim mPassword As String
Dim mProtocol As String         ' // "snp" or "gntp" only for now
Dim mName As String             ' // friendly name

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = Me.AddressAndPort

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mProtocol

End Function

Public Function Password() As String

    Password = mPassword

End Function

Public Sub TidyUp()

    Set theSocket = Nothing

End Sub

Public Function Socket() As CSocket

    Set Socket = theSocket

End Function

Public Function RemoteHost() As String

    If NOTNULL(theSocket) Then _
        RemoteHost = theSocket.RemoteHostIP

End Function

Public Function SetTo(ByVal Protocol As String, ByRef RemoteSocket As CSocket, Optional ByRef Data As BPackedData) As Boolean

    g_Debug "TSubscriber.SetTo()", LEMON_LEVEL_PROC_ENTER

    mProtocol = LCase$(Protocol)

    Select Case mProtocol
    Case "snp"
        Set theSocket = RemoteSocket
        mAddress = RemoteSocket.RemoteHostIP
        mPort = RemoteSocket.RemotePort
        If NOTNULL(Data) Then
            mName = Data.ValueOf("name")
            mPassword = Data.ValueOf("password")

        End If
        SetTo = True

    Case "gntp"
        ' /* GNTP subscriptions work differently: after sending the request, the remote
        '    computer closes the port.  Any subsequent forwards should thus create a
        '    new connection to the supplied port */
        If NOTNULL(Data) Then
            mAddress = RemoteSocket.RemoteHostIP
            mPort = g_SafeLong(Data.ValueOf("reply-port"))
            mName = Data.ValueOf("name")
            SetTo = True

        End If

    Case Else
        g_Debug "Protocol " & g_Quote(mProtocol) & " not supported", LEMON_LEVEL_CRITICAL

    End Select


    If SetTo Then
        mLastBump = Now()
        ' /* notify */
        If g_ConfigGet("notify_when_subscriber_added") = "1" Then _
            g_PrivateNotify , "New subscriber added", Me.FormattedName & " subscribed to notifications from this computer", , ".sub-" & Protocol & "-add"

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function Protocol() As String

    Protocol = mProtocol

End Function

Public Function AddressAndPort() As String

    AddressAndPort = mAddress & ":" & CStr(mPort)

End Function

Public Function FormattedName() As String

    If mName = "" Then
        FormattedName = mAddress

    Else
        FormattedName = mName & " (" & mAddress & ")"

    End If

End Function

Public Function Port() As Long

    Port = mPort

End Function

Public Sub Bump()

    ' /* for GNTP (but not used at present): record the last time we
    '    received a SUBSCRIBE request from the remote computer */

    mLastBump = Now()

End Sub

Friend Sub Forward(ByRef Content As T_NOTIFICATION_INFO)
Dim pf As TForward

    Select Case mProtocol
    Case "snp"
        ' /* socket should be open */
        If NOTNULL(theSocket) Then
            ' /* yes, so send it through the socket as a forward */
            theSocket.SendData snp_CreateForward(Content)

        Else
            g_Debug "TSubscriber.Forward(): snp: socket invalid", LEMON_LEVEL_CRITICAL

        End If

    Case "gntp"
        ' /* nice and easy :) create a new TForward and tell it to do its thang */
        Set pf = New TForward
        pf.ForwardTo Content, mAddress, mProtocol
        g_SubsRoster.AddForwardOrSubscription pf

    End Select

End Sub

Private Sub theSocket_OnClose()

    g_Debug "TSubscriber.OnClose()", LEMON_LEVEL_PROC_ENTER
    g_Debug mAddress & ":" & CStr(mPort) & " dropped its connection", LEMON_LEVEL_INFO

    Select Case mProtocol
    Case "snp"

        If g_IsAlphaBuild() Then _
            g_PrivateNotify "", "[Internal Test Message]", "Subscriber " & mAddress & ":" & mPort & "dropped its conntection"

        g_Debug "is SNP: removing subscription..."
        g_SubsRoster.RemoveSubscriber theSocket, Nothing

    Case "gntp"
        g_Debug "is GNTP: expected..."

    End Select

    Set theSocket = Nothing
    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

