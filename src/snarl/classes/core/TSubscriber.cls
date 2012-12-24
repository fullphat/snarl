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
    '/  Description:    Represents a subscribed GNTP and SNP3 client
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Enum SN_SUBSCRIBER_TYPE

    ' /* - - - - | P D M M | T T T T | V V V V
    '   ---------+---------+---------+---------
    '
    '   P = Persistence (1 = persistent, 0 = transient)
    '       persistent connections are recorded in network.snarl and are loaded/managed
    '       by the subscription roster
    '   D = Direction (1 = outbound, 0 = inbound)
    '       this indicates the direction of travel of the notification packets
    '   M = Medium (00 = TCP, 01 = HTTP, 10 = reserved, 11 = reserved)
    '   T = Type (00 = SNP, 01 = GNTP, 10 = reserved, 11 = reserved)
    '   V = Version (based Type), currently defined:
    '            T     V
    '         ------+---------------
    '          SNP  | 1, 2 or 3
    '          GNTP | 1
    '
    ' */

    SN_ST_TRANSIENT = 0                 ' // can't use as mask - use SN_ST_PERSISTENT instead
    SN_ST_PERSISTENT = &H800&
    SN_ST_INBOUND = 0                   ' // can't use as mask - use SN_ST_OUTBOUND instead
    SN_ST_OUTBOUND = &H400&

    ' /* medium */
    
    SN_ST_MEDIUM_MASK = &H300&
    SN_ST_TCP = 0                       ' // direct TCP socket
    SN_ST_HTTP = &H100&                 ' // HTTP transfer

    ' /* type */

    SN_ST_TYPE_MASK = &HF0&
    SN_ST_SNP = 0                       ' // Snarl Network Protocol
    SN_ST_GNTP = &H10&                  ' // Growl Network Transport Protocol

    ' /* version */
    
    SN_ST_VERSION_MASK = &HF&

    ' /*
    '   non-persistent types (not recorded in network.snarl)
    ' */

    ' // subscriptions made *to* this computer from a remote computer
    SN_ST_SUBSCRIBER = SN_ST_TRANSIENT Or SN_ST_INBOUND Or SN_ST_TCP
    SN_ST_SNP3_SUBSCRIBER = SN_ST_SUBSCRIBER Or SN_ST_SNP Or 3
    SN_ST_GNTP_SUBSCRIBER = SN_ST_SUBSCRIBER Or SN_ST_GNTP Or 1

    ' /*
    '   persistent types (recorded in network.snarl)
    ' */
    SN_ST_FORWARD = SN_ST_PERSISTENT Or SN_ST_OUTBOUND Or SN_ST_TCP
    SN_ST_SNP3_FORWARD = SN_ST_FORWARD Or SN_ST_SNP Or 3
    SN_ST_GNTP_FORWARD = SN_ST_FORWARD Or SN_ST_GNTP Or 1

    SN_ST_SUBSCRIPTION = SN_ST_PERSISTENT Or SN_ST_INBOUND Or SN_ST_TCP
    SN_ST_SNP3_SUBSCRIPTION = SN_ST_SUBSCRIPTION Or SN_ST_SNP Or 3
    SN_ST_GNTP_SUBSCRIPTION = SN_ST_SUBSCRIPTION Or SN_ST_GNTP Or 1

    ' /* special types */

    SN_ST_COMET = SN_ST_PERSISTENT Or SN_ST_INBOUND Or SN_ST_HTTP

End Enum

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Dim mType As SN_SUBSCRIBER_TYPE
Dim mApps As BTagList               ' // which app(s) we're subscribed to - an empty list means "all"

Dim mRemoteAddr As String           ' // where to forward notifications or computer to subscribe to
Dim mDescription As String
Dim mPassword As String
'Dim mGuid As String

Dim mSuccess As Boolean
Dim mCache As BTagList              ' // requests are cached while we wait for the socket to connect

Dim mEnabled As Boolean
Dim mConfig As ConfigSection


Dim mAddressAndPort As String
Dim mProtocol As String

    ' /* a word on granularity: a subscriber is an IP address and a port and any number of apps - password is per subscriber */

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mAddressAndPort

End Function

Private Function BTagItem_Value() As String

'    BTagItem_Value = mType

End Function

Private Sub Class_Initialize()

    Set mCache = new_BTagList()
    Set mApps = new_BTagList()

    mEnabled = True     ' // temp

End Sub

Private Sub theSocket_OnClose()

'    g_Debug "TSubscriber.OnClose(): remote computer '" & mRemoteAddr & "' closed the connection", LEMON_LEVEL_INFO
'
'    Select Case mType
'
'    Case SN_ST_SNP3_FORWARD, SN_ST_GNTP_FORWARD, SN_ST_SNP3_SUBSCRIPTION, SN_ST_GNTP_SUBSCRIPTION
'        g_Debug "TSubscriber.OnClose(): was a forward/subscription"
'        frmAbout.bForwardersChanged
'
'    Case SN_ST_SNP3_SUBSCRIBER, SN_ST_GNTP_SUBSCRIBER
'        g_Debug "TSubscriber.OnClose(): removing remote subscription"
'        g_SubsRoster.RemoveByGuid mGuid
'        frmAbout.bSubsChanged
'
'    End Select
'
'    uReset

End Sub

'Public Function SetFrom(ByRef Config As ConfigSection) As Boolean
'
''    ' /* configure from an existing (or recently created ConfigSection).  This implies
''    '    persistency and thus means we're either a forwarder or a subscription to a
''    '    remote computer */
''
''
''    ' /* provided IP cannot be same as this computer (or NULL) */
''    mRemoteAddr = Config.GetValueWithDefault("remotehost", "")
''
''    If g_IsLocalAddress(mRemoteAddr) Then
''        g_Debug "TSubscriber.SetFrom(): invalid IP address '" & mRemoteAddr & "'", LEMON_LEVEL_CRITICAL
''        Exit Function
''
''    End If
''
''    mType = g_SafeLong(Config.GetValueWithDefault("type", ""))
''
''    Select Case mType
''    Case SN_ST_GNTP_FORWARD, SN_ST_GNTP_SUBSCRIPTION, SN_ST_SNP3_FORWARD, SN_ST_SNP3_SUBSCRIPTION
''
''    Case Else
''        g_Debug "TSubscriber.SetFrom(): type must be one of SN_ST_GNTP_FORWARD, SN_ST_GNTP_SUBSCRIPTION, SN_ST_SNP3_FORWARD, SN_ST_SNP3_SUBSCRIPTION", LEMON_LEVEL_CRITICAL
''        Exit Function
''
''    End Select
''
''    ' /* initialize */
''
''    Set mConfig = Config
''    mGuid = Config.Name
''    mDescription = Config.GetValueWithDefault("description", "")
''    mEnabled = (Config.GetValueWithDefault("enabled", "0") = "1")
''
''    ' /* TO-DO: allow forwarding or subscribing to named apps */
''    '// Config.GetValueWithDefault("apps", "")
''
''    ' /* create the socket */
''
''    uReset
''
''    SetFrom = True
'
'End Function

'Public Function SetTo(ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef RemoteSocket As CSocket, ByRef Data As BPackedData) As Boolean
'
'    ' /* configure from the provided information.  This implies transiency and thus
'    '    means we're recording a subscription received from a remote computer */
'
'    ' /* error checking */
'
'    If (SubscriberType <> SN_ST_GNTP_SUBSCRIBER) And (SubscriberType <> SN_ST_SNP3_SUBSCRIBER) Then
'        g_Debug "TSubscriber.SetTo(): SubType can only be SN_ST_GNTP_SUBSCRIBER or SN_ST_SNP3_SUBSCRIBER", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End If
'
'    mType = SubscriberType
'    mGuid = g_CreateGUID(True)          ' // we'll create our own...
'    mRemoteAddr = RemoteSocket.RemoteHostIP
'    Set theSocket = RemoteSocket
'
''    mDescription = Description
'
'    g_Debug "TSubscriber.SetTo(): subscribed to from " & mRemoteAddr, LEMON_LEVEL_INFO
'
'
'    SetTo = True
'
'End Function

Public Function send(ByVal Content As String) As Boolean

    If ISNULL(theSocket) Then _
        Exit Function

    ' /* if we're connected to the remote destination send immediately
    '    otherwise cache content */

    Select Case mType
    Case SN_ST_SNP3_FORWARD, SN_ST_SNP3_SUBSCRIBER
        If g_IsLocalAddress(theSocket.RemoteHostIP, True) Then
            Debug.Print "TSubscriber.Send(): invalid or local address '" & Me.RemoteHost & "' (" & Me.Description & ")"

        ElseIf theSocket.State = sckConnected Then
            Debug.Print "TSubscriber.Send(): sending content to '" & Me.RemoteHost & "'..."
            theSocket.SendData Content

        Else
            Debug.Print "TSubscriber.Send(): caching content..."
            mCache.Add new_BTagItem(Content, "")
            If mCache.CountItems > 32 Then
                g_Debug "TSubscriber.Send(): pruning cache..."
                mCache.Remove 1
        
            End If

        End If

    Case SN_ST_GNTP_FORWARD
        ' /* need to create a REGISTRATION packet and a NOTIFY
        '    packet and send them both */

'        sz = uCreateGNTPForward(Content)
'        szn = uCreateGNTPNotify(Content)
'        If theSocket.State = sckConnected Then
'            Debug.Print "TSubscriber.Send(): sending registration..."
'            theSocket.SendData sz
'            theSocket.SendData szn
'
'        Else
'            Debug.Print "TSubscriber.Send(): caching content..."
'            mCache.Add new_BTagItem(sz, "")
'            mCache.Add new_BTagItem(szn, "")
'            If mCache.CountItems > 64 Then
'                g_Debug "TSubscriber.Send(): pruning cache..."
'                mCache.Remove 1
'                mCache.Remove 1
'
'            End If
'
'        End If



    End Select




''        ' /* SNP 3.0 subscription (sent as FORWARD) */
''
''        sz = "SNP/3.0 FORWARD" & vbCrLf
''
''        ' /* add app info */
''
''        sz = sz & "app-signature: " & App.Signature & vbCrLf
''        sz = sz & "app-name: " & App.Name & vbCrLf
''
''        ' /* add selected parts of the content */
''
'''        Set ppd = New BPackedData
'''        content.SetTo Content
''
''        ' /* standard components */
''
''        sz = sz & "class: " & Content.ValueOf("id") & vbCrLf
''        sz = sz & "title: " & Replace$(Content.ValueOf("title"), vbCrLf, "\n") & vbCrLf
''        sz = sz & "text: " & Replace$(Content.ValueOf("text"), vbCrLf, "\n") & vbCrLf
''        sz = sz & "priority: " & Content.ValueOf("priority") & vbCrLf
''        sz = sz & "uid: " & Content.ValueOf("uid") & vbCrLf                 ' // TO DO: should we prefix with sender host name?
''
''        ' /* icon */
''
''        szn = Content.ValueOf("icon")
''        If szn <> "" Then
''            If (LCase$(g_SafeLeftStr(szn, 7)) = "http://") Or _
''               (LCase$(g_SafeLeftStr(szn, 8)) = "https://") Or _
''               (LCase$(g_SafeLeftStr(szn, 1)) = "!") Then
''                ' /* add verbatim */
''                sz = sz & "icon: " & szn & vbCrLf
''
''            Else
''                ' /* encode it in a slightly modified Base64 format (CRLF's are replaced with #'s) */
''                If uEncodeIcon(szn, szv) Then _
''                    sz = sz & "icon-phat64: " & szv & vbCrLf
''
''            End If
''
''            szn = ""
''            szv = ""
''
''        End If
''
''        ' /* add data-*, label-* and value-* elements */
''
''        With Content
''            .Rewind
''            Do While .GetNextItem(szn, szv)
''                If (LCase$(g_SafeLeftStr(szn, 5)) = "data-") Or _
''                   (LCase$(g_SafeLeftStr(szn, 6)) = "label-") Or _
''                   (LCase$(g_SafeLeftStr(szn, 6)) = "value-") Then _
''                    sz = sz & szn & ": " & szv & vbCrLf
''
''            Loop
''
''        End With
''
''        ' /* standard headers */
''
''        sz = sz & "x-timestamp: " & Format$(Now(), "d mmm yyyy hh:mm:ss") & vbCrLf
''        sz = sz & "x-daemon: " & "Snarl " & CStr(APP_VER) & "." & CStr(APP_SUB_VER) & vbCrLf
''        sz = sz & "x-host: " & LCase$(g_GetComputerName()) & vbCrLf
''
''        sz = sz & "END" & vbCrLf
''        theSocket.SendData sz
''
''
''    Case SN_ST_GNTP
'
'
'
'
'    End Select

End Function

Public Function Destination() As String

    Destination = mRemoteAddr

End Function

Public Function UID() As String

    UID = BTagItem_Name()

End Function

Public Function SubscriptionType() As SN_SUBSCRIBER_TYPE

    SubscriptionType = mType

End Function

Public Function Detail() As String
'Dim pt As BTagItem
'Dim i As Long
'
'    Detail = IIf(mType = SN_ST_SNP3_FORWARD, "Forwarding ", "Subscribed to ")
'
'    If mApps.CountItems = 0 Then
'        Detail = Detail & "everything"
'
'    Else
'        With mApps
'            .Rewind
'            Do While .GetNextTag(pt) = B_OK
'                i = i + 1
'                If i > 4 Then
'                    Detail = Detail & "and " & CStr(mApps.CountItems - 4) & " more, "
'                    Exit Do
'
'                Else
'                    Detail = Detail & pt.Name & ", "
'
'                End If
'
'            Loop
'
'        End With
'
'        Detail = g_SafeLeftStr(Detail, Len(Detail) - 2)
'
'    End If
'
'
'    If (mType <> SN_ST_GNTP_SUBSCRIBER) And (mType <> SN_ST_SNP3_SUBSCRIBER) Then _
'        Detail = Detail & IIf((mType And SN_ST_OUTBOUND), " to ", " from ") & mRemoteAddr
'
'
'    If (mType And SN_ST_TYPE_MASK) = SN_ST_GNTP Then _
'        Detail = Detail & " (using GNTP)"

End Function

Public Function Socket() As CSocket

    Set Socket = theSocket

End Function

Public Function AddApps(ByRef Data As BPackedData) As Boolean
Dim sPassword As String
Dim szn As String
Dim szv As String

    ' /* returns TRUE if at least one subscription succeeded */

    If Data.Exists("password") Then
        If Data.ValueOf("password") <> mPassword Then
            g_Debug "TSubscriber.AddApps(): password mismatch", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_AUTH_FAILURE
            Exit Function

        End If

    End If

    With Data
        .Rewind
        Do While .GetNextItem(szn, szv)
            If szn = "app-sig" Then
                If mApps.IndexOf(szv) = 0 Then
                    mApps.Add new_BTagItem(szv, "")
                    g_Debug "TSubscriber.AddApps(): subscribed to '" & szv & "'" '& IIf(sPassword <> "", " (using password)", "")
                    AddApps = True

                Else
                    g_Debug "TSubscriber.AddApps(): already subscribed to '" & szv & "'", LEMON_LEVEL_WARNING

                End If
            End If
        Loop
    End With

End Function

Public Function Password() As String

    Password = mPassword

End Function

Public Function IsSubscribedTo(ByVal AppSig As String) As Boolean

    If mApps.CountItems = 0 Then
        IsSubscribedTo = True

    Else
        IsSubscribedTo = (mApps.IndexOf(AppSig) <> 0)

    End If

End Function

Public Function Apps() As BTagList

    Set Apps = mApps

End Function

Public Function TypeStr() As String

    Select Case mType
    Case SN_ST_GNTP_SUBSCRIBER

    Case SN_ST_SNP3_SUBSCRIBER

    Case SN_ST_SNP3_FORWARD
        TypeStr = "FWD/SNP"
    
    Case SN_ST_SNP3_SUBSCRIPTION

    Case Else
        TypeStr = "UNK?"

    End Select

End Function

Public Function Description() As String

    If mDescription <> "" Then
        Description = mDescription

    Else
        Description = mRemoteAddr

    End If

End Function

'Public Function Guid() As String
'
'    Guid = mGuid
'
'End Function

Private Sub theSocket_OnConnect()
Dim pt As BTagItem

    Select Case mType
    Case SN_ST_SNP3_SUBSCRIPTION
        ' /* subscription from this computer to a remote one */
        g_Debug "TSubscriber.OnConnect(SN_ST_SNP3_SUBSCRIPTION): now connected to '" & theSocket.RemoteHost & "'"
        theSocket.SendData snp3_CreateSubscribeRequest()
        g_Debug "TSubscriber.OnConnect(SN_ST_SNP3_SUBSCRIPTION): sent subscription request"
        frmAbout.bForwardersChanged

    Case SN_ST_GNTP_SUBSCRIPTION
        g_Debug "TSubscriber.OnConnect(): GNTP subscriptions are not currently implemented", LEMON_LEVEL_WARNING


    Case SN_ST_SNP3_FORWARD, SN_ST_GNTP_FORWARD
        g_Debug "TSubscriber.OnConnect(): [SN_ST_x_FORWARD] now connected to '" & theSocket.RemoteHost & "'"
        With mCache
            .Rewind
            Do While .GetNextTag(pt) = B_OK
                theSocket.SendData pt.Name

            Loop
            .MakeEmpty

        End With
        frmAbout.bForwardersChanged

    Case Else
        g_Debug "TSubscriber.OnConnect(): spurious connection (dest=" & mRemoteAddr & " type=" & g_HexStr(mType, 2) & ")", LEMON_LEVEL_WARNING

    End Select
    
'    Set theTimer = Nothing

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim sz As String

    On Error GoTo er

    Select Case mType
    
    Case SN_ST_SNP3_SUBSCRIPTION
        ' /* subscription from this computer to a remote one so this should be incoming
        '    registrations and notifications */

        theSocket.PeekData sz

'        Debug.Print "()"
'        Debug.Print sz
'        Debug.Print "() " & InStr(sz, "SNP/") & " " & InStr(sz, vbCrLf & "END" & vbCrLf)

        If (InStr(sz, "SNP/") > 0) And (InStr(sz, vbCrLf & "END" & vbCrLf) > 0) Then
            theSocket.GetData sz
            snp3_Translate sz, theSocket, SN_NF_IS_SNP3 Or SN_NF_FORWARD Or SN_NF_REMOTE

        Else
            g_Debug "TSubscriber.OnDataArrival(SN_ST_SNP3_SUBSCRIPTION): waiting for complete packet..."

        End If

    End Select

    Exit Sub

er:
    g_Debug "TSubscriber.OnDataArrival(): " & err.Description, LEMON_LEVEL_CRITICAL
    theSocket.CloseSocket

End Sub

Private Sub theSocket_OnSendComplete()
'//
End Sub

'Private Sub theTimer_Pulse()
'
'    g_Debug "TSubscriber.Pulse(): timed out connecting to '" & mRemoteAddr & "' (type=" & g_HexStr(mType, 2) & ")", LEMON_LEVEL_CRITICAL
'    theSocket.CloseSocket
'    Set theSocket = Nothing
'
'End Sub

Public Sub SetEnabled(ByVal Enabled As Boolean)

    mEnabled = Enabled
    g_Debug "TSubscriber.SetEnabled(): setting '" & mDescription & "' to " & mEnabled & "..."

    If Not (mConfig Is Nothing) Then
        mConfig.Update "enabled", IIf(Enabled, "1", "0")
        g_SubsRoster.WriteConfig

    End If

    Select Case mType
    Case SN_ST_SNP3_FORWARD, SN_ST_SNP3_SUBSCRIPTION
        If mEnabled Then
            uReset

        Else
            If NOTNULL(theSocket) Then
                theSocket.CloseSocket
                Set theSocket = Nothing

            End If
        End If

    End Select

End Sub

Public Function IsEnabled() As Boolean

    IsEnabled = mEnabled

End Function

Public Function IsConnected() As Boolean

    If NOTNULL(theSocket) Then _
        IsConnected = (theSocket.State = sckConnected)

End Function

Public Sub Quit()

    If NOTNULL(theSocket) Then
        theSocket.CloseSocket
        Set theSocket = Nothing

    End If

End Sub

Private Sub uReset()

    Select Case mType

    Case SN_ST_SNP3_FORWARD
        If mEnabled Then
            g_Debug "TSubscriber.uReset(): SN_ST_SNP3_FORWARD: connecting to '" & mRemoteAddr & "'..."
            Set theSocket = New CSocket
            theSocket.Connect mRemoteAddr, SNP_DEFAULT_PORT

        End If

    Case SN_ST_GNTP_FORWARD
        g_Debug "TSubscriber.uReset(): GNTP forwarder not currently implemented", LEMON_LEVEL_WARNING
'        Set theSocket = New CSocket
'        theSocket.Connect mRemoteAddr, GNTP_DEFAULT_PORT

    Case SN_ST_SNP3_SUBSCRIPTION
        If mEnabled Then
            g_Debug "TSubscriber.uReset(): SN_ST_SNP3_SUBSCRIPTION: subscribing to '" & mRemoteAddr & "'..."
            Set theSocket = New CSocket
            theSocket.Connect mRemoteAddr, SNP_DEFAULT_PORT

        End If

    Case SN_ST_GNTP_SUBSCRIPTION
        g_Debug "TSubscriber.uReset(): GNTP subscriber not currently implemented", LEMON_LEVEL_WARNING
'        Set theSocket = New CSocket
'        theSocket.Connect mRemoteAddr, GNTP_DEFAULT_PORT

    Case SN_ST_GNTP_SUBSCRIBER, SN_ST_SNP3_SUBSCRIBER
        Set theSocket = Nothing

    End Select

End Sub

Private Function uGetFirstLine(ByVal str As String) As String
Dim i As Long

    i = InStr(str, vbCrLf)
    If i Then _
        uGetFirstLine = g_SafeLeftStr(str, i - 1)

End Function

Public Function RemoteHostEx() As String

    If NOTNULL(theSocket) Then _
        RemoteHostEx = theSocket.RemoteHostIP & ":" & CStr(theSocket.RemotePort)

End Function

Public Function RemoteHost() As String

    If NOTNULL(theSocket) Then
        RemoteHost = theSocket.RemoteHostIP

    Else
        RemoteHost = mRemoteAddr

    End If

End Function

Public Function SetTo2(ByVal Protocol, ByRef RemoteSocket As CSocket, ByRef Data As BPackedData) As Boolean

    g_Debug "TSubscriber.SetTo2()", LEMON_LEVEL_PROC_ENTER

    mProtocol = LCase$(Protocol)
    Set theSocket = RemoteSocket
    mAddressAndPort = RemoteSocket.RemoteHostIP & ":" & CStr(RemoteSocket.RemotePort)

    Select Case mProtocol
    Case "snp"
'        mGuid = g_CreateGUID(True)          ' // we'll create our own...
        SetTo2 = True

    Case Else
        g_Debug "Protocol " & g_Quote(mProtocol) & " not supported", LEMON_LEVEL_CRITICAL

    End Select


    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function Protocol() As String

    Protocol = mProtocol

End Function

Public Function AddressAndPort() As String

    AddressAndPort = mAddressAndPort

End Function
