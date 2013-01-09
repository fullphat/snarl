VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNetworkRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TSubscriberRoster.cls
    '/
    '/  Description:    Manages subscribed GNTP and SNP3 clients
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mRefs As Long

'Dim mListenerConfig As ConfigFile
'Dim mListeners As BTagList

Dim mConfig As ConfigFile
Dim mSubscribers As BTagList        ' // dynamic list of current subscribers
Dim mForwardsAndSubs As BTagList    ' // dynamic list of forwards and subscriptions in progress

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT
Dim ps As TSubscriber
Dim pf As TForward

    mRefs = mRefs - 1
    If mRefs = 0 Then
        ' /* run tidyup code */

        With mSubscribers
            .Rewind
            Do While .GetNextTag(ps) = B_OK
                ps.TidyUp

            Loop
        End With

        With mForwardsAndSubs
            .Rewind
            Do While .GetNextTag(pf) = B_OK
                pf.TidyUp

            Loop
        End With

    End If

End Function

Private Sub MLibrary_Initialize()

    ' /* not currently used */

End Sub

Private Function MLibrary_Magic() As Long

    MLibrary_Magic = &H23232323

End Function

Private Function MLibrary_Open() As melon.M_RESULT
Dim pc As ConfigSection
'Dim ps As TSubscriber
'Dim pl As TWebListener

    g_Debug "TSubscriberRoster.Open()", LEMON_LEVEL_PROC_ENTER

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        Set mSubscribers = new_BTagList()
        Set mForwardsAndSubs = new_BTagList()

        Set mConfig = New ConfigFile
        With mConfig
            .File = g_MakePath(gPrefs.SnarlConfigPath) & "etc\network.snarl"
            .Load

        End With

'        Set mListeners = new_BTagList()
'        Set mListenerConfig = New ConfigFile
'        With mListenerConfig
'            .File = g_SettingsPath() & "listeners.snarl"
'            .Load
'
'            ' /* load up persistent entries - specifically:
'            ' */
'
'            .Rewind
'
'            Do While .GetNextSection(pc)
'                Set pl = New TWebListener
'                If pl.SetFrom(pc) Then
'                    mListeners.Add pl
'                    g_Debug "got weblistener '" & pl.Name & "' url=" & pl.URL
'
'                End If
'            Loop
'
'        End With

Dim ps As TForward

        With mConfig
            .Rewind
            Do While .GetNextSection(pc)
                If pc.GetValueWithDefault("type", "") = "subscription" Then
                    Set ps = New TForward
                    If ps.SubscribeTo(pc.GetValueWithDefault("name", ""), pc.GetValueWithDefault("hosts", ""), pc.GetValueWithDefault("protocol", ""), pc.Name, pc.GetValueWithDefault("password", "")) Then _
                        mForwardsAndSubs.Add ps

                End If
            Loop
        End With

        g_Debug "roster initialised"

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()

    ' /* not currently used */

End Sub

Private Function MRoster_CountItems() As Long

    MRoster_CountItems = mSubscribers.CountItems '+ mGNTPSubscribers.CountItems

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long
Dim i As Long

'    If mCount Then
'        Name = LCase$(Name)
'        For i = 1 To mCount
'            If LCase$(mApp(i).Name) = Name Then
'                MRoster_FindItem = i
'                Exit Function
'
'            End If
'        Next i
'
'        g_Debug "TApplicationRoster.FindItem(): app '" & Name & "' not found", LEMON_LEVEL_WARNING
'
'    Else
'        g_Debug "TApplicationRoster.FindItem(): no apps registered!", LEMON_LEVEL_CRITICAL
'
'    End If

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.mObject

'    If (Index > 0) And (Index <= mCount) Then _
        Set MRoster_ItemAt = mApp(Index)

End Function

Private Function MRoster_NextItem(Item As melon.mObject) As Boolean

'    If mIndex <= mCount Then
'        Set Item = mApp(mIndex)
'        mIndex = mIndex + 1
'        MRoster_NextItem = True
'
'    End If

End Function

Private Sub MRoster_Rewind()

'    mIndex = 1

End Sub

Private Property Get MVersionInfo_Date() As String

    MVersionInfo_Date = "24-Jun-2011"           ' // creation date, not last modification date

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "subscriber.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

''Public Function Add(ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Socket As CSocket, ByRef Data As BPackedData) As Boolean
'Public Function Add(ByVal SubType As SN_SUBSCRIBER_TYPE, ByVal Description As String, ByVal RemoteHost As String, ByVal Apps As String, ByRef Error As String) As Boolean
'Dim pc As ConfigSection
'Dim ps As TSubscriber
'
'    g_Debug "TSubscriberRoster.Add(): type=0x" & g_HexStr(SubType, 4) & " dest=" & RemoteHost
'
'    Select Case SubType
'    Case SN_ST_SNP3_FORWARD, SN_ST_SNP3_SUBSCRIPTION
'        ' /*
'        '   This represents a persistent forward or subscription from this computer to a remote destination using SNP3
'        '
'        '   Description: optional
'        '   RemoteHost: required, specifies where to forward the notifications to
'        '   Apps: optional, identifies which app(s) to forward notifications for
'        '
'        '   On successful add: add to config and save
'        '
'        ' */
'
'        ' /* error checking */
'
'        If g_IsLocalAddress(RemoteHost) Then
'            g_Debug "TSubscriberRoster.Add(): RemoteHost cannot be ourself", LEMON_LEVEL_CRITICAL
'            Error = "The specified IP address is assigned to this computer"
'            Exit Function
'
'        End If
'
'        ' /* check this forward/subscription doesn't already exist */
'
'        If uFindByTypeAndDestination(SubType, RemoteHost, Nothing) Then
'            g_Debug "TSubscriberRoster.Add(): already exists", LEMON_LEVEL_CRITICAL
'            Error = "A " & IIf(SubType = SN_ST_SNP3_FORWARD, "forward", "subscription") & " to this computer already exists"
'            Exit Function
'
'        End If
'
'        ' /* create persistent config */
'
'        Set pc = mSubscriberConfig.AddSectionObj(g_CreateGUID(True))
'        With pc
'            .Add "type", CStr(SubType)
'            .Add "description", Description
'            .Add "remotehost", RemoteHost
'            .Add "apps", Apps
'            .Add "enabled", "1"
'
'        End With
'
'        ' /* create and initialise the subscriber */
'
'        Set ps = New TSubscriber
'        If Not ps.SetFrom(pc) Then
'            g_Debug "TSubscriberRoster.Add(): failed to initialise", LEMON_LEVEL_CRITICAL
'            mSubscriberConfig.RemoveSection mSubscriberConfig.FindSection(pc.Name)      ' // remove from the config
'            Exit Function
'
'        End If
'
'        ' /* success! so add it to the list */
'
'        mSubscribers.Add ps
'        mSubscriberConfig.Save
'        Add = True
'
'    End Select
'
''    ps.SetTo "", SubType, Description, RemoteHost, Apps, mConfig.AddSectionObj("")
'
'End Function

'Public Function AddSubscriber(ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef RemoteSocket As CSocket, ByRef Data As BPackedData) As Boolean
'
'    ' /* called in response to an incoming {subscribe} command */
'
'    If (RemoteSocket Is Nothing) Or (Data Is Nothing) Then
'        g_SetLastError SNARL_ERROR_SYSTEM
'        Exit Function
'
'    End If
'
'    If g_ConfigGet("allow_subs") = "0" Then
'        g_Debug "TSubscriberRoster.AddSubscriber(): not accepting subscriptions", LEMON_LEVEL_WARNING
'        g_SetLastError SNARL_ERROR_ACCESS_DENIED
'        Exit Function
'
'    End If
'
'    If (Not gDebugMode) And (g_IsLocalAddress(RemoteSocket.RemoteHostIP)) Then
'        g_Debug "TSubscriberRoster.AddSubscriber(): local subscriptions are not allowed", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_ACCESS_DENIED
'        Exit Function
'
'    End If
'
'
'Dim ps As TSubscriber
'
'    If uIndexOf(RemoteSocket.RemoteHostIP & ":" & CStr(RemoteSocket.RemotePort)) <> 0 Then
'        ' /* already subscribed */
'        g_Debug "TSubscriberRoster.AddSubscriber(): '" & RemoteSocket.RemoteHostIP & "' is already subscribed", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_ALREADY_SUBSCRIBED
'        Exit Function
'
'    End If
'
''        ' /* check password */
''        If ps.Password <> Data.ValueOf("password") Then
''            g_Debug "TSubscriberRoster.uAddNewSubscriber(): password mismatch", LEMON_LEVEL_CRITICAL
''            g_SetLastError SNARL_ERROR_AUTH_FAILURE
''
''        ElseIf Data.Exists("app-sig") Then
''            ' /* add specific apps */
''            uAddNewSubscriber = ps.AddApps(Data)
''
''        Else
''            ' /* cannot do this if already subscribed */
''            g_Debug "TSubscriberRoster.uAddNewSubscriber(): cannot do wildcard subscribe", LEMON_LEVEL_CRITICAL
''            g_SetLastError SNARL_ERROR_FAILED
''
''        End If
''
''    Else
'
'    ' /* not already subscribed */
'    Set ps = New TSubscriber
'    If ps.SetTo(SubscriberType, RemoteSocket, Data) Then
'        mSubscribers.Add ps
'        g_SetLastError SNARL_SUCCESS
'        AddSubscriber = True
'
'    Else
'        g_Debug "TSubscriberRoster.AddSubscriber(): failed to add subscriber from " & RemoteSocket.RemoteHostIP, LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_FAILED
'
'    End If
'
'End Function
'
'Public Function Remove(ByRef Socket As CSocket, ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Data As BPackedData) As Boolean
'
'    If (Socket Is Nothing) Or (Data Is Nothing) Then _
'        Exit Function
'
'Dim ps As TSubscriber
'Dim i As Long
'
'    i = uIndexOf(Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort))
'    If i Then
'        Debug.Print "TSubscriberRoster.Remove(): subscriber '" & Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort) & "' found"
'
'        Set ps = mSubscribers.TagAt(i)
'
'        If ps.SubscriptionType <> SubscriberType Then
'            ' /* wrong type */
'            g_Debug "TSubscriberRoster.Remove(): '" & Socket.RemoteHostIP & "' type mismatch", LEMON_LEVEL_CRITICAL
'            g_SetLastError SNARL_ERROR_INVALID_ARG
'
'        ElseIf ps.Password <> Data.ValueOf("password") Then
'            ' /* wrong password */
'            g_Debug "TSubscriberRoster.Remove(): password mismatch", LEMON_LEVEL_CRITICAL
'            g_SetLastError SNARL_ERROR_AUTH_FAILURE
'
'        Else
'            ' /* remove */
'            uRemove i
'            Remove = True
'
'        End If
'
'    Else
'        g_Debug "TSubscriberRoster.Remove(): '" & Socket.RemoteHostIP & "' is not subscribed", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_NOT_SUBSCRIBED
'
'    End If
'
'End Function

'Private Function uIndexOf(ByVal IPAddressAndPort As String) As Long
'Dim ps As TSubscriber
'Static i As Long
'
'    If mSubscribers.CountItems Then
'        For i = 1 To mSubscribers.CountItems
'            Set ps = mSubscribers.TagAt(i)
'            If ps.RemoteHostEx = IPAddressAndPort Then
'                uIndexOf = i
'                Exit Function
'
'            End If
'        Next i
'    End If
'
'End Function

'Friend Sub Broadcast(ByRef Content As T_NOTIFICATION_INFO)
'
'    If mSubscribers.CountItems = 0 Then _
'        Exit Sub
'
'Dim szRequest As String
'Dim ps As TSubscriber
'Dim pa As TApp
'
'    szRequest = uCreateSNP3Forward(Content)
'    Set pa = Content.ClassObj.App
'
'    With mSubscribers
'        .Rewind
'        Do While .GetNextTag(ps) = B_OK
'            Debug.Print "TSubscriberRoster.Broadcast(): '" & ps.Destination & "' enabled=" & ps.IsEnabled
'
'            If ps.IsEnabled Then
'                ' /* TO-DO: if the subscriber uses GNTP we will need to do
'                '    a two-pass process.  Step 1 will do the REGISTER;
'                '    STEP 2 will do the NOTIFY */
'                If ps.IsSubscribedTo(pa.Signature) Then _
'                    ps.send szRequest
'
'            End If
'        Loop
'
'    End With
'
'End Sub

Public Function Subscribers() As BTagList

    Set Subscribers = mSubscribers

End Function

'Public Function RemoveByGuid(ByVal Guid As String) As Boolean
'Dim i As Long
'
'    i = mSubscribers.IndexOf(Guid)
'    If i <> 0 Then
'        uRemove i
'        RemoveByGuid = True
'
'    Else
'        g_Debug "TSubscriberRoster.RemoveByGuid(): '" & Guid & "' not in roster", LEMON_LEVEL_WARNING
'
'    End If
'
'    ' /* only persistent forwarders and subscriptions will have a config entry */
'
'    i = mSubscriberConfig.FindSection(Guid)
'    If i <> 0 Then
'        g_Debug "TSubscriberRoster.RemoveByGuid(): updating settings..."
'        mSubscriberConfig.RemoveSection i
'        mSubscriberConfig.Save
'
'    Else
'        g_Debug "TSubscriberRoster.RemoveByGuid(): '" & Guid & "' not in config", LEMON_LEVEL_WARNING
''        f = False
'
'    End If
'
'End Function

'Private Sub uRemove(ByVal Index As Long)
'Dim ps As TSubscriber
'
'    Set ps = mSubscribers.TagAt(Index)
'    If ISNULL(ps) Then _
'        Exit Sub
'
''    If (ps.SubscriptionType = SN_ST_GNTP_SUBSCRIBER) Or (ps.SubscriptionType = SN_ST_SNP3_SUBSCRIBER) Then _
'        g_PrivateNotify "", "Subscriber removed", ps.RemoteHost & " unsubscribed", , ".sub-snp-rem"
'
'    g_Debug "TSubscriberRoster.uRemove(): '" & ps.RemoteHost & "' removed"
'    mSubscribers.Remove Index
'
'End Sub

'Public Function RemoveListenerByGuid(ByVal Guid As String) As Boolean
'Dim i As Long
'
'    i = mListeners.IndexOf(Guid)
'    If i = 0 Then
'        g_Debug "TSubscriberRoster.RemoveListenerByGuid(): '" & Guid & "' not in roster", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End If
'
'    mListeners.Remove i
'
'    i = mListenerConfig.FindSection(Guid)
'    If i <> 0 Then
'        mListenerConfig.RemoveSection i
'        mListenerConfig.Save
'
'    Else
'        g_Debug "TSubscriberRoster.RemoveListenerByGuid(): '" & Guid & "' not in config", LEMON_LEVEL_WARNING
'
'    End If
'
'    RemoveListenerByGuid = True
'
'End Function
'
'Public Function AddWebListener(ByVal Description As String, ByVal URL As String) As Boolean
'Dim pc As ConfigSection
'Dim pl As TWebListener
'
'    g_Debug "TSubscriberRoster.AddWebListener(): name=" & Description & " URL=" & URL
'
'    ' /* create and initialise the subscriber */
'
'    Set pl = New TWebListener
'    If Not pl.CreateNew(Description, URL) Then
'        g_Debug "TSubscriberRoster.AddWebListener(): failed to initialise", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End If
'
'    mListeners.Add pl
'
'    ' /* create persistent config */
'
''    Set pc = mListenerConfig.AddSectionObj(pl.Guid)
''    With pc
''        .Add "type", CStr(SN_ST_COMET)                  ' // for now, in case we support other models in the future
''        .Add "name", Description
''        .Add "url", URL
''        .Add "enabled", "1"
''
''    End With
''
''    mListenerConfig.Save
''    AddWebListener = True
'
'End Function

Friend Function QuickForward(ByVal ForwarderGuid As String, ByRef Info As T_NOTIFICATION_INFO) As Boolean
Dim i As Long

    i = mConfig.FindSection(ForwarderGuid)
    If i Then
        QuickForward = Me.ForwardNotification(mConfig.SectionAt(i), Info)

    Else
        g_Debug "TNetworkRoster.QuickForward(): forwarder " & g_Quote(ForwarderGuid) & " not found", LEMON_LEVEL_CRITICAL

    End If

End Function

Public Function Remove(ByVal Guid As String) As Boolean
Dim pf As TForward
Dim i As Long

    ' /* removes the specified forwarder or subscription from network.snarl config file */

    i = mConfig.FindSection(Guid)
    If i Then
        mConfig.RemoveSection i
        mConfig.Save
        Remove = True

    End If

    ' /* remove the actual subscription object from the roster */

    i = mForwardsAndSubs.IndexOf(Guid)
    If i Then
        Set pf = mForwardsAndSubs.TagAt(i)
        If pf.IsSubscription Then _
            pf.Unsubscribe

        mForwardsAndSubs.Remove i

    End If

End Function

Public Function Find(ByVal Name As String, ByRef Forwarder As ConfigSection) As Boolean
Dim i As Long

    With mConfig
        i = .FindSection(Name)
        If i Then
            Set Forwarder = .SectionAt(i)
            Find = True

        End If
    End With

End Function

Public Function Config() As ConfigFile

    Set Config = mConfig

End Function

Public Sub BuildForwarderList(ByRef Control As BControl, Optional ByRef Existing As BTagList)
Dim n As Long

    n = Val(Control.GetValue())

Dim ps As ConfigSection
Dim szList As String
Dim szProto As String

    With g_SubsRoster.Config
        .Load
        .Rewind
        Do While .GetNextSection(ps)
            If ps.GetValueWithDefault("type", "") = "forwarder" Then
                ' /* name, hosts, protocol */
                If (ps.GetValueWithDefault("name", "") <> "") And (ps.Name <> "") Then
                    Select Case LCase$(ps.GetValueWithDefault("protocol", ""))
                    Case "snp"
                        szProto = "SNP"

                    Case "gntp"
                        szProto = "GNTP"
                        
                    Case "smb"
                        szProto = "Server Message Block (SMB)"

                    Case Else
                        szProto = "<unknown protocol>"

                    End Select

                    szList = szList & ps.GetValueWithDefault("name", "") & "#?" & ps.Name & "#?" & _
                                      ps.GetValueWithDefault("hosts", "<none>") & " using " & szProto & "|"

                End If
            End If
        Loop

        szList = g_SafeLeftStr(szList, Len(szList) - 1)
        Control.SetText szList

Dim szIcon As String
Dim i As Long

        If szList <> "" Then
            .Rewind
            Do While .GetNextSection(ps)
                If ps.GetValueWithDefault("type", "") = "forwarder" Then
                    If (ps.GetValueWithDefault("name", "") <> "") And (ps.Name <> "") Then
                        i = i + 1
                        szIcon = LCase$(ps.GetValueWithDefault("protocol", ""))
                        Select Case szIcon
                        Case "snp", "gntp", "smb"
                        
                        Case Else
                            szIcon = "unknown"
                            
                        End Select
    
                        prefskit_SetItemObject Control, i, "image-object", load_image_obj(g_MakePath(App.Path) & "etc\icons\protocol-" & szIcon & ".png")
    
                        If NOTNULL(Existing) Then
                            If Existing.IndexOf(ps.Name) <> 0 Then _
                                prefskit_SetItem Control, i, "checked", 1&
    
                        End If
                    End If
                End If
            Loop
        End If
    End With

    Control.SetValue CStr(Max(n, 1))

End Sub

Public Sub BuildSubscriptionList(ByRef Control As BControl, Optional ByRef Existing As BTagList)
Dim n As Long

    n = Val(Control.GetValue())

Dim ps As ConfigSection
Dim szList As String
Dim szProto As String

    With g_SubsRoster.Config
        .Load
        .Rewind
        Do While .GetNextSection(ps)
            If ps.GetValueWithDefault("type", "") = "subscription" Then
                ' /* name, host, protocol */
                If (ps.GetValueWithDefault("name", "") <> "") And (ps.Name <> "") Then
                    Select Case LCase$(ps.GetValueWithDefault("protocol", ""))
                    Case "snp"
                        szProto = "SNP"

                    Case "gntp"
                        szProto = "GNTP"
                        
                    Case "smb"
                        szProto = "Server Message Block (SMB)"

                    Case Else
                        szProto = "<unknown protocol>"

                    End Select

                    szList = szList & ps.GetValueWithDefault("name", "") & "#?" & ps.Name & "#?" & _
                                      ps.GetValueWithDefault("hosts", "<none>") & " using " & szProto & "|"

                End If
            End If
        Loop

        szList = g_SafeLeftStr(szList, Len(szList) - 1)
        Control.SetText szList

Dim pf As TForward
Dim szIcon As String
Dim i As Long

        If szList <> "" Then
            .Rewind
            Do While .GetNextSection(ps)
                If ps.GetValueWithDefault("type", "") = "subscription" Then
                    If (ps.GetValueWithDefault("name", "") <> "") And (ps.Name <> "") Then
                        i = i + 1
                        szIcon = LCase$(ps.GetValueWithDefault("protocol", ""))
                        Select Case szIcon
                        Case "snp", "gntp", "smb"
                        
                        Case Else
                            szIcon = "unknown"
                            
                        End Select
    
                        prefskit_SetItemObject Control, i, "image-object", load_image_obj(g_MakePath(App.Path) & "etc\icons\protocol-" & szIcon & ".png")
                        If mForwardsAndSubs.Find(ps.Name, pf) Then
                            prefskit_SetItem Control, i, "greyscale", IIf(pf.IsConnected, 0&, -1&)

                        Else
                            prefskit_SetItem Control, i, "greyscale", -1&
    
                        End If

                        If NOTNULL(Existing) Then
                            If Existing.IndexOf(ps.Name) <> 0 Then _
                                prefskit_SetItem Control, i, "checked", 1&
    
                        End If
                    End If
                End If
            Loop
        End If
    End With

    Control.SetValue CStr(Max(n, 1))

End Sub

Friend Sub BroadcastToSubscribers(ByRef Content As T_NOTIFICATION_INFO)

    ' /* send Content to all subscribed computers in the same way we would
    '    forward a notification to them */

Dim ps As TSubscriber

    With mSubscribers
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            ps.Forward Content

        Loop

    End With

End Sub

Public Function AddSubscriber(ByVal Protocol As String, ByRef RemoteSocket As CSocket, ByRef Data As BPackedData) As SNARL_STATUS_CODE

    ' /* called in response to an incoming {subscribe} command */

    g_Debug "TSubscriberRoster.AddSubscriber()", LEMON_LEVEL_PROC_ENTER

    If (RemoteSocket Is Nothing) Or (Data Is Nothing) Then
        g_Debug "missing arg", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        AddSubscriber = SNARL_ERROR_SYSTEM
        Exit Function

    End If

    If g_ConfigGet("allow_subs") = "0" Then
        g_Debug "sorry, we're not accepting subscriptions", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        AddSubscriber = SNARL_ERROR_ACCESS_DENIED
        Exit Function

    End If

'    If (Not gDebugMode) And (g_IsLocalAddress(RemoteSocket.RemoteHostIP)) Then
'        g_Debug "TSubscriberRoster.AddSubscriber(): local subscriptions are not allowed", LEMON_LEVEL_CRITICAL
'        g_SetLastError SNARL_ERROR_ACCESS_DENIED
'        Exit Function
'
'    End If

Dim szDestination As String

    If Protocol = "gntp" Then
        szDestination = RemoteSocket.RemoteHostIP & ":" & Data.ValueOf("reply-port")

    Else
        szDestination = RemoteSocket.RemoteHostIP & ":" & CStr(RemoteSocket.RemotePort)

    End If

Dim ps As TSubscriber

    If Not mSubscribers.Find(szDestination, ps) Then
        ' /* not already subscribed */
        Set ps = New TSubscriber
        If ps.SetTo(Protocol, RemoteSocket, Data) Then
            ' /* success */
            mSubscribers.Add ps
            ' /* refresh list */
            frmAbout.bSubsChanged
            AddSubscriber = SNARL_SUCCESS

        Else
            g_Debug "failed to add subscriber " & g_Quote(szDestination), LEMON_LEVEL_CRITICAL
            AddSubscriber = SNARL_ERROR_FAILED

        End If
    Else
        ' /* already subscribed */
        g_Debug g_Quote(szDestination) & " is already subscribed", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        ps.Bump
        AddSubscriber = SNARL_ERROR_ALREADY_SUBSCRIBED

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function RemoveSubscriber(ByRef Socket As CSocket, Optional ByRef Data As BPackedData) As Boolean

    g_Debug "TSubscriberRoster.RemoveSubscriber()", LEMON_LEVEL_PROC_ENTER

    If (Socket Is Nothing) Then
        g_Debug "missing arg", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        g_SetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

Dim szDest As String
Dim i As Long

    szDest = Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort)

'MsgBox "removing subscriber: " & szDest

    i = mSubscribers.IndexOf(szDest)
    If i Then
        g_Debug g_Quote(szDest) & " removed", LEMON_LEVEL_INFO
        mSubscribers.Remove i
        frmAbout.bSubsChanged
        RemoveSubscriber = True

    Else
        g_Debug g_Quote(szDest) & " is not subscribed", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_NOT_SUBSCRIBED

    End If

'        ElseIf ps.Password <> Data.ValueOf("password") Then
'            ' /* wrong password */
'            g_Debug "TSubscriberRoster.Remove(): password mismatch", LEMON_LEVEL_CRITICAL
'            g_SetLastError SNARL_ERROR_AUTH_FAILURE

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function IsSubscriber(ByVal AddressAndPort As String) As Boolean

    IsSubscriber = (mSubscribers.IndexOf(AddressAndPort) <> 0)

End Function

Public Function IsSNPSubscriber(ByVal AddressAndPort As String) As Boolean
Dim ps As TSubscriber

    If mSubscribers.Find(AddressAndPort, ps) Then _
        IsSNPSubscriber = (ps.Protocol = "snp")

End Function


    ' /* forwards *

Public Sub AddForwardOrSubscription(ByRef Forwarder As TForward)

    mForwardsAndSubs.Add Forwarder

End Sub

Public Sub RemoveForwardOrSubscription(ByVal Guid As String)

    mForwardsAndSubs.Remove mForwardsAndSubs.IndexOf(Guid)

End Sub

Friend Function ForwardNotification(ByRef Forwarder As ConfigSection, ByRef Info As T_NOTIFICATION_INFO) As Boolean
Dim szp As String

    szp = LCase$(Forwarder.GetValueWithDefault("protocol", ""))
    If (szp <> "snp") And (szp <> "gntp") And (szp <> "smb") Then
        g_Debug "TSubscriberRoster.ForwardNotification(): protocol " & g_Quote(szp) & " not supported", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo Forwarder.GetValueWithDefault("hosts", "")
    If ppd.Count = 0 Then
        g_Debug "TSubscriberRoster.ForwardNotification(): no hosts defined", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pf As TForward
Dim szn As String

    ' /* each defined forwarder can have multiple remote destinations, here we
    '    send the notification content to each destination asynchronously */

    With ppd
        .Rewind
        Do While .GetNextItem(szn, "")
            Set pf = New TForward
            pf.ForwardTo Info, szn, szp
            mForwardsAndSubs.Add pf

        Loop
    End With

    ForwardNotification = True

End Function

Public Function AddSubscription(ByVal Name As String, ByVal Destination As String, ByVal Protocol As String, ByVal Guid As String, ByVal Password As String)
Dim ps As TForward

    Set ps = New TForward
    If ps.SubscribeTo(Name, Destination, Protocol, Guid, Password) Then _
        mForwardsAndSubs.Add ps

End Function
