VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSubscriberRoster"
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

Dim mSubscriberConfig As ConfigFile
Dim mRefs As Long

Dim mListenerConfig As ConfigFile
Dim mListeners As BTagList

    ' /* list of defined forwards: held in network.config
Dim mForwards As BTagList
Dim mForwarders As ConfigFile

    ' /* dynamic list of current subscribers */
Dim mSubscribers As BTagList


Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT
Dim ps As TSubscriber

    mRefs = mRefs - 1
    If mRefs = 0 Then
        ' /* run tidyup code */

        With mSubscribers
            .Rewind
            Do While .GetNextTag(ps) = B_OK
                ps.Quit

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
Dim ps As TSubscriber
Dim pl As TWebListener

    g_Debug "TSubscriberRoster.Open()", LEMON_LEVEL_PROC_ENTER

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        Set mSubscribers = new_BTagList()
'        Set mGNTPSubscribers = new_BTagList()

'        Set mSubscriberConfig = New ConfigFile
'        With mSubscriberConfig
'            .File = g_SettingsPath() & "network.snarl"
'            .Load
'
'            ' /* load up persistent entries - specifically:
'            '       forwards from this computer to a remote computer
'            '       subscriptions to a remote computer
'            ' */
'
'            .Rewind
'
'            Do While .GetNextSection(pc)
'                Set ps = New TSubscriber
'                If ps.SetFrom(pc) Then
'                    mSubscribers.Add ps
'                    g_Debug "got " & ps.TypeStr & " to " & ps.Destination
'
'                End If
'            Loop
'
'        End With

        Set mListeners = new_BTagList()
        Set mListenerConfig = New ConfigFile
        With mListenerConfig
            .File = g_SettingsPath() & "listeners.snarl"
            .Load

            ' /* load up persistent entries - specifically:
            ' */

            .Rewind

            Do While .GetNextSection(pc)
                Set pl = New TWebListener
                If pl.SetFrom(pc) Then
                    mListeners.Add pl
                    g_Debug "got weblistener '" & pl.Name & "' url=" & pl.URL

                End If
            Loop

        End With

        Set mForwards = new_BTagList()
        Set mForwarders = New ConfigFile
        With mForwarders
            .File = g_MakePath(gPrefs.SnarlConfigPath) & "etc\network.snarl"
            .Load

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
'        gSetLastError SNARL_ERROR_SYSTEM
'        Exit Function
'
'    End If
'
'    If g_ConfigGet("allow_subs") = "0" Then
'        g_Debug "TSubscriberRoster.AddSubscriber(): not accepting subscriptions", LEMON_LEVEL_WARNING
'        gSetLastError SNARL_ERROR_ACCESS_DENIED
'        Exit Function
'
'    End If
'
'    If (Not gDebugMode) And (g_IsLocalAddress(RemoteSocket.RemoteHostIP)) Then
'        g_Debug "TSubscriberRoster.AddSubscriber(): local subscriptions are not allowed", LEMON_LEVEL_CRITICAL
'        gSetLastError SNARL_ERROR_ACCESS_DENIED
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
'        gSetLastError SNARL_ERROR_ALREADY_SUBSCRIBED
'        Exit Function
'
'    End If
'
''        ' /* check password */
''        If ps.Password <> Data.ValueOf("password") Then
''            g_Debug "TSubscriberRoster.uAddNewSubscriber(): password mismatch", LEMON_LEVEL_CRITICAL
''            gSetLastError SNARL_ERROR_AUTH_FAILURE
''
''        ElseIf Data.Exists("app-sig") Then
''            ' /* add specific apps */
''            uAddNewSubscriber = ps.AddApps(Data)
''
''        Else
''            ' /* cannot do this if already subscribed */
''            g_Debug "TSubscriberRoster.uAddNewSubscriber(): cannot do wildcard subscribe", LEMON_LEVEL_CRITICAL
''            gSetLastError SNARL_ERROR_FAILED
''
''        End If
''
''    Else
'
'    ' /* not already subscribed */
'    Set ps = New TSubscriber
'    If ps.SetTo(SubscriberType, RemoteSocket, Data) Then
'        mSubscribers.Add ps
'        gSetLastError SNARL_SUCCESS
'        AddSubscriber = True
'
'    Else
'        g_Debug "TSubscriberRoster.AddSubscriber(): failed to add subscriber from " & RemoteSocket.RemoteHostIP, LEMON_LEVEL_CRITICAL
'        gSetLastError SNARL_ERROR_FAILED
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
'            gSetLastError SNARL_ERROR_INVALID_ARG
'
'        ElseIf ps.Password <> Data.ValueOf("password") Then
'            ' /* wrong password */
'            g_Debug "TSubscriberRoster.Remove(): password mismatch", LEMON_LEVEL_CRITICAL
'            gSetLastError SNARL_ERROR_AUTH_FAILURE
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
'        gSetLastError SNARL_ERROR_NOT_SUBSCRIBED
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

Public Sub WriteConfig()

    If Not (mSubscriberConfig Is Nothing) Then _
        mSubscriberConfig.Save


End Sub

Private Function uFindByTypeAndDestination(ByVal aType As SN_SUBSCRIBER_TYPE, ByVal Destination As String, ByRef Subscriber As TSubscriber) As Boolean
Dim ps As TSubscriber

    With mSubscribers
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            If (ps.SubscriptionType = aType) And (ps.Destination = Destination) Then
                Set Subscriber = ps
                uFindByTypeAndDestination = True
                Exit Function

            End If

        Loop
    End With

End Function

'Private Function uTranslateIcon(ByVal Icon As String) As String
'
'    If Icon = "" Then _
'        Exit Function
'
'Dim sz As String
'Dim b As String
'
'    If (g_SafeLeftStr(Icon, 1) = "!") Or (g_IsURL(Icon)) Then
'        ' /* add verbatim */
'        sz = sz & "&icon=" & Icon
'
'    Else
'        ' /* encode it in a slightly modified Base64 format (CRLF's are replaced with #'s) */
'        If uEncodeIcon(Icon, b) Then _
'            sz = sz & "&icon-phat64=" & b
'
'    End If
'
'    uTranslateIcon = sz
'
'End Function
'
'Private Function uEncodeIcon(ByVal IconPath As String, ByRef Base64 As String) As Boolean
'
'    If IconPath = "" Then _
'        Exit Function
'
'    If Not g_Exists(IconPath) Then _
'        Exit Function
'
'Dim sz As String
'Dim i As Integer
'
'    On Error Resume Next
'
'    i = FreeFile()
'
'    err.Clear
'    Open IconPath For Binary Access Read Lock Write As #i
'    If err.Number = 0 Then
'        sz = String$(LOF(i), Chr$(0))
'        Get #i, , sz
'        Close #i
'
'        sz = Encode64orig(sz)                   ' // encode as standard Base64
'        If sz <> "" Then
'            Base64 = Replace$(sz, vbCrLf, "#")  ' // replace CRLFs
'            Base64 = Replace$(Base64, "=", "%") ' // replace ='s
'            uEncodeIcon = True
'
'        End If
'
'    Else
'        g_Debug "TSubscriberRoster.uEncodeIcon(): " & err.Description, LEMON_LEVEL_CRITICAL
'
'    End If
'
'End Function

Public Function Listeners() As BTagList

    Set Listeners = mListeners

End Function

Public Function RemoveListenerByGuid(ByVal Guid As String) As Boolean
Dim i As Long

    i = mListeners.IndexOf(Guid)
    If i = 0 Then
        g_Debug "TSubscriberRoster.RemoveListenerByGuid(): '" & Guid & "' not in roster", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    mListeners.Remove i

    i = mListenerConfig.FindSection(Guid)
    If i <> 0 Then
        mListenerConfig.RemoveSection i
        mListenerConfig.Save

    Else
        g_Debug "TSubscriberRoster.RemoveListenerByGuid(): '" & Guid & "' not in config", LEMON_LEVEL_WARNING

    End If

    RemoveListenerByGuid = True

End Function

Public Function AddWebListener(ByVal Description As String, ByVal URL As String) As Boolean
Dim pc As ConfigSection
Dim pl As TWebListener

    g_Debug "TSubscriberRoster.AddWebListener(): name=" & Description & " URL=" & URL

    ' /* create and initialise the subscriber */

    Set pl = New TWebListener
    If Not pl.CreateNew(Description, URL) Then
        g_Debug "TSubscriberRoster.AddWebListener(): failed to initialise", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    mListeners.Add pl

    ' /* create persistent config */

    Set pc = mListenerConfig.AddSectionObj(pl.Guid)
    With pc
        .Add "type", CStr(SN_ST_COMET)                  ' // for now, in case we support other models in the future
        .Add "name", Description
        .Add "url", URL
        .Add "enabled", "1"

    End With

    mListenerConfig.Save
    AddWebListener = True

End Function

'Private Function uCreateSNP3Forward(ByRef Content As T_NOTIFICATION_INFO) As String
'Dim szc As String
'Dim pa As TApp
'
'    Set pa = Content.ClassObj.App
'
'    ' /* base content */
'
'    szc = "SNP/3.0" & vbCrLf & _
'          "register?app-sig=" & pa.Signature & "&title=" & pa.Name & uTranslateIcon(pa.Icon) & vbCrLf
'
'    ' /* add classes */
'
'Dim i As Long
'
'    With pa
'        If .CountAlerts Then
'            For i = 1 To .CountAlerts
'                With .AlertAt(i)
'                    szc = szc & "addclass?app-sig=" & pa.Signature & "&id=" & .Name & "&name=" & .Description & vbCrLf
'
'                End With
'            Next i
'        End If
'    End With
'
'    ' /* build notification content */
'
'    szc = szc & "notify?app-sig=" & pa.Signature
'
'Dim szn As String
'Dim szv As String
'Dim sz As String
'
'    With New BPackedData
'        If .SetTo(Content.OriginalContent) Then
'            .Rewind
'            Do While .GetNextItem(szn, szv)
'                Select Case szn
'                Case "icon"
'                    sz = sz & "&" & uTranslateIcon(szv)
'
'                Case Else
'                    sz = sz & "&" & szn & "=" & szv
'
'                End Select
'            Loop
'        End If
'    End With
'
'    If sz <> "" Then
'        ' /* add notification content and footer */
'        sz = Replace$(sz, vbCrLf, "\n")
'        szc = szc & sz
'
'    End If
'
'    szc = szc & vbCrLf & "END" & vbCrLf
'
'        Debug.Print "---uCreateSNP3Forward--- (" & mSubscribers.CountItems & ")"
'        Debug.Print szc
'        Debug.Print "---"
'
'    uCreateSNP3Forward = szc
'
'End Function
'
'Private Function uCreateGNTPForward(ByRef Content As T_NOTIFICATION_INFO) As String
'Dim szc As String
'Dim sz As String
'
'    sz = Content.OriginalContent
'    sz = Replace$(sz, "::", "=")
'    sz = Replace$(sz, "#?", "&")
'    sz = Replace$(sz, vbCrLf, "\n")
'
'Dim pa As TApp
'
'    Set pa = Content.ClassObj.App
'
'    ' /* base content */
'
'    szc = "GNTP/1.0 REGISTER NONE" & vbCrLf & _
'          "Application-Name: " & pa.Name & vbCrLf & _
'          "Notifications-Count: " & CStr(pa.CountAlerts) & vbCrLf & vbCrLf
'
'    ' /* add classes */
'
'Dim i As Long
'
'    With pa
'        If .CountAlerts Then
'            For i = 1 To .CountAlerts
'                With .AlertAt(i)
'                    szc = szc & "Notification-Name: " & .Name & vbCrLf & _
'                                "Notification-Display-Name: " & .Description & vbCrLf & _
'                                "Notification-Enabled: " & IIf(.IsEnabled, "True", "False") & vbCrLf & vbCrLf
'
'                End With
'            Next i
'        End If
'    End With
'
'    ' /* add notification content and footer */
'
'    szc = szc & vbCrLf & vbCrLf
'
'    uCreateGNTPForward = szc
'
'End Function
'
'Private Function uCreateGNTPNotify(ByRef Content As T_NOTIFICATION_INFO) As String
'Dim szc As String
'Dim sz As String
'
'    sz = Content.OriginalContent
'    sz = Replace$(sz, "::", "=")
'    sz = Replace$(sz, "#?", "&")
'    sz = Replace$(sz, vbCrLf, "\n")
'
'Dim pa As TApp
'
'    Set pa = Content.ClassObj.App
'
'    ' /* base content */
'
'    szc = "GNTP/1.0 NOTIFY NONE" & vbCrLf & _
'          "Application-Name: " & pa.Name & vbCrLf
'
'
''Notification-Name: <string>
''Required - The name (type) of the notification (must match a previously registered notification name registered by the application specified in Application-Name)
''
''Notification-ID: <string>
''Optional - A unique ID for the notification. If used, this should be unique for every request, even if the notification is replacing a current notification (see Notification-Coalescing-ID)
''
''Notification-Title: <string>
''Required - The notification's title
''
''Notification-Text: <string>
''Optional - The notification's text. (defaults to "")
''
''Notification-Sticky: <boolean>
''Optional - Indicates if the notification should remain displayed until dismissed by the user. (default to False)
''
''Notification-Priority: <int>
''Optional - A higher number indicates a higher priority. This is a display hint for the receiver which may be ignored. (valid values are between -2 and 2, defaults to 0)
''
''Notification-Icon: <url> | <uniqueid>
''Optional - The icon to display with the notification.
''
''Notification-Coalescing-ID: <string>
''Optional - If present, should contain the value of the Notification-ID header of a previously-sent notification. This serves as a hint to the notification system that this notification should replace/update the matching previous notification. The notification system may ignore this hint.
''
''Notification-Callback-Context: <string>
''Optional - Any data (will be passed back in the callback unmodified)
''
''Notification-Callback-Context-Type: <string>
''Optional, but Required if 'Notification-Callback-Context' is passed - The type of data being passed in Notification-Callback-Context (will be passed back in the callback unmodified). This does not need to be of any pre-defined type, it is only a convenience to the sending application.
''
''Notification-Callback-Target: <string>
''Optional - An alternate target for callbacks from this notification. If passed, the standard behavior of performing the callback over the original socket will be ignored and the callback data will be passed
'
'
'    ' /* end marker */
'
'    szc = szc & vbCrLf & vbCrLf
'
'    uCreateGNTPNotify = szc
'
'End Function

Friend Function DoForward(ByRef Forwarder As ConfigSection, ByRef Info As T_NOTIFICATION_INFO) As Boolean
Dim szp As String

    szp = LCase$(Forwarder.GetValueWithDefault("protocol", ""))
    If (szp <> "snp") And (szp <> "gntp") And (szp <> "smb") Then
        g_Debug "TSubscriberRoster.DoForward(): protocol " & g_Quote(szp) & " not supported", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo Forwarder.GetValueWithDefault("hosts", "")
    If ppd.Count = 0 Then
        g_Debug "TSubscriberRoster.DoForward(): no hosts defined", LEMON_LEVEL_CRITICAL
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
            pf.Forward Info.ClassObj.App, Info, szn, szp
            mForwards.Add pf

        Loop
    End With

    DoForward = True

End Function

Public Sub RemoveForward(ByVal Guid As String)

    mForwards.Remove mForwards.IndexOf(Guid)

End Sub

Public Function RemoveForwarder(ByVal Guid As String) As Boolean
Dim i As Long

    ' /* removes the specified forwarder from network.snarl config file */

    i = mForwarders.FindSection(Guid)
    If i Then
        mForwarders.RemoveSection i
        mForwarders.Save
        RemoveForwarder = True

    End If

End Function

Public Function Forwarders() As ConfigFile

    Set Forwarders = mForwarders

End Function

Public Function FindForwarder2(ByVal Name As String, ByRef Forwarder As ConfigSection) As Boolean
Dim i As Long

    With mForwarders
        i = .FindSection(Name)
        If i Then
            Set Forwarder = .SectionAt(i)
            FindForwarder2 = True

        End If
    End With

End Function

Public Sub BuildForwarderList(ByRef Control As BControl, Optional ByRef Existing As BTagList)
Dim n As Long

    n = Val(Control.GetValue())

Dim ps As ConfigSection
Dim szList As String
Dim szProto As String

    With g_SubsRoster.Forwarders
        .Load
        .Rewind
        Do While .GetNextSection(ps)
            If ps.GetValueWithDefault("type", "") = "forwarder" Then
                ' /* name, hosts, protocol */
                If (ps.GetValueWithDefault("name", "") <> "") And (ps.Name <> "") Then
                    Select Case LCase$(ps.GetValueWithDefault("protocol", ""))
                    Case "snp"
                        szProto = "SNP 3.0"

                    Case "gntp"
                        szProto = "GNTP"
                        
                    Case "smb"
                        szProto = "Server Message Block (SMB)"

                    Case Else
                        szProto = "<unknown protocol>"

                    End Select

                    szList = szList & ps.GetValueWithDefault("name", "") & "#?" & ps.Name & "#?" & _
                                      "To " & ps.GetValueWithDefault("hosts", "<none>") & " using " & szProto & "|"

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
    
                        prefskit_SetItem Control, i, "image-file", g_MakePath(App.Path) & "etc\icons\protocol-" & szIcon & ".png"
    
                        If NOTNULL(Existing) Then
                            If Existing.IndexOf(ps.Name) <> 0 Then _
                                prefskit_SetItem Control, i, "checked", 1&
    
                        End If
                    End If
                End If
            Loop
        End If
    End With

    Control.SetValue CStr(MAX(n, 1))

End Sub

Friend Sub Broadcast(ByRef Content As T_NOTIFICATION_INFO)

    ' /* send Content to all subscribed computers in the same was as we would
    '    forward a notification to them */

    ' /* obviously, each subscriber can only have one destination address but
    '    the protocol can vary */



End Sub

Public Function AddSubscriber2(ByVal Protocol, ByRef RemoteSocket As CSocket, ByRef Data As BPackedData) As Boolean

    ' /* called in response to an incoming {subscribe} command */

    g_Debug "TSubscriberRoster.AddSubscriber2()", LEMON_LEVEL_PROC_ENTER

    If (RemoteSocket Is Nothing) Or (Data Is Nothing) Then
        g_Debug "missing RemoteSocket or Data", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        gSetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

    If g_ConfigGet("allow_subs") = "0" Then
        g_Debug "sorry, we're not accepting subscriptions", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        gSetLastError SNARL_ERROR_ACCESS_DENIED
        Exit Function

    End If

'    If (Not gDebugMode) And (g_IsLocalAddress(RemoteSocket.RemoteHostIP)) Then
'        g_Debug "TSubscriberRoster.AddSubscriber(): local subscriptions are not allowed", LEMON_LEVEL_CRITICAL
'        gSetLastError SNARL_ERROR_ACCESS_DENIED
'        Exit Function
'
'    End If

Dim szDestination As String

    szDestination = RemoteSocket.RemoteHostIP & ":" & CStr(RemoteSocket.RemotePort)

Dim ps As TSubscriber

    If mSubscribers.IndexOf(szDestination) = 0 Then
        ' /* not already subscribed */
        
        
'        ' /* check password */
'        If ps.Password <> Data.ValueOf("password") Then
'            g_Debug "TSubscriberRoster.uAddNewSubscriber(): password mismatch", LEMON_LEVEL_CRITICAL
'            gSetLastError SNARL_ERROR_AUTH_FAILURE
'
'        ElseIf Data.Exists("app-sig") Then
'            ' /* add specific apps */
'            uAddNewSubscriber = ps.AddApps(Data)
'
'        Else
'            ' /* cannot do this if already subscribed */
'            g_Debug "TSubscriberRoster.uAddNewSubscriber(): cannot do wildcard subscribe", LEMON_LEVEL_CRITICAL
'            gSetLastError SNARL_ERROR_FAILED
'
'        End If
        
        
        
        Set ps = New TSubscriber
        If ps.SetTo2(Protocol, RemoteSocket, Data) Then
            mSubscribers.Add ps
            gSetLastError SNARL_SUCCESS
            AddSubscriber2 = True

        Else
            g_Debug "failed to add subscriber " & g_Quote(szDestination), LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_FAILED

        End If
    Else
        ' /* already subscribed */
        g_Debug g_Quote(szDestination) & " is already subscribed", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        gSetLastError SNARL_ERROR_ALREADY_SUBSCRIBED

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function RemoveSubscriber(ByRef Socket As CSocket, Optional ByRef Data As BPackedData) As Boolean

    g_Debug "TSubscriberRoster.RemoveSubscriber()", LEMON_LEVEL_PROC_ENTER

    If (Socket Is Nothing) Then
        g_Debug "missing RemoteSocket or Data", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        gSetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

Dim szDest As String
Dim i As Long

    szDest = Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort)

    i = mSubscribers.IndexOf(szDest)
    If i Then
        g_Debug g_Quote(szDest) & " removed", LEMON_LEVEL_INFO
        mSubscribers.Remove i
        frmAbout.bSubsChanged
        RemoveSubscriber = True

    Else
        g_Debug g_Quote(szDest) & " is not subscribed", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_NOT_SUBSCRIBED

    End If

'        ElseIf ps.Password <> Data.ValueOf("password") Then
'            ' /* wrong password */
'            g_Debug "TSubscriberRoster.Remove(): password mismatch", LEMON_LEVEL_CRITICAL
'            gSetLastError SNARL_ERROR_AUTH_FAILURE

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function IsSubscriber(ByVal AddressAndPort As String) As Boolean

    IsSubscriber = (mSubscribers.IndexOf(AddressAndPort) <> 0)

End Function
