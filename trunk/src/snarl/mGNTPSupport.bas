Attribute VB_Name = "mGNTPSupport"
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           mGNTPSupport.bas
    '/
    '/  Description:    GNTP support routines
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Enum E_GNTP_CODES
    INVALID_REQUEST = 300           '// The request contained an unsupported directive, invalid headers or values, or was otherwise malformed
    UNKNOWN_PROTOCOL = 301          '// The request was not a GNTP request
    UNKNOWN_PROTOCOL_VERSION = 302  '// The request specified an unknown or unsupported GNTP version
    REQUIRED_HEADER_MISSING = 303   '// The request was missing required information
    NOT_AUTHORIZED = 400            '// The request supplied a missing or wrong password/key or was otherwise not authorized
    UNKNOWN_APPLICATION = 401       '// Application is not registered to send notifications
    UNKNOWN_NOTIFICATION = 402      '// Notification type is not registered by the application
    INTERNAL_SERVER_ERROR = 500     '// An internal server error occurred while processing the request

End Enum


Private Type T_NOTIFY_TYPE
    Name As String
    DisplayName As String
    Enabled As Boolean
    Icon As String

End Type

Private Type T_REG
    Token As Long
    AppName As String
    NotificationType() As T_NOTIFY_TYPE
    Count As Long
    Signature As String
    IconPath As String

End Type

Dim mRegistration As T_REG

Private Type T_GNTP_NOTIFICATION
    AppName As String
    Name As String
    Title As String
    Id As String
    Text As String
    Sticky As Boolean
    Priority As Long
    Icon As String
    CoalesceID As String
    CallbackContext As String
    CallbackContextType As String
    CallbackContextTarget As String

End Type

Dim mSection() As String
Dim mDirective As String
Dim mCustomHeaders As String
Dim mRedactResponse As Boolean

' /*********************************************************************************************
'   gntp_Process() -- Master GNTP request handler
'
'   Inputs
'       Request - unabridged request content
'       Sender - sending socket object (only used for notifications)
'
'   Outputs
'       Response - GNTP response that should be sent back to the source socket
'       KeepSocketOpen - Set to TRUE if the socket should remain open
'
'   Return Value
'       None
'
' *********************************************************************************************/

Public Sub gntp_Process(ByVal Request As String, ByRef Sender As CSocket, ByRef Response As String, ByRef KeepSocketOpen As Boolean)

    ' /* return a GNTP error code here */

    On Error GoTo er

    mCustomHeaders = ""
    KeepSocketOpen = False
    mRedactResponse = False

    uOutput "gntp_Process()"
    uIndent

    ' /* split into sections */

    mSection = Split(Request, vbCrLf & vbCrLf)
    uOutput "section count = " & UBound(mSection) - 1

    If UBound(mSection) < 1 Then
        uOutput "failed: no sections"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Sub

    End If

    ' /* parse section 1 - must be the info block */

    mDirective = ""

    If Not uParse(0, Response) Then
        uOutput "failed: invalid info block"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Sub

    End If

    uOutput "info block okay, directive is '" & mDirective & "'"

    Select Case mDirective
    Case "REGISTER"
        Debug.Print uDoRegistration(Response)

    Case "NOTIFY"
        Debug.Print uDoNotification(Response, Sender, KeepSocketOpen)

    Case "SUBSCRIBE"
        Debug.Print uDoSubscription(Response, Sender)

    Case Else
        uOutput "unsupported directive '" & mDirective & "'"
        Response = uCreateResponse(INVALID_REQUEST)

    End Select

    uOutput "done"
    uOutdent
    Exit Sub

er:
    Debug.Print "  panic: " & err.Description
    uOutdent

End Sub

Private Function uBool(ByVal str As String) As Boolean

    Select Case LCase$(str)
    Case "yes", "true"
        uBool = True

    End Select

End Function

Private Function uDoRegistration(ByRef Response As String) As Boolean

    mCustomHeaders = "Response-Action: REGISTER"

Dim pp As TPackedData

    Set pp = New TPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        uOutput "uDoRegistration(): bad data"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

#If GNTP_TEST = 1 Then
    uListPackedData pp
#End If

    ' /* 44.51: supports X-Response-Redacted: <bool> - credit to Viscount for suggesting this */

    If pp.Exists("X-Response-Redacted") Then _
        mRedactResponse = uBool(pp.ValueOf("X-Response-Redacted"))

    uOutput "uDoRegistration(): redact_response = " & CStr(mRedactResponse)

    ' /* required items */

Dim szAppName As String
Dim szAppSig As String

    If Not pp.Exists("Application-Name") Then
        uOutput "uDoRegistration(): missing app name"
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        Exit Function

    Else
        szAppName = trim(pp.ValueOf("Application-Name"))
        szAppSig = "application/x-gntp-" & Replace$(szAppName, " ", "_")

    End If

Dim dwCount As Long

    If Not pp.Exists("Notifications-Count") Then
        uOutput "uDoRegistration(): missing count"
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        Exit Function

    Else
        dwCount = g_SafeLong(pp.ValueOf("Notifications-Count"))

    End If

    ' /* special Snarl feature: zero notifications means unregister */

    If dwCount = 0 Then
        uOutput "uDoRegistration(): app requested an unregister"
        Response = uCreateResponse(0)

#If GNTP_TEST Then
        uOutput "uDoRegistration(): snarl_unregister() returned " & snarl_unregister(szAppSig)

#Else
        g_DoAction "unreg", 0, g_newBPackedData("app-sig::" & szAppSig)

#End If
        Exit Function

    End If


Dim px As T_REG

    LSet mRegistration = px

    With mRegistration
        .AppName = szAppName
        .Signature = szAppSig
        .Count = dwCount
        ReDim .NotificationType(.Count)

        ' /* 44.55: translate GNTP icon resources into path to saved image */

        If pp.Exists("Application-Icon") Then _
            .IconPath = uTranslateIconPath(trim(pp.ValueOf("Application-Icon")))

    End With

    ' /* otherwise must have the right number of sections */

    If UBound(mSection) < mRegistration.Count Then
        uOutput "uDoRegistration(): not enough sections"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    ' /* sections 1 to pr.Count should be notification types */

Dim i As Long
Dim c As Long

    For i = 1 To mRegistration.Count
        ' /* if adding any of the notification types fails, we fail */
        If Not uAddNotificationType(mSection(i), i - 1, Response) Then _
            Exit Function

    Next i

    ' /* sections pr.Count to end should be resource identifiers */

    c = UBound(mSection) - 1

    If c >= mRegistration.Count Then
        uOutput "uDoRegistration(): parsing resource identifiers " & CStr(mRegistration.Count) & " to " & CStr(c)
        For i = mRegistration.Count To c
            uParse i, ""

        Next i

    End If

    uOutput "uDoRegistration(): registering with Snarl..."

    ' /* register here */

    With mRegistration

#If GNTP_TEST Then
        .Token = snarl_register(.Signature, .AppName, .IconPath)
        uOutput "--"
        uOutput "uDoRegistration(): Snarl replied with " & CStr(.Token) & " to app-sig='" & .Signature & "' app-name='" & .AppName & "' icon=" & .IconPath & "'"
        uOutput "--"
        .Token = 999        ' // don't fail

#Else
        .Token = g_DoAction("register", 0, _
                                         g_newBPackedData("app-sig::" & .Signature & _
                                                          "#?title::" & .AppName & _
                                                          "#?icon::" & .IconPath))
#End If

    End With

    ' /* if registration fails, quit now */

    If mRegistration.Token < 1 Then
        uOutput "uDoRegistration(): registration failed (" & CStr(Abs(mRegistration.Token)) & ")"
        Response = uCreateResponse(INTERNAL_SERVER_ERROR)
        Exit Function

    End If

    ' /* add notification types as classes */

Dim szReq As String
Dim hr As Long

    With mRegistration
        For i = 0 To .Count - 1

#If GNTP_TEST Then
            szReq = "addclass?app-sig=" & .Signature & _
                    "&id=" & .NotificationType(i).Name & _
                    "&name=" & .NotificationType(i).DisplayName & _
                    "&enabled=" & IIf(.NotificationType(i).Enabled, "1", "0") & _
                    "&icon=" & .NotificationType(i).Icon

            hr = snDoRequest(szReq)
            uOutput "uDoRegistration(): Snarl replied with " & CStr(hr) & " to '" & szReq & "'"

#Else

            g_DoAction "addclass", 0, _
                       g_newBPackedData("app-sig::" & mRegistration.Signature & _
                                        "#?id::" & .NotificationType(i).Name & _
                                        "#?name::" & .NotificationType(i).DisplayName & _
                                        "#?enabled::" & IIf(.NotificationType(i).Enabled, "1", "0") & _
                                        "#?icon::" & .NotificationType(i).Icon)
#End If

        Next i

    End With

    ' /* done */

    Response = uCreateResponse(0)
    uDoRegistration = True

End Function

Private Function uAddNotificationType(ByVal str As String, ByVal Index As Long, ByRef Response As String) As Boolean
Dim pp As BPackedData

    Set pp = New BPackedData
    If Not pp.SetTo(str, vbCrLf, ": ") Then
        uOutput "uAddNotificationType(): invalid data for notification type #" & CStr(Index)
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    'For each notification being registered:
    '
    'Notification-Name: <string>
    'Required - The name (type) of the notification being registered

    If Not pp.Exists("Notification-Name") Then
        uOutput "uAddNotificationType(): missing required arg Notification-Name for notification type #" & CStr(Index)
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        Exit Function

    End If

    'Notification-Display-Name: <string>
    'Optional - The name of the notification that is displayed to the user (defaults to the same value as Notification-Name)
    '
    'Notification-Enabled: <boolean>
    'Optional - Indicates if the notification should be enabled by default (defaults to False)
    '
    'Notification-Icon: <url> | <uniqueid>
    'Optional - The default icon to use for notifications of this type
    '
    'Each notification being registered should be seperated by a blank line, including the first notification.

Dim sx() As String

    With mRegistration.NotificationType(Index)
        .Name = trim(pp.ValueOf("Notification-Name"))
        .DisplayName = trim(pp.ValueOf("Notification-Display-Name"))
        .Enabled = uBool(pp.ValueOf("Notification-Enabled"))
        .Icon = trim(pp.ValueOf("Notification-Icon"))

        ' /* if the icon is actually an identifier, map it to the locally saved copy */

        If g_SafeLeftStr(.Icon, 19) = "x-growl-resource://" Then
            sx = Split(.Icon, "://")
            .Icon = g_GetTempPath() & "gntp-res-" & sx(1) & ".png"

        End If

        ' /* as per GNTP specification */

        If .DisplayName = "" Then _
            .DisplayName = .Name

        uOutput "uAddNotificationType(): got notification type " & CStr(Index) & " (" & .Name & "¶" & .DisplayName & "¶" & IIf(.Enabled, "Enabled", "Disabled") & "¶" & .Icon & ")"

    End With

    uAddNotificationType = True

End Function

Private Function uParse(ByVal SectionIndex As Long, ByRef Response As String) As Boolean
Dim s() As String

    uOutput "uParse()"
    uIndent

    ' /* parses the first line of the given section and returns
    '    the appropriate result */

    s = Split(mSection(SectionIndex), vbCrLf)
    If UBound(s) < 1 Then
        uOutput "failed: invalid section"
        uOutdent
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    uOutput "section #" & CStr(SectionIndex) & " header is '" & s(0) & "'"

    ' /* identify section type from the first line */

Dim x() As String

    If g_SafeLeftStr(s(0), 4) = "GNTP" Then
        ' /* information line first */
        uOutput "is information block"
        uParse = uParseInfoLine(s(0), Response)

    ElseIf g_SafeLeftStr(s(0), 12) = "Identifier: " Then
        ' /* resource identifier */
        uOutput "is resource identifier"
        x = Split(s(0), ": ")
        uSaveBinary SectionIndex + 1, x(1)
        uParse = True

    Else
        ' /* other headers... */
        uOutput "is unknown!"

    End If

    uOutdent

End Function

Private Function uParseInfoLine(ByVal str As String, ByRef Response As String) As Boolean
Dim s() As String

    s = Split(str, " ")
    If UBound(s) < 2 Then
        uOutput "uParseInfoLine(): not enough params"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function           ' // not enough params

    End If

Dim v() As String

    v = Split(s(0), "/")
    If UBound(v) <> 1 Then
        uOutput "uParseInfoLine(): not GNTP"
        Response = uCreateResponse(UNKNOWN_PROTOCOL)
        Exit Function           ' // not GNTP
    
    End If

    If v(0) <> "GNTP" Then
        uOutput "uParseInfoLine(): not GNTP"
        Response = uCreateResponse(UNKNOWN_PROTOCOL)
        Exit Function           ' // not GNTP
    
    End If

    If v(1) <> "1.0" Then
        uOutput "uParseInfoLine(): not 1.0"
        Response = uCreateResponse(UNKNOWN_PROTOCOL_VERSION)
        Exit Function

    End If

    mDirective = ""

    Select Case s(1)
    Case "REGISTER", "NOTIFY", "SUBSCRIBE"
        mDirective = s(1)

    Case Else
        uOutput "uParseInfoLine(): unsupported directive"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function           ' // bad directive

    End Select

    Select Case s(2)
    Case "NONE"
    
    Case Else
        uOutput "uParseInfoLine(): unsupported encryption (" & s(2) & ")"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function           ' // unsupported encryption

    End Select

    uParseInfoLine = True

End Function

Private Sub uSaveBinary(ByVal SectionIndex As Long, ByVal Identifier As String)
Dim i As Integer

    On Error Resume Next

    Identifier = "gntp-res-" & Identifier
    uOutput "uSaveBinary(): writing binary content to " & g_GetTempPath() & Identifier '& ".png"

    i = FreeFile()
    Open g_GetTempPath() & Identifier For Binary Access Write As #i
    Put #i, , mSection(SectionIndex)
    Close #i

End Sub

Private Function uDoNotification(ByRef Response As String, ByRef Sender As CSocket, ByRef KeepSocketOpen As Boolean) As Boolean

    uOutput "uDoNotification()"
    uIndent

    mCustomHeaders = "Response-Action: NOTIFY"

Dim pp As TPackedData

    ' /* convert the mime-style section into packed data */

    Set pp = New TPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        uOutput "bad data"
        Response = uCreateResponse(INVALID_REQUEST)
        uOutdent
        Exit Function

    End If

    ' /* 44.51: supports X-Response-Redacted: <bool> - credit to Viscount for suggesting this */

    If pp.Exists("X-Response-Redacted") Then _
        mRedactResponse = uBool(pp.ValueOf("X-Response-Redacted"))

    uOutput "redact_response = " & CStr(mRedactResponse)

    ' /* required items */

    If (Not pp.Exists("Application-Name")) Or (Not pp.Exists("Notification-Name")) Or (Not pp.Exists("Notification-Title")) Then
        uOutput "missing required arg"
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        uOutdent
        Exit Function

    End If

Dim c As Long
Dim i As Long

    ' /* sections pr.Count to end should be resource identifiers */

    c = UBound(mSection) - 1
    If c > 0 Then
        uOutput "scanning for resource identifiers in blocks 1 to " & CStr(c) & "..."
        For i = 1 To c
            uParse i, ""

        Next i

    Else
        uOutput "no blocks to check"

    End If

    uOutput "building notification content..."

#If GNTP_TEST = 1 Then
    uListPackedData pp
#End If

    ' /* build the Snarl packet */

Dim px As TPackedData

    Set px = New TPackedData
    With px
        'Application-Name: <string>
        'Required - The name of the application that sending the notification (must match a previously registered application)
        .Add "app-sig", "application/x-gntp-" & Replace$(trim(pp.ValueOf("Application-Name")), " ", "_")

        'Notification-Name: <string>
        'Required - The name (type) of the notification (must match a previously registered notification name registered by the
        'application specified in Application-Name)
        .Add "id", pp.ValueOf("Notification-Name")

        'Notification-Title: <string>
        'Required - The notification's title
        .Add "title", g_toUnicodeUTF8(Replace$(trim(pp.ValueOf("Notification-Title")), Chr$(13), vbCrLf))
        
        'Notification-Text: <string>
        'Optional - The notification's text. (defaults to "")
        .Add "text", g_toUnicodeUTF8(Replace$(trim(pp.ValueOf("Notification-Text")), Chr$(13), vbCrLf))

        'Notification-Sticky: <boolean>
        'Optional - Indicates if the notification should remain displayed until dismissed by the user. (default to False)
        ' /* sticky == zero duration under Snarl */
        If uBool(pp.ValueOf("Notification-Sticky")) Then
            .Add "timeout", "0"

        ElseIf pp.Exists("X-Notification-Duration") Then
            ' /* 44.51: supports X-Notification-Duration: <seconds> - credit to Viscount for suggesting this */
            .Add "timeout", pp.ValueOf("X-Notification-Duration")

        End If

        'Notification-Priority: <int>
        'Optional - A higher number indicates a higher priority. This is a display hint for the receiver which may be ignored. (valid
        'values are between -2 and 2, defaults to 0)
        If pp.Exists("Notification-Priority") Then _
            .Add "priority", pp.ValueOf("Notification-Priority")

        'Notification-Coalescing-ID: <string>
        'Optional - If present, should contain the value of the Notification-ID header of a previously-sent notification. This serves
        'as a hint to the notification system that this notification should replace/update the matching previous notification. The
        'notification system may ignore this hint.
        If pp.Exists("Notification-Coalescing-ID") Then _
            .Add "update-uid", pp.ValueOf("Notification-Coalescing-ID")

        'Notification-ID: <string>
        'Optional - A unique ID for the notification. If used, this should be unique for every request, even if the notification is
        'replacing a current notification (see Notification-Coalescing-ID)
        If pp.Exists("Notification-ID") Then _
            .Add "uid", trim(pp.ValueOf("Notification-ID"))

        'Notification-Callback-Target: <string>
        'Optional - An alternate target for callbacks from this notification. If passed, the standard behavior of performing the
        'callback over the original socket will be ignored and the callback data will be passed to this target instead. See the 'Url
        'Callbacks' section for more information.
        If pp.ValueOf("Notification-Callback-Target") <> "" Then
            ' /* static callback */
            .Add "callback", trim(pp.ValueOf("Notification-Callback-Target"))

        ElseIf (pp.ValueOf("Notification-Callback-Context") <> "") And (pp.ValueOf("Notification-Callback-Context-Type") <> "") Then
            ' /* we have a dynamic callback */

            'Notification-Callback-Context: <string>
            'Optional - Any data (will be passed back in the callback unmodified)

            'Notification-Callback-Context-Type: <string>
            'Optional, but Required if 'Notification-Callback-Context' is passed - The type of data being passed in
            'Notification-Callback-Context (will be passed back in the callback unmodified). This does not need to be of any pre-defined
            'type, it is only a convenience to the sending application.

            .Add "callback-context", trim(pp.ValueOf("Notification-Callback-Context"))
            .Add "callback-type", trim(pp.ValueOf("Notification-Callback-Context-Type"))

            uOutput "uDoNotification(): dynamic callback specified: context=" & pp.ValueOf("Notification-Callback-Context") & " type=" & pp.ValueOf("Notification-Callback-Context-Type")
            KeepSocketOpen = True

        End If

        ' /* sort out the icon */

Dim sx() As String
Dim sz As String

        'Notification-Icon: <url> | <uniqueid>
        'Optional - The icon to display with the notification.
        sz = trim(pp.ValueOf("Notification-Icon"))
        If g_SafeLeftStr(sz, 19) = "x-growl-resource://" Then
            sx = Split(sz, "://")
            .Add "icon", g_GetTempPath() & "gntp-res-" & sx(1) & ".png"

        ElseIf sz <> "" Then
            .Add "icon", sz

        End If

    End With

    ' /* add anything starting with Data- */

Dim szd As String

    With pp
        .Rewind
        Do While .GetNextItem(sz, szd)
            If g_SafeLeftStr(sz, 5) = "Data-" Then _
                px.Add sz, trim(szd)

        Loop

    End With

    ' /* do the notification */

Dim hr As Long

#If GNTP_TEST = 1 Then

    sz = px.AsString()
    sz = Replace$(sz, "::", "=")
    sz = Replace$(sz, "#?", "&")
    hr = snDoRequest("notify?" & sz)

    uOutput "Snarl replied with " & CStr(hr) & " to 'notify?" & sz & "'"
    If hr < 0 Then _
        uOutput "(error is ignored by GNTP Listener)"

    Response = uCreateResponse(0)
    uDoNotification = True

#Else

Dim lFlags As SN_NOTIFICATION_FLAGS
Dim pxSnarl As BPackedData

    If Sender.LocalIP <> "127.0.0.1" Then _
        lFlags = lFlags Or SN_NF_REMOTE

    ' /* new for R2.4.2 */
        lFlags = lFlags Or SN_NF_IS_GNTP

'    If (px.Exists("callback-context")) And (px.Exists("callback-type")) Then _
        lFlags = lFlags Or SN_NF_GNTP_CALLBACK

    Set pxSnarl = New BPackedData
    pxSnarl.SetTo px.AsString

    hr = g_DoAction("notify", 0, pxSnarl, lFlags Or App.Major, Sender)
    If hr = 0 Then
        ' /* failed */

        uOutput "<notify> failed: " & CStr(g_QuickLastError())
        Select Case g_QuickLastError()
        Case SNARL_ERROR_AUTH_FAILURE
            Response = uCreateResponse(NOT_AUTHORIZED)

        Case SNARL_ERROR_NOT_REGISTERED
            Response = uCreateResponse(UNKNOWN_APPLICATION)

        Case Else
            Response = uCreateResponse(INTERNAL_SERVER_ERROR)

        End Select

    Else
        Response = uCreateResponse(0)
        uDoNotification = True

    End If

#End If

    uOutdent

End Function

Private Function uError(ByVal ResponseCode As E_GNTP_CODES) As String

    Select Case ResponseCode
    Case INVALID_REQUEST
        'The request contained an unsupported directive, invalid headers or values, or was otherwise malformed
        uError = "Invalid request"

    Case UNKNOWN_PROTOCOL
        'The request was not a GNTP request
        uError = "Unknown protocol"

    Case UNKNOWN_PROTOCOL_VERSION
        'The request specified an unknown or unsupported GNTP version
        uError = "Unknown protocol version"

    Case REQUIRED_HEADER_MISSING
        'The request was missing required information
        uError = "Required header missing"

    Case NOT_AUTHORIZED
        'The request supplied a missing or wrong password/key or was otherwise not authorized
        uError = "Not authorized"

    Case UNKNOWN_APPLICATION
        'Application is not registered to send notifications
        uError = "Unknown application"

    Case UNKNOWN_NOTIFICATION
        'Notification type is not registered by the application
        uError = "Unknown notification"

    Case INTERNAL_SERVER_ERROR
        'An internal server error occurred while processing the request
        uError = "Internal server error"

    End Select

End Function

Private Function uCreateResponse(ByVal ResponseCode As E_GNTP_CODES) As String
Dim sz As String

    sz = "GNTP/1.0 "

    Select Case ResponseCode
    Case 0
        sz = sz & "-OK"

    Case Else
        sz = sz & "-ERROR"

    End Select

    sz = sz & " NONE" & vbCrLf

    ' /* error details */

    If ResponseCode <> 0 Then
         sz = sz & "Error-Code: " & CStr(ResponseCode) & vbCrLf & _
                   "Error-Description: " & uError(ResponseCode) & vbCrLf

    End If

    ' /* custom headers */

    If mCustomHeaders <> "" Then _
        sz = sz & mCustomHeaders & vbCrLf

    ' /* generic headers */
    
    uAddStandardHeaders sz

    uCreateResponse = sz & vbCrLf & vbCrLf

End Function

Private Sub uOutput(ByVal Text As String)

#If GNTP_TEST = 1 Then
    Form1.output Text

#Else
    g_Debug Text

#End If

End Sub

Public Function gntp_CreateCallbackResponse(ByVal Name As String, ByVal Application As String, ByRef OriginalContent As String, ByRef Response As String) As Boolean
Dim pp As BPackedData

    Set pp = New BPackedData
    If Not pp.SetTo(OriginalContent) Then
        uOutput "gntp_CreateCallbackResponse(): OriginalContent missing or invalid"
        Exit Function

    End If

    If (Not pp.Exists("callback-context")) Or (Not pp.Exists("callback-type")) Then
        uOutput "gntp_CreateCallbackResponse(): callback-context and/or callback-type missing"
        Exit Function

    End If

    Response = "GNTP/1.0 -CALLBACK NONE" & vbCrLf

    'Application-Name: <string>
    'Required - The name of the application that sent the original request
    Response = Response & "Application-Name: " & Application & vbCrLf

    'Notification-ID: <string>
    'Required - The value of the 'Notification-ID' header from the original request
    Response = Response & "Notification-ID: " & pp.ValueOf("uid") & vbCrLf

    'Notification-Callback-Result: <string>
    'Required - [CLICKED|CLOSED|TIMEDOUT] | [CLICK|CLOSE|TIMEOUT]
    Response = Response & "Notification-Callback-Result: " & Name & vbCrLf

    'Notification-Callback-Timestamp: <date>
    'Required - The date and time the callback occurred
    Response = Response & "Notification-Callback-Timestamp: " & Format$(Now(), "MM/DD/YYYY HH:MM:SS AMPM") & vbCrLf

    'Notification-Callback-Context: <string>
    'Required - The value of the 'Notification-Callback-Context' header from the original request
    Response = Response & "Notification-Callback-Context: " & pp.ValueOf("callback-context") & vbCrLf

    'Notification-Callback-Context-Type: <string>
    'Required - The value of the 'Notification-Callback-Context-Type' header from the original request
    Response = Response & "Notification-Callback-Context-Type: " & pp.ValueOf("callback-type") & vbCrLf

    ' /* add anything with 'Data-' */

Dim szn As String
Dim szv As String

    With pp
        Do While .GetNextItem(szn, szv)
            If g_SafeLeftStr(szn, 5) = "Data-" Then _
                Response = Response & szn & ": " & szv & vbCrLf

        Loop

    End With

    'Callbacks may also contain the generic headers defined in the 'Responses' section, as well as custom headers as defined in the 'Custom Headers' portion of the 'Requests' section above.

    uAddStandardHeaders Response

    uOutput "++ gntp_CreateCallbackResponse() +++++++++++++++++++++++++++"
    uOutput Response
    uOutput "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    Response = Response & vbCrLf & vbCrLf

    gntp_CreateCallbackResponse = True

End Function

Private Sub uAddStandardHeaders(ByRef Response As String)

    If mRedactResponse Then
        g_Debug "uAddStandardHeaders(): redacting...", LEMON_LEVEL_INFO
        Exit Sub

    End If

    Response = Response & "Origin-Machine-Name: " & get_host_name() & vbCrLf

#If GNTP_TEST = 1 Then
    Response = Response & "Origin-Software-Name: " & App.ProductName & vbCrLf
    Response = Response & "Origin-Software-Version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf

#Else
    Response = Response & "Origin-Software-Name: Snarl" & vbCrLf
    Response = Response & "Origin-Software-Version: " & CStr(APP_VER) & "." & CStr(APP_SUB_VER) & " (" & CStr(App.Major) & "." & CStr(App.Revision) & ")" & vbCrLf

#End If
    Response = Response & "Origin-Platform-Name: Windows " & g_GetOSVersionString(True) & vbCrLf
    Response = Response & "Origin-Platform-Version: " & g_GetOSVersionString() & vbCrLf

    Response = Response & "X-Message-Daemon: Snarl" & vbCrLf
    Response = Response & "X-Timestamp: " & Format$(Now(), "MM/DD/YYYY HH:MM:SS AMPM")

End Sub

Private Sub uListPackedData(ByRef pp As TPackedData)

    If (pp Is Nothing) Then _
        Exit Sub

Dim zz As String
Dim vv As String

    uOutput "++"

    With pp
        .Rewind
        Do While .GetNextItem(zz, vv)
            uOutput g_Quote(zz) & " == " & g_Quote(vv)

        Loop

    End With

    uOutput "++"

End Sub

Private Sub uIndent()

#If GNTP_TEST = 1 Then
    Form1.indent
#End If

End Sub

Private Sub uOutdent()

#If GNTP_TEST = 1 Then
    Form1.outdent
#End If

End Sub

Private Function uDoSubscription(ByRef Response As String, ByRef Sender As CSocket) As Boolean

    mCustomHeaders = "Response-Action: SUBSCRIBE"

'    ' /* for now... */
'
'    Response = uCreateResponse(INTERNAL_SERVER_ERROR)
'    Exit Function


Dim pp As TPackedData

    Set pp = New TPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        uOutput "uDoSubscription(): bad data"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

#If GNTP_TEST = 1 Then
    uListPackedData pp
#End If

'Dim szSubID As String
'Dim szSubName As String

    ' /* required items */

'Subscriber-ID: <string>
'Required - A unique id (UUID) that identifies the subscriber
'
'Subscriber-Name: <string>
'Required - The friendly name of the subscribing machine

    If (Not pp.Exists("Subscriber-ID")) Or (Not pp.Exists("Subscriber-Name")) Then
        uOutput "uDoRegistration(): missing subscriber ID or name"
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        Exit Function

    End If

Dim pArgs As BPackedData

    Set pArgs = New BPackedData
    pArgs.Add "id", pp.ValueOf("Subscriber-ID")
    pArgs.Add "name", pp.ValueOf("Subscriber-Name")

'Subscriber-Port: <int>
'Optional - The port that the subscriber will listen for notifications on (defaults to the standard 23053)

    If pp.Exists("Subscriber-Port") Then
        pArgs.Add "reply-port", pp.ValueOf("Subscriber-Port")

    Else
        pArgs.Add "reply-port", CStr(GNTP_DEFAULT_PORT)

    End If

Dim hr As SNARL_STATUS_CODE

    hr = g_SubsRoster.AddSubscriber("gntp", Sender, pArgs)
    If hr = SNARL_SUCCESS Then
        ' /* done */
        Response = uCreateResponse(0)
        uDoSubscription = True

    Else
        ' /* TO DO: create appropriate GNTP error based on result */
        Select Case hr
        Case SNARL_ERROR_ALREADY_SUBSCRIBED
            Response = uCreateResponse(0)
            uDoSubscription = True

        Case Else
            Response = uCreateResponse(INTERNAL_SERVER_ERROR)

        End Select

    End If

End Function

'SUBSCRIBE
'Subscriber-ID: <string>
'Required - A unique id (UUID) that identifies the subscriber
'
'Subscriber-Name: <string>
'Required - The friendly name of the subscribing machine
'
'Subscriber-Port: <int>
'Optional - The port that the subscriber will listen for notifications on (defaults to the standard 23053)
'
'Subscription requests MUST include the key hash (In practice, only instances subscribing from another machine make sense, so they would have to provide the key hash anyway since they are communicating over the network.)
'
'Once the subscription is made, the subscribed-to machine should forward requests to the subscriber, just as if the end user had configured the subscriber using the standard forwarding mechanisms, with one important distinction: the forwarding machine should use a specially-constructed password consisting of -its own- password plus the Subscriber-ID value when constructing the forwarded messages (as opposed to the usual behavior of using the receiver's password to construct the message - see example below). This is due to the fact that the subscribed-to machine does not know the subscriber's password (but the subscriber must know/provide the subscribed-to machine's password when subscribing). Note that the subscriber should be prepared to accept incoming requests using this password for the lifetime of the subscription. This also ensures that if the password is changed on the subscribed-to machine, the subscribed machine can no longer receive subscribed notifications.
'
'Example:
'Subscribed-to Machine's password: foo
'Subscriber-ID value sent by subscribing machine: 0f8e3530-7a29-11df-93f2-0800200c9a66
'Resulting password used by subscribed-to machine when forwarding: foo0f8e3530-7a29-11df-93f2-0800200c9a66If the subscriber wants to remain subscribed, they must issue another SUBSCRIBE request before the TTL period has elapsed. It is recommended that the subscribed-to machine not set the TTL to less than 60 seconds to give the subscriber time to process notifications and issue their renewal requests.
'
'If the subscriber does not issue another SUBSCRIBE request before the TTL period has elapsed, the subscribed-to machine should stop forwarding requests to the subscriber.

Public Function gntp_CreateSubscribeRequest(ByVal UUID As String, ByVal Name As String, ByVal Hash As String, Optional ByVal Port As Long = 23053) As String

    gntp_CreateSubscribeRequest = "GNTP/1.0 SUBSCRIBE NONE " & Hash & vbCrLf & _
                                  "Subscriber-ID: " & UUID & vbCrLf & _
                                  "Subscriber-Name: " & Name & vbCrLf & _
                                  "Subscriber-Port: " & CStr(Port) & vbCrLf & vbCrLf

End Function

Public Function gntp_IsResponse(ByVal RawData As String, Optional ByRef ResponseType As String) As Boolean
Dim sz As String
Dim s() As String

    sz = uGetFirstLine(RawData)
    If uIsResponseHeader(sz) Then
        s() = Split(sz, " ")
        ResponseType = s(1)
        gntp_IsResponse = True

    End If

End Function

Private Function uGetFirstLine(ByVal str As String, Optional ByVal EndMarker As String = vbCrLf) As String
Dim i As Long

    i = InStr(str, EndMarker)
    If i Then _
        uGetFirstLine = g_SafeLeftStr(str, i - 1)

End Function

Private Function uIsResponseHeader(ByVal Header As String) As Boolean

    uIsResponseHeader = (InStr(Header, " -OK") > 0) Or (InStr(Header, " -ERROR") > 0)

End Function

Private Function uTranslateIconPath(ByVal Path As String) As String
Dim sz() As String

    If g_SafeLeftStr(Path, 19) = "x-growl-resource://" Then
        sz = Split(Path, "://")
        uTranslateIconPath = g_GetTempPath() & "gntp-res-" & sz(1) '& ".png"

    Else
        uTranslateIconPath = Path

    End If

End Function
