Attribute VB_Name = "mSNP3Support"
Option Explicit

Public Function snp3_GotPacketEnd(ByVal Packet As String) As Boolean

    snp3_GotPacketEnd = (InStr(Packet, vbCrLf & "END" & vbCrLf) > 0)

End Function

Public Function snp3_IsSNP3(ByVal Packet As String) As Boolean
Dim sz As String

    sz = uGetFirstLine(Packet)
    snp3_IsSNP3 = (g_SafeLeftStr(sz, 7) = "SNP/3.0")

End Function

'Public Function snp3_CreateSubscribeRequest() As String
'
'    snp3_CreateSubscribeRequest = "SNP/3.0" & vbCrLf & "subscribe" & vbCrLf & "END" & vbCrLf
'
'End Function

'Private Sub uSNP3Parse(ByRef Data As BPackedData)
'Dim sx As String
'Dim sr As String
'
'    ' /* get last entry */
'
'    Data.EntryAt Data.Count, sx, ""
'
'    If sx = "END" Then
'        Debug.Print "SNP3 packet received!!!"
'        theSocket.GetData sx                    ' // empty the socket buffer
'        mDoingSNP3 = False
'
'        ' /* if translation fails, we must close the socket immediately after responding */
'
'        If Not uSNP3Translate(Data, sr) Then
'            theSocket.SendData sr
'            theSocket.CloseSocket
'
'        Else
'            theSocket.SendData sr
'
'        End If
'
'    End If
'
'End Sub


'Private Function uSNP3Translate(ByRef Request As BPackedData, ByRef Response As String) As Boolean
'Dim bErr As Boolean
'Dim szr As String
'
'    ' /* return False if the socket should be closed */
'
'    Debug.Print "----SNP3---"
'    Debug.Print Request.AsString
'    Debug.Print "----SNP3---"
'
'Dim sz As String
'
'    ' /* get the header */
'
'    Request.EntryAt 1, sz, ""
'    If g_SafeLeftStr(sz, 7) <> "SNP/3.0" Then
'        Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_VERSION, "Invalid version in header")
'        Exit Function
'
'    End If
'
'Dim ppdHeader As BPackedData
'
'    ' /* split header into component parts */
'
'    Set ppdHeader = New BPackedData
'    ppdHeader.SetTo sz, " ", ":"
'
'Dim i As Long
'Dim n As Long
'
'    ' /* check for empty message */
'
'    With Request
'        If .Count > 2 Then
'            For i = 2 To .Count - 1
'                .EntryAt i, sz, ""
'                If (g_SafeLeftStr(sz, 1) <> "#") Then _
'                    n = n + 1
'
'            Next i
'        End If
'
'    End With
'
'    If n = 0 Then
'        Response = snp3_BuildResponse(SNARL_ERROR_NO_ACTIONS_PROVIDED, "Must supply at least one action")
'        Exit Function
'
'    End If
'
'
'Dim szExpectedKey As String
'
'    ' /* get current password */
'
'    szExpectedKey = g_GetPassword()
'
'    ' /* decode header - currently SNP/3.0 [HASH:value.salt] [ENCRYPTION:value] */
'
'Dim szProvidedKey As String
'Dim szProvidedSalt As String
'Dim szHashType As String
'
'    If ppdHeader.Count > 1 Then
'        ' /* get provided hashing algorithm, key and salt */
'        ppdHeader.EntryAt 2, szHashType, szProvidedKey
'
'        If Not uGetSalt(szProvidedKey, szProvidedKey, szProvidedSalt) Then
'            ' /* key and/or salt invalid */
'            Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING)
'            Exit Function
'
'        End If
'
'        Debug.Print "type=" & szHashType & " key=" & szProvidedKey & " salt=" & szProvidedSalt
'
'    End If
'
'    ' /* check invalid cases */
'
'    If (szExpectedKey = "") And (szProvidedKey <> "") Then
'        ' /* invalid case: no password set but one provided by sender (we don't tell the sender this, however) */
'        Debug.Print "uSNP3Translate(): no system password set but request contained one"
'        Response = snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE)
'        Exit Function
'
'    End If
'
'    If (szExpectedKey <> "") And (szProvidedKey = "") Then
'        ' /* invalid case: password set but none provided by sender (we don't tell the sender this, however) */
'        Debug.Print "uSNP3Translate(): password required"
'        Response = snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE)
'        Exit Function
'
'    End If
'
'    ' /* valid combos */
'
'    If (szExpectedKey <> "") And (szProvidedKey <> "") Then
'        ' /* compare supplied password with ours */
'
'        szExpectedKey = szExpectedKey & szProvidedSalt
'
'        Select Case szHashType
'        Case "MD5"
'            szExpectedKey = MD5DigestStrToHexStr(szExpectedKey)
'
'        Case "SHA1"
'            szExpectedKey = SHA1DigestStrToHexStr(szExpectedKey)
'
'        Case "SHA256"
'            SHA256Init
'            szExpectedKey = SHA256DigestStrToHexStr(szExpectedKey)
'
'        Case "SHA384", "SHA512"
'            Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING, "")
'            Exit Function
'
'        Case Else
'            ' /* unknown hashing algorithm */
'            Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING, "")
'            Exit Function
'
'        End Select
'
'    End If
'
'
''Dim szEncType As String
''Dim szProvidedEncKey As String
''
''        If ppdHeader.Count > 1 Then
''            ' /* get encryption algorithm */
''            ppdHeader.EntryAt 2, szEncType, szProvidedEncKey
''            Debug.Print "CYPHER ALGORITHM=" & szEncType & " value=" & szProvidedEncKey
''
''            Select Case szEncType
''            Case "AES"
''                Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "AES not currently implemented")
''                Exit Function
''
''            Case "DES"
''                Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "DES not currently implemented")
''                Exit Function
''
''            Case "3DES"
''                Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "Triple-DES not currently implemented")
''                Exit Function
''
''            Case Else
''                Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "")
''                Exit Function
''
''            End Select
''
''        End If
'
'
'
'
'    ' /* check key */
'
'    Debug.Print "expected key: " & szExpectedKey
'
'    If szExpectedKey <> szProvidedKey Then
'        Response = snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE, "Authorization failure")
'        Exit Function
'
'    End If
'
'    szr = ""
'
'    ' /* parse actions */
'
'    With Request
'        For i = 2 To .Count - 1
'            .EntryAt i, sz, ""
'
'            If (g_SafeLeftStr(sz, 1) <> "#") Then
'                n = g_DoV42Request(sz, GetCurrentProcessId(), theSocket, SN_NF_IS_SNP3)
'                If n < 0 Then
''                        bErr = True     ' // set overall response flag
'                    n = Abs(n)      ' // convert to SNP code
'
'                Else
'                    n = 0           ' // convert to SNP code
'
'                End If
'
'                szr = szr & uGetAction(sz) & ": " & CStr(n) & " " & snp3_StatusName(n) & vbCrLf
'
'            End If
'        Next i
'    End With
'
'    Response = snp3_BuildResponse(IIf(bErr, SNARL_ERROR_FAILED, SNARL_SUCCESS), , szr)
'
'    ' /* leave connected irrespective of individual action success/failure */
'
'    uSNP3Translate = True
'
'End Function

Public Function snp3_Translate(ByVal RawPacket As String, ByRef SendingSocket As CSocket, ByVal Flags As SN_NOTIFICATION_FLAGS) As Boolean

    If ISNULL(SendingSocket) Then _
        Exit Function

Static i As Long

    ' /* return False if the socket should be closed */

    i = InStr(RawPacket, "SNP/")
    If i = 0 Then
        SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_BAD_PACKET, "Missing header")
        Exit Function

    End If

    RawPacket = g_SafeRightStr(RawPacket, Len(RawPacket) - (i - 1))

    i = InStr(RawPacket, vbCrLf & "END" & vbCrLf)
    If i = 0 Then
        SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_BAD_PACKET, "Missing end marker")
        Exit Function

    End If

    ' /* trim off the surplus */

    RawPacket = g_SafeLeftStr(RawPacket, i - 1)
    
    Debug.Print "**"
    Debug.Print RawPacket
    Debug.Print "**"

Dim ppd As BPackedData

    Set ppd = New BPackedData
    If Not ppd.SetTo(RawPacket, vbCrLf, "") Then
        SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_BAD_PACKET, "Invalid data found")
        Exit Function

    End If

Dim sz As String

    ' /* get the header */

    ppd.EntryAt 1, sz, ""
    If Not g_BeginsWith(sz, "SNP/3.0") Then
        SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_VERSION, "Invalid version in header")
        Exit Function

    End If

    If uIsResponseHeader(sz) Then
        Debug.Print "snp3_Translate(): is SNP/3.0 response"
        Exit Function

    End If

Dim ppdHeader As BPackedData

    ' /* split header into component parts */

    Set ppdHeader = New BPackedData
    ppdHeader.SetTo sz, " ", ":"

Dim n As Long

    ' /* check for empty message */

    With ppd
        If .Count > 1 Then
            For i = 2 To .Count
                .EntryAt i, sz, ""
                If (g_SafeLeftStr(sz, 1) <> "#") Then _
                    n = n + 1

            Next i
        End If

    End With

    If n = 0 Then
        SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_NO_ACTIONS_PROVIDED, "Must supply at least one action")
        Exit Function

    End If


Dim szExpectedKey As String

    ' /* get current password */

    szExpectedKey = g_GetPassword()

    ' /* decode header - currently SNP/3.0 [HASH:value.salt] [ENCRYPTION:value] */

Dim szProvidedKey As String
Dim szProvidedSalt As String
Dim szHashType As String

    If ppdHeader.Count > 1 Then
        ' /* get provided hashing algorithm, key and salt */
        ppdHeader.EntryAt 2, szHashType, szProvidedKey

        If Not uGetSalt(szProvidedKey, szProvidedKey, szProvidedSalt) Then
            ' /* key and/or salt invalid */
            SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING)
            Exit Function

        End If

        Debug.Print "type=" & szHashType & " key=" & szProvidedKey & " salt=" & szProvidedSalt

    End If

    ' /* check invalid cases */

    If (szExpectedKey = "") And (szProvidedKey <> "") Then
        ' /* invalid case: no password set but one provided by sender (we don't tell the sender this, however) */
        Debug.Print "uSNP3Translate(): no system password set but request contained one"
        SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE)
        Exit Function

    End If

    If (szExpectedKey <> "") And (szProvidedKey = "") Then
        ' /* invalid case: password set but none provided by sender (we don't tell the sender this, however) */
        Debug.Print "uSNP3Translate(): password required"
        SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE)
        Exit Function

    End If

    ' /* valid combos */

    If (szExpectedKey <> "") And (szProvidedKey <> "") Then
        ' /* compare supplied password with ours */

        szExpectedKey = szExpectedKey & szProvidedSalt

        Select Case szHashType
        Case "MD5"
            szExpectedKey = MD5DigestStrToHexStr(szExpectedKey)

        Case "SHA1"
            szExpectedKey = SHA1DigestStrToHexStr(szExpectedKey)

        Case "SHA256"
            SHA256Init
            szExpectedKey = SHA256DigestStrToHexStr(szExpectedKey)

        Case "SHA384", "SHA512"
            SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING, "")
            Exit Function

        Case Else
            ' /* unknown hashing algorithm */
            SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING, "")
            Exit Function

        End Select

    End If


'Dim szEncType As String
'Dim szProvidedEncKey As String
'
'        If ppdHeader.Count > 1 Then
'            ' /* get encryption algorithm */
'            ppdHeader.EntryAt 2, szEncType, szProvidedEncKey
'            Debug.Print "CYPHER ALGORITHM=" & szEncType & " value=" & szProvidedEncKey
'
'            Select Case szEncType
'            Case "AES"
'                SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "AES not currently implemented")
'                Exit Function
'
'            Case "DES"
'                SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "DES not currently implemented")
'                Exit Function
'
'            Case "3DES"
'                SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "Triple-DES not currently implemented")
'                Exit Function
'
'            Case Else
'                SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "")
'                Exit Function
'
'            End Select
'
'        End If




    ' /* check key */

    Debug.Print "expected key: " & szExpectedKey

    If szExpectedKey <> szProvidedKey Then
        SendingSocket.SendData snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE, "Authorization failure")
        Exit Function

    End If

Dim szr As String

    szr = ""

    ' /* parse actions */

    With ppd
        For i = 2 To .Count
            .EntryAt i, sz, ""

            If (g_SafeLeftStr(sz, 1) <> "#") Then
                n = g_DoV42Request(sz, GetCurrentProcessId(), SendingSocket, Flags)
                If n < 0 Then
'                        bErr = True     ' // set overall response flag
                    n = Abs(n)      ' // convert to SNP code

                Else
                    n = 0           ' // convert to SNP code

                End If

                szr = szr & "result: " & uGetAction(sz) & " " & CStr(n) & " " & snp3_StatusName(n) & vbCrLf

            End If
        Next i
    End With

Dim bErr As Boolean

    SendingSocket.SendData snp3_BuildResponse(IIf(bErr, SNARL_ERROR_FAILED, SNARL_SUCCESS), , szr)

    ' /* leave connected irrespective of individual action success/failure */

    snp3_Translate = True

End Function

Public Function snp3_BuildResponse(ByVal StatusCode As SNARL_STATUS_CODE, Optional ByVal Hint As String, Optional ByVal Actions As String) As String
Dim sz As String

    sz = "SNP/" & SNP_VERSION & " " & IIf(StatusCode = SNARL_SUCCESS, "OK", "FAILED") & vbCrLf

    ' /* only if StatusCode indicates error */

    If (StatusCode <> SNARL_SUCCESS) And (Actions = "") Then
        sz = sz & "error-code: " & CStr(StatusCode) & vbCrLf & _
                  "error-name: " & snp3_StatusName(StatusCode) & vbCrLf

        If Hint <> "" Then _
            sz = sz & "error-hint: " & Hint & vbCrLf

    End If

    If Actions <> "" Then _
        sz = sz & Actions

    ' /* standard headers */

'    sz = sz & "--" & vbCrLf
    sz = sz & "x-timestamp: " & Format$(Now(), "d mmm yyyy hh:mm:ss") & vbCrLf
    sz = sz & "x-daemon: " & "Snarl " & CStr(APP_VER) & "." & CStr(APP_SUB_VER) & IIf(APP_SUB_SUB_VER <> 0, "." & CStr(APP_SUB_SUB_VER), "") & vbCrLf
    sz = sz & "x-host: " & LCase$(g_GetComputerName()) & vbCrLf
    sz = sz & "END" & vbCrLf

    snp3_BuildResponse = sz

End Function

Public Function snp3_StatusName(ByVal StatusCode As SNARL_STATUS_CODE) As String
Dim sz As String

    Select Case StatusCode

    Case SNARL_SUCCESS
        sz = "Ok"

    Case SNARL_ERROR_FAILED                      '// miscellaneous failure
        sz = "Failed"

    Case SNARL_ERROR_UNKNOWN_COMMAND             '// specified command not recognised
        sz = "BadCommand"

    Case SNARL_ERROR_TIMED_OUT                   '// Snarl took too long to respond
        sz = "TimedOut"

    Case SNARL_ERROR_BAD_SOCKET                  '// invalid socket (or some other socket-related error)
        sz = "BadSocket"

    Case SNARL_ERROR_BAD_PACKET                  '// badly formed request
        sz = "BadPacket"

    Case SNARL_ERROR_INVALID_ARG                 '// R2.4B4: arg supplied was invalid
        sz = "InvalidArg"

    Case SNARL_ERROR_ARG_MISSING                 '// required argument missing
        sz = "ArgMissing"

    Case SNARL_ERROR_SYSTEM                      '// internal system error
        sz = "InternalError"

    Case SNARL_ERROR_ACCESS_DENIED               '// libsnarl only
        sz = "AccessDenied"

    Case SNARL_ERROR_UNSUPPORTED_VERSION
        sz = "UnsupportedVersion"

    Case SNARL_ERROR_NO_ACTIONS_PROVIDED
        sz = "NothingToDo"

    Case SNARL_ERROR_UNSUPPORTED_ENCRYPTION
        sz = "UnsupportedEncryption"

    Case SNARL_ERROR_UNSUPPORTED_HASHING
        sz = "UnsupportedHashing"



    Case SNARL_ERROR_NOT_RUNNING                 '// Snarl handling window not found
        sz = "NotRunning"

    Case SNARL_ERROR_NOT_REGISTERED
        sz = "NotRegistered"

    Case SNARL_ERROR_ALREADY_REGISTERED          '// not used yet; sn41RegisterApp() returns existing token
        sz = "SpuriousRegister"

    Case SNARL_ERROR_CLASS_ALREADY_EXISTS        '// not used yet
        sz = "SpuriousClass"

    Case SNARL_ERROR_CLASS_BLOCKED
        sz = "ClassBlocked"

    Case SNARL_ERROR_CLASS_NOT_FOUND
        sz = "InvalidClass"

    Case SNARL_ERROR_NOTIFICATION_NOT_FOUND
        sz = "InvalidNotification"

    Case SNARL_ERROR_FLOODING                    '// notification generated by same class within quantum
        sz = "FloodingAlert"

    Case SNARL_ERROR_DO_NOT_DISTURB              '// DnD mode is in effect was not logged as missed
        sz = "DoNotDisturb"

    Case SNARL_ERROR_COULD_NOT_DISPLAY           '// not enough space on-screen to display notification
        sz = "DisplayFailed"

    Case SNARL_ERROR_AUTH_FAILURE                '// password mismatch
        sz = "AuthenticationFailure"

    Case SNARL_ERROR_DISCARDED                   '// discarded for some reason, e.g. foreground app match
        sz = "WasDiscarded"

    Case SNARL_ERROR_NOT_SUBSCRIBED                 '// 2.4.2 DR3: subscriber not found
        sz = "NotSubscribed"

    Case SNARL_ERROR_ALREADY_SUBSCRIBED
        sz = "AlreadySubscribed"


    Case Else
        sz = "(Unknown)"

    End Select

    snp3_StatusName = sz

End Function

Private Function uGetSalt(ByVal SaltedKey As String, ByRef Key As String, ByRef Salt As String) As Boolean
Dim i As Long

    i = InStr(SaltedKey, ".")
    If i = 0 Then _
        Exit Function

    Key = g_SafeLeftStr(SaltedKey, i - 1)
    Salt = g_SafeRightStr(SaltedKey, Len(SaltedKey) - i)

    uGetSalt = ((Key <> "") And (Salt <> ""))

End Function

Private Function uGetAction(ByVal Request As String) As String
Dim i As Long

    i = InStr(Request, "?")
    If i Then
        uGetAction = g_SafeLeftStr(Request, i - 1)

    Else
        uGetAction = Request

    End If

End Function

Private Function uIsResponseHeader(ByVal Header As String) As Boolean

    uIsResponseHeader = (InStr(Header, " OK") > 0) Or (InStr(Header, " FAILED") > 0) Or (InStr(Header, " CALLBACK") > 0)

End Function

Public Function snp_CreateForward(ByRef Info As T_NOTIFICATION_INFO) As String

    ' /* base content */

    snp_CreateForward = "SNP/3.0" & vbCrLf & _
                        "register?app-sig=" & Info.ClassObj.App.Signature & "&title=" & Info.ClassObj.App.Name & _
                        uTranslateIcon(Info.ClassObj.App.Icon) & vbCrLf

    ' /* add classes */

Dim i As Long

    With Info.ClassObj.App
        If .CountAlerts Then
            For i = 1 To .CountAlerts
                With .AlertAt(i)
                    snp_CreateForward = snp_CreateForward & "addclass?app-sig=" & Info.ClassObj.App.Signature & "&id=" & .Name & "&name=" & .Description & vbCrLf

                End With
            Next i
        End If
    End With

    ' /* build notification content */

    snp_CreateForward = snp_CreateForward & "notify?app-sig=" & Info.ClassObj.App.Signature

Dim szn As String
Dim szv As String

    With New BPackedData
        .SetTo Info.OriginalContent
        .Rewind
        Do While .GetNextItem(szn, szv)
            Select Case szn
            Case "icon"
                snp_CreateForward = snp_CreateForward & uTranslateIcon(szv)

            Case Else
                snp_CreateForward = snp_CreateForward & "&" & szn & "=" & Replace$(szv, vbCrLf, "\n")

            End Select
        Loop
    End With

    snp_CreateForward = snp_CreateForward & vbCrLf & "END" & vbCrLf

'    MsgBox snp_CreateForward

End Function

Private Function uTranslateIcon(ByVal Icon As String) As String

    If Icon = "" Then _
        Exit Function

Dim sz As String
Dim b As String

    If (g_SafeLeftStr(Icon, 1) = "!") Or (g_IsURL(Icon)) Then
        ' /* add verbatim */
        uTranslateIcon = "&icon=" & Icon

    ElseIf g_ConfigGet("include_icon_when_forwarding") = "1" Then
        ' /* encode it in a slightly modified Base64 format (CRLF's are replaced with #'s) */
        If uEncodeIcon(Icon, b) Then _
            uTranslateIcon = "&icon-phat64=" & b

    End If

End Function

Private Function uEncodeIcon(ByVal IconPath As String, ByRef Base64 As String) As Boolean

    If Not g_Exists(IconPath) Then _
        Exit Function

Dim sz As String
Dim i As Integer

    On Error Resume Next

    i = FreeFile()

    err.Clear
    Open IconPath For Binary Access Read Lock Write As #i
    If err.Number = 0 Then
        sz = String$(LOF(i), Chr$(0))
        Get #i, , sz
        Close #i

        sz = Encode64orig(sz)                   ' // encode as standard Base64
        If sz <> "" Then
            Base64 = Replace$(sz, vbCrLf, "#")  ' // replace CRLFs
            Base64 = Replace$(Base64, "=", "%") ' // replace ='s
            uEncodeIcon = True

        End If

    Else
        g_Debug "TSubscriberRoster.uEncodeIcon(): " & err.Description, LEMON_LEVEL_CRITICAL

    End If

End Function

Public Function snp_CreateSubscription(ByVal Name As String, Optional ByVal Password As String) As String
Dim sz As String

    ' /* base content */

    snp_CreateSubscription = "SNP/3.0" & vbCrLf & "subscribe"
                             
    With New BPackedData
        .SetTo "", "&", "="

        If Name <> "" Then _
            .Add "name", Name

        If Password <> "" Then _
            .Add "password", Password

        sz = .AsString()
        If sz <> "" Then _
            snp_CreateSubscription = snp_CreateSubscription & "?" & sz

    End With

    snp_CreateSubscription = snp_CreateSubscription & vbCrLf & "END" & vbCrLf

End Function

Public Function snp3_IsResponse(ByVal RawData As String, Optional ByRef ResponseType As String) As Boolean
Dim sz As String
Dim s() As String

    sz = uGetFirstLine(RawData)
    If uIsResponseHeader(sz) Then
        s() = Split(sz, " ")
        ResponseType = s(1)
        snp3_IsResponse = True

    End If

End Function

Private Function uGetFirstLine(ByVal str As String, Optional ByVal EndMarker As String = vbCrLf) As String
Dim i As Long

    i = InStr(str, EndMarker)
    If i Then _
        uGetFirstLine = g_SafeLeftStr(str, i - 1)

End Function

