VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSOSHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Key
'/A Required argument
'/F Final argument in list
'/K Keyword must be entered with argument
'/M Multiple arguments
'/N Number
'/S Switch (optional)

Dim WithEvents theSender As CSocket
Attribute theSender.VB_VarHelpID = -1

Implements BDoForEach

Private Function BDoForEach_DoForEach(Item As melon4.BObject, Optional ByVal Arg2 As Long = 0&) As Boolean
Dim pn As TNotification
Dim ps As TSubscriber
Dim sArgs As String
Dim sCmd As String
Dim pt As BTagItem
Dim pl As BTagList
Dim sz As String
Dim pa As TApp
Dim ppa As T_SNARL_APP
Dim i As Long

    If (Item Is Nothing) Then
        ' /* SOS event */
        Select Case Arg2
        Case SOS_WARM_BOOT
            SOS_printline "\nSnarl Debug 1.0\n" & _
                          "(C) 2011 full phat products\n\n'help' -- command list; '!' -- hide the debugger\n\n"

        End Select

    Else
        ' /* args received -- inside Item->Name */

        i = InStr(Item.Name, " ")
        If i Then
            sCmd = g_SafeLeftStr(Item.Name, i - 1)
            sArgs = g_SafeRightStr(Item.Name, Len(Item.Name) - i)

        Else
            sCmd = Item.Name

        End If

        Select Case LCase$(sCmd)

        Case "help"
            SOS_printline "\nHelp Page\n---------"
            SOS_printline "about  - displays version information"
            SOS_printline "apps   - list registered apps"
            SOS_printline "do     - process API-based request"
            SOS_printline "exts   - list installed extensions"
            SOS_printline "quit   - quit Snarl"
            SOS_printline "rel    - release a banned application"
            SOS_printline "list   - display notification information"
            SOS_printline "subs   - list subscribers"
            SOS_printline "xclose - close the local TCP connection"
            SOS_printline "xopen  - open a TCP connection on localhost"
            SOS_printline "xsend  - send local TCP data"
            SOS_printline ""

        Case "about"
            sz = "Snarl " & App.Comments & " (V" & App.Major & "." & App.Revision & ")"
            SOS_printline "\n" & sz
            SOS_printline String$(Len(sz), "`")
            SOS_printline "Started:  " & CStr(gStartTime)
            SOS_printline "Local IP: " & get_ip_address_table()
            SOS_printline "Handler:  $" & g_HexStr(ghWndMain)

            If Not (g_AppRoster Is Nothing) Then
                SOS_printline "Apps:     " & CStr(g_AppRoster.CountApps)

            Else
                SOS_printline "** Application Roster not available **"

            End If

            If Not (g_ExtnRoster Is Nothing) Then
                SOS_printline "Exts:     " & CStr(g_ExtnRoster.CountItems)

            Else
                SOS_printline "** Extension Roster not available **"

            End If

            SOS_printline ""


        Case "app"
            If (sArgs = "") Or (sArgs = "?") Then
                SOS_printline "TOKEN/N/A\n"

            ElseIf (g_AppRoster Is Nothing) Then
                SOS_printline "\n**App Roster not available\n"

            Else
                If g_AppRoster.PrivateFindByToken(g_SafeLong("&H" & sArgs), pa) Then
                    SOS_printline ""
                    With pa
                        .PrivateGetInfo ppa
                        SOS_printline "  Signature: " & .Signature
                        SOS_printline "       Name: " & .Name
                        SOS_printline "      Owner: " & ppa.Pid
                        SOS_printline "      Reply: $" & g_HexStr(ppa.hWnd) & "/$" & g_HexStr(ppa.uMsg, 4)
                        SOS_printline " Registered: " & CStr(ppa.Timestamp)
                        SOS_printline "    IsNew(): " & CStr(.IsNew)
                        SOS_printline " IsSecure(): " & CStr(.IsSecure)
                        SOS_printline " IsBanned(): " & CStr(.IsBanned)
                        SOS_printline "  IsLocal(): " & CStr(.IsLocal)
                        SOS_printline " IsRemote(): " & CStr(.IsRemote)
                        SOS_printline " RemoteHost: " & una(.RemoteHostName)
                        If Not (ppa.Socket Is Nothing) Then _
                            SOS_printline "     Socket: " & ppa.Socket.RemoteHostIP & ":" & CStr(ppa.Socket.RemotePort)
                        SOS_printline "IsDaemon(): " & CStr(.IsDaemon)
                        SOS_printline "       Tool: " & una(ppa.Tool)
                        SOS_printline "       Hint: " & una(ppa.Hint)

                    End With
                    SOS_printline ""

                Else
                    SOS_printline "Not found\n"

                End If
            End If


        Case "apps"
            SOS_printline ""
            With g_AppRoster
                For i = 1 To .CountApps
                    With .AppAt(i)
                        SOS_printline g_HexStr(.Token, 8) & " " & _
                                      IIf(.Password <> "", "S", ".") & _
                                      IIf(.IsLocal, "L", IIf(.IsRemote, "R", ".")) & _
                                      IIf(.IsNew, "N", ".") & _
                                      IIf(.IsBanned, "X", ".") & _
                                      IIf(.hWnd <> 0, "W", ".") & _
                                      IIf(.uMsg <> 0, "M", ".") & _
                                      " " & _
                                      g_Pad(.Signature, 36, True) & " " & g_RightPad(CStr(.CountAlerts), 3, True) & " " & .Name

                    End With
                Next i
            End With
            SOS_printline ""



        Case "err"
            If (sArgs = "") Or (sArgs = "?") Then
                SOS_printline "ERR/N/A\n"

            Else
                SOS_printline uError(g_SafeLong(sArgs)) & "\n"

            End If


        Case "exts"
            SOS_printline ""
            With g_ExtnRoster
                For i = 1 To .CountItems
                    With .ExtensionAt(i)
                        SOS_printline IIf(.State = SN_ES_FAILED, "Failed  ", IIf(.State = SN_ES_LOADED, "Loaded  ", "Unloaded")) & "  " & g_Pad(.Class, 24, True) & .VerString

                    End With
                Next i
            End With
            SOS_printline ""


        Case "quit"
            ExitProcess 0
            Exit Function


        Case "do"
            sArgs = g_RemoveQuotes(sArgs)
            If sArgs = "" Then
                SOS_printline "REQUEST/S/A\n"

            Else
                i = g_DoV42Request(g_RemoveQuotes(sArgs), GetCurrentProcessId())
                SOS_printline CStr(i) & "\n"

            End If


        Case "ban"
            If (sArgs = "") Or (sArgs = "?") Then
                SOS_printline "APPTOKEN/N/A\n"

            Else
                With g_AppRoster
                    If .FindByToken(g_SafeLong("&H" & sArgs), pa, "") Then
                        If Not pa.IsBanned Then
                            pa.SetBlocked True
                            SOS_printline sArgs & " now banned\n"
    
                        Else
                            SOS_printline sArgs & " is not banned\n"
    
                        End If
                    
                    Else
                        SOS_printline "not found\n"
    
                    End If
                End With
            End If


        Case "rel"
            If (sArgs = "") Or (sArgs = "?") Then
                SOS_printline "APPTOKEN/N/A\n"

            Else
                With g_AppRoster
                    If .FindByToken(g_SafeLong("&H" & sArgs), pa, "") Then
                        If pa.IsBanned Then
                            pa.SetBlocked False
                            SOS_printline sArgs & " now released\n"
    
                        Else
                            SOS_printline sArgs & " is not banned\n"
    
                        End If

                    Else
                        SOS_printline "not found\n"

                    End If
                End With
            End If


        Case "list"
            Select Case sArgs
            Case "missed"
                Set pl = g_NotificationRoster.MissedList()

            Case "active"
                Set pl = g_NotificationRoster.ActiveList()

            Case "wait"
                Set pl = g_NotificationRoster.WaitList()

            Case Else
                SOS_printline "MISSED/K ACTIVE/K WAIT/K\n"
                Exit Function

            End Select

            SOS_printline ""
            i = 0

            With pl
                .Rewind
                Do While .GetNextTag(pn) = B_OK
                    With pn
                        SOS_printline g_HexStr(.Info.Token) & " t:" & Format$(.Info.Timeout, "00") & " " & g_Pad(IIf(.Info.Title <> "", .Info.Title, "<n/a>"), 12, True) & " " & g_Pad(.Info.Text, 24, True)
                        i = i + 1

                    End With
                Loop

            End With

            SOS_printline CStr(i) & " listed\n"


        Case "xopen"
            If (theSender Is Nothing) Then
                If (sArgs = "") Or (sArgs = "?") Then
                    SOS_printline "PORT/N/A\n"

                Else
                    i = g_SafeLong(sArgs)
                    Select Case i
                    Case GNTP_DEFAULT_PORT, SNP_DEFAULT_PORT, JSON_DEFAULT_PORT
                        Set theSender = New CSocket
                        theSender.Connect "127.0.0.1", i
                        SOS_printline "Ok\n"
    
                    Case Else
                        SOS_printline "Invalid port\n"
    
                    End Select
                End If
            Else
                SOS_printline "Socket in use\n"

            End If


        Case "xclose"
            If (theSender Is Nothing) Then
                SOS_printline "Socket not open\n"

            Else
                theSender.CloseSocket
                Set theSender = Nothing
                SOS_printline "Socket closed\n"

            End If


        Case "xsend"
            If (theSender Is Nothing) Then
                SOS_printline "Socket not open\n"

            ElseIf theSender.State <> sckConnected Then
                SOS_printline "Socket not connected\n"

            Else
                sArgs = g_RemoveQuotes(sArgs)
                theSender.SendData Replace$(sArgs, "\n", vbCrLf) & vbCrLf

            End If


        Case "subs"
            With g_SubsRoster.Subscribers
                .Rewind
                Do While .GetNextTag(ps) = B_OK
                    With ps
                        SOS_printline g_Pad(.Socket.RemoteHostIP & ":" & CStr(.Socket.RemotePort), 24) & " " & _
                                      g_Pad(.Protocol, 6) & _
                                      IIf(.Password <> "", "(secure)", "")

'                        If .Apps.CountItems Then
'                            With .Apps
'                                .Rewind
'                                Do While .GetNextTag(pt) = B_OK
'                                    SOS_printline "     " & pt.Name
'
'                                Loop
'
'                            End With
'                        End If

                        i = i + 1

                    End With
                Loop

            End With

            SOS_printline CStr(i) & " listed\n"

        Case "password"
            If g_IsIDE() Then
                SOS_printline g_GetPassword() & "\n"

            Else
                SOS_printline "unknown command; try ""help""\n"

            End If

        Case "setpassword"
            If g_SetPassword(sArgs) Then
                SOS_printline "Password changed\n"

            Else
                SOS_printline "Invalid password\n"

            End If

        Case "away"
            g_SetPresence SN_PF_AWAY_SOS
            SOS_printline "Presence changed to away\n"

        Case "back"
            g_ClearPresence SN_PF_AWAY_SOS
            SOS_printline "Presence changed to back\n"

Dim pe As TExtension

        Case "cfgx"
            If g_ExtnRoster.Find(sArgs, pe) Then
'                If Not pe.IsConfigOpen Then
                    SOS_printline "ok\n"
                    pe.DoPrefs 0
                
'                Else
'                    SOS_printline "extension " & g_Quote(sArgs) & ": config already open\n"
'
'                End If

            Else
                SOS_printline "extension " & g_Quote(sArgs) & " not found\n"

            End If

        Case Else
            SOS_printline "unknown command; try ""help""\n"

        End Select

    End If

End Function

Private Sub Class_Initialize()

    If Not (theSender Is Nothing) Then
        theSender.CloseSocket
        Set theSender = Nothing

    End If

End Sub

Private Sub theSender_OnClose()

    SOS_printline "Socket closed\n"
    Set theSender = Nothing

End Sub

Private Sub theSender_OnConnect()

    SOS_printline "Connected\n"

End Sub

Private Sub theSender_OnDataArrival(ByVal bytesTotal As Long)
Dim sz As String

    theSender.GetData sz
    SOS_printline sz

End Sub

Private Function uError(ByVal Error As Long) As String

    Select Case Error
    Case SNARL_SUCCESS:                     uError = "SNARL_SUCCESS"
    Case SNARL_CALLBACK_R_CLICK:            uError = "SNARL_CALLBACK_R_CLICK"
    Case SNARL_CALLBACK_TIMED_OUT:          uError = "SNARL_CALLBACK_TIMED_OUT"
    Case SNARL_CALLBACK_INVOKED:            uError = "SNARL_CALLBACK_INVOKED"
    Case SNARL_CALLBACK_MENU_SELECTED:      uError = "SNARL_CALLBACK_MENU_SELECTED"
    Case SNARL_CALLBACK_M_CLICK:            uError = "SNARL_CALLBACK_M_CLICK"
    Case SNARL_CALLBACK_CLOSED:             uError = "SNARL_CALLBACK_CLOSED"
    Case SNARL_ERROR_FAILED:                uError = "SNARL_ERROR_FAILED"
    Case SNARL_ERROR_UNKNOWN_COMMAND:       uError = "SNARL_ERROR_UNKNOWN_COMMAND"
    Case SNARL_ERROR_TIMED_OUT:             uError = "SNARL_ERROR_TIMED_OUT"
    Case SNARL_ERROR_BAD_SOCKET:            uError = "SNARL_ERROR_BAD_SOCKET"
    Case SNARL_ERROR_BAD_PACKET:            uError = "SNARL_ERROR_BAD_PACKET"
    Case SNARL_ERROR_INVALID_ARG:           uError = "SNARL_ERROR_INVALID_ARG"
    Case SNARL_ERROR_ARG_MISSING:           uError = "SNARL_ERROR_ARG_MISSING"
    Case SNARL_ERROR_SYSTEM:                uError = "SNARL_ERROR_SYSTEM"
    Case SNARL_ERROR_ACCESS_DENIED:         uError = "SNARL_ERROR_ACCESS_DENIED"
    Case SNARL_ERROR_UNSUPPORTED_VERSION:   uError = "SNARL_ERROR_UNSUPPORTED_VERSION"
    Case SNARL_ERROR_NO_ACTIONS_PROVIDED:   uError = "SNARL_ERROR_NO_ACTIONS_PROVIDED"
    Case SNARL_ERROR_UNSUPPORTED_ENCRYPTION:    uError = "SNARL_ERROR_UNSUPPORTED_ENCRYPTION"
    Case SNARL_ERROR_UNSUPPORTED_HASHING:   uError = "SNARL_ERROR_UNSUPPORTED_HASHING"
    Case SNARL_ERROR_NOT_RUNNING:           uError = "SNARL_ERROR_NOT_RUNNING"
    Case SNARL_ERROR_NOT_REGISTERED:        uError = "SNARL_ERROR_NOT_REGISTERED"
    Case SNARL_ERROR_ALREADY_REGISTERED:    uError = "SNARL_ERROR_ALREADY_REGISTERED"
    Case SNARL_ERROR_CLASS_ALREADY_EXISTS:  uError = "SNARL_ERROR_CLASS_ALREADY_EXISTS"
    Case SNARL_ERROR_CLASS_BLOCKED:         uError = "SNARL_ERROR_CLASS_BLOCKED"
    Case SNARL_ERROR_CLASS_NOT_FOUND:       uError = "SNARL_ERROR_CLASS_NOT_FOUND"
    Case SNARL_ERROR_NOTIFICATION_NOT_FOUND:    uError = "SNARL_ERROR_NOTIFICATION_NOT_FOUND"
    Case SNARL_ERROR_FLOODING:              uError = "SNARL_ERROR_FLOODING"
    Case SNARL_ERROR_DO_NOT_DISTURB:        uError = "SNARL_ERROR_DO_NOT_DISTURB"
    Case SNARL_ERROR_COULD_NOT_DISPLAY:     uError = "SNARL_ERROR_COULD_NOT_DISPLAY"
    Case SNARL_ERROR_AUTH_FAILURE:          uError = "SNARL_ERROR_AUTH_FAILURE"
    Case SNARL_ERROR_DISCARDED:             uError = "SNARL_ERROR_DISCARDED"
    Case SNARL_ERROR_NOT_SUBSCRIBED:        uError = "SNARL_ERROR_NOT_SUBSCRIBED"
    Case SNARL_WAS_MERGED:                  uError = "SNARL_WAS_MERGED"

    Case SNARL_NOTIFY_GONE:                 uError = "SNARL_NOTIFY_GONE"
    Case 302:                               uError = "SNARL_NOTIFY_CLICK"
    Case SNARL_NOTIFY_EXPIRED:              uError = "SNARL_NOTIFY_EXPIRED"
    Case SNARL_NOTIFY_INVOKED:              uError = "SNARL_NOTIFY_INVOKED"
    Case SNARL_NOTIFY_MENU:                 uError = "SNARL_NOTIFY_MENU"
    Case 306:                               uError = "SNARL_NOTIFY_EX_CLICK"
    Case SNARL_NOTIFY_CLOSED:               uError = "SNARL_NOTIFY_CLOSED"
    Case SNARL_NOTIFY_ACTION:               uError = "SNARL_NOTIFY_ACTION"
    Case SNARL_NOTIFY_APP_DO_ABOUT:         uError = "SNARL_NOTIFY_APP_DO_ABOUT"
    Case SNARL_NOTIFY_APP_DO_PREFS:         uError = "SNARL_NOTIFY_APP_DO_PREFS"

    Case Else:                              uError = "Undefined error"

    End Select

End Function

Private Function una(ByVal str As String) As String

    una = IIf(str = "", "<n/a>", str)

End Function
