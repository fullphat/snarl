Attribute VB_Name = "mSnarlDeprecated_i"
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           mSnarlDeprecated_i.bas
    '/
    '/  Description:    Deprecated functions and declarations
    '/
    '/  © 2004-2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/


    ' /* MSG_QUIT is deprecated in favour of using WM_CLOSE.
    '    Handling of MSG_QUIT is retained in R2.3 and R2.4 */
Public Const MSG_QUIT = WM_USER + 81

    ' /* deprecated - use "hello" action instead */
Public Const WM_SNARLTEST = WM_USER + 237

    ' /* still used and still private but deprecated in favour of using command-line switches */
Public Const WM_MANAGE_SNARL = WM_USER + 238

    ' /* WM_MANAGE_SNARL constants */
Public Enum E_MANAGE_SNARL
    E_MISC_CMDS = 0                 ' // lParam is command
                                    ' // 1 = enable sticky notifications
                                    ' // 2 = show missed panel

    E_STOP_SNARL = 1
    E_START_SNARL = 2
    E_RECYCLE_SNARL = 3
    E_SHOW_PREFS = 4
    E_RELOAD_EXTS = 5
    E_UNLOAD_EXTS = 6
    E_LOAD_EXTS = 7
    E_UNLOAD_EXT = 8
    E_RELOAD_CONFIG = 9
    E_MANAGE_STYLE_ROSTER = 10      ' // lParam is control value: 1=open, 2=close
    E_SET_DND_MODE = 11             ' // lParam is control value: 0=disabled, 1=enabled

End Enum

    ' /*
    '
    '   Pre-V41 SNARL_COMMANDS Enumeration
    '
    ' */

Public Enum S_SNARL_COMMANDS
    
    ' /* -------------------------------------------------------------------
    '
    '    Standard commands -- all use a SNARLSTRUCT
    '
    '    -----------------------------------------------------------------*/
    
    SNARL_SHOW = 1
    SNARL_HIDE_COMMAND
    SNARL_UPDATE
    SNARL_IS_VISIBLE
    SNARL_GET_VERSION
    SNARL_REGISTER_CONFIG_WINDOW
    SNARL_REVOKE_CONFIG_WINDOW

    ' /* following introduced in Snarl V37 (R1.6) */

    SNARL_REGISTER_ALERT
    SNARL_REVOKE_ALERT                          '// for future use
    SNARL_GET_REVISION = SNARL_REVOKE_ALERT     '// note dual-use of command value!
    SNARL_REGISTER_CONFIG_WINDOW_2
    SNARL_GET_VERSION_EX
    SNARL_SET_TIMEOUT

    ' /* following introduced in Snarl V39 (R2.1) */

    SNARL_SET_CLASS_DEFAULT
    SNARL_CHANGE_ATTR
    SNARL_REGISTER_APP
    SNARL_UNREGISTER_APP
    SNARL_ADD_CLASS_

    ' /* following were/are private */

    SNARL_LOAD_EXTENSION                '// private (V38.107)
    SNARL_UNLOAD_EXTENSION              '// private (V38.107)
    SNARL_COUNT_NOTIFICATIONS           '// private (V39.12)
    SNARL_PREVIEW_SCHEME                '// private (V39.66)
    SNARL_LOAD_STYLE_ENGINE             '// private (V39.77)
    SNARL_UNLOAD_STYLE_ENGINE           '// private (V39.77)

    ' /* -------------------------------------------------------------------
    '
    '    Extended commands -- all use a SNARLSTRUCTEX
    '
    '    -----------------------------------------------------------------*/

    SNARL_EX_SHOW = &H20
    SNARL_SHOW_NOTIFICATION                '// V39

End Enum

    ' /* original low-level transport */

Private Type SNARLSTRUCT
    Cmd As S_SNARL_COMMANDS     ' // what to do...
    Id As Long                  ' // snarl message id (returned by snShowMessage())
    Timeout As Long             ' // timeout in seconds (0=sticky)
    LngData2 As Long            ' // reserved
    Title As String * 512
    Text As String * 512        ' // VB defines these as wide so they're actually 1024 bytes
    Icon As String * 512

End Type

    ' /* V39 Class attributes */

Public Enum SNARL_ATTRIBUTES
    SNARL_ATTRIBUTE_TITLE = 1
    SNARL_ATTRIBUTE_TEXT
    SNARL_ATTRIBUTE_ICON
    SNARL_ATTRIBUTE_TIMEOUT
    SNARL_ATTRIBUTE_SOUND
    SNARL_ATTRIBUTE_ACK               '// file to run on ACK
    SNARL_ATTRIBUTE_MENU

End Enum

    ' /* V41 Commands */

Public Enum SNARLX41_COMMANDS
    SNARLX41_REGISTER_APP = 1        '// for this command, SNARLMSG->Token is actually the sending app's PID
    SNARLX41_UNREGISTER_APP
    SNARLX41_UPDATE_APP
    SNARLX41_SET_CALLBACK
    SNARLX41_ADD_CLASS
    SNARLX41_REMOVE_CLASS
    SNARLX41_NOTIFY
    SNARLX41_UPDATE_NOTIFICATION
    SNARLX41_HIDE_NOTIFICATION
    SNARLX41_IS_NOTIFICATION_VISIBLE
    SNARLX41_LAST_ERROR              '// deprecated but retained for backwards compatability
'    SNARL42_ADD_ACTION
'    SNARL42_CLEAR_ACTIONS
'    SNARL42_SHOW_REQUEST
'    SNARL42_PARSE

End Enum

    ' /* Notification flags (V41 only) */

Public Enum SNARL41_NOTIFICATION_FLAGS
'    SNARL41_NOTIFICATION_ALLOWS_MERGE = 1
'    SNARL41_NOTIFICATION_AUTO_DISMISS = 2

    XYAZ = 1

End Enum

Public Function g_XCommandStr(ByVal Command As Long) As String

    Select Case Command
    Case SNARL_SHOW
        g_XCommandStr = "SNARL_SHOW"

    Case SNARL_HIDE_COMMAND
        g_XCommandStr = "SNARL_HIDE"

    Case SNARL_UPDATE
        g_XCommandStr = "SNARL_UPDATE"

    Case SNARL_IS_VISIBLE
        g_XCommandStr = "SNARL_IS_VISIBLE"

    Case SNARL_GET_VERSION
        g_XCommandStr = "SNARL_GET_VERSION"

    Case SNARL_REGISTER_CONFIG_WINDOW
        g_XCommandStr = "SNARL_REGISTER_CONFIG_WINDOW"

    Case SNARL_REVOKE_CONFIG_WINDOW
        g_XCommandStr = "SNARL_REVOKE_CONFIG_WINDOW"

    ' /* R1.6 onwards */
    Case SNARL_REGISTER_ALERT
        g_XCommandStr = "SNARL_REGISTER_ALERT"

    Case SNARL_REVOKE_ALERT                  '// for future use
        g_XCommandStr = "SNARL_REVOKE_ALERT"

    Case SNARL_REGISTER_CONFIG_WINDOW_2
        g_XCommandStr = "SNARL_REGISTER_CONFIG_WINDOW_2"



    Case SNARL_LOAD_EXTENSION
        g_XCommandStr = "SNARL_LOAD_EXTENSION"

    Case SNARL_UNLOAD_EXTENSION
        g_XCommandStr = "SNARL_UNLOAD_EXTENSION"

    Case SNARL_LOAD_STYLE_ENGINE
        g_XCommandStr = "SNARL_LOAD_STYLE_ENGINE"
    
    Case SNARL_UNLOAD_STYLE_ENGINE
        g_XCommandStr = "SNARL_UNLOAD_STYLE_ENGINE"


    Case Else
        g_XCommandStr = "unknown: " & CStr(Command)

    End Select

End Function

Public Function xzShowMessage(ByVal Title As String, ByVal Text As String, Optional ByVal Timeout As Long, Optional ByVal IconPath As String, Optional ByVal hWndReply As Long, Optional ByVal uReplyMsg As Long) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_SHOW
        .Title = uToUTF8(Title)
        .Text = uToUTF8(Text)
        .Icon = uToUTF8(IconPath)
        .Timeout = Timeout
        ' /* R0.3 */
        .LngData2 = hWndReply
        .Id = uReplyMsg

    End With

    xzShowMessage = uSend(pss)

End Function

Private Function uSend(pss As SNARLSTRUCT) As Long
Dim hWnd As Long
Dim pcds As COPYDATASTRUCT
Dim dw As Long

    hWnd = FindWindow("w>Snarl", "Snarl")
    If IsWindow(hWnd) <> 0 Then
        pcds.dwData = 2                 '// SNARLSTRUCT
        pcds.cbData = LenB(pss)
        pcds.lpData = VarPtr(pss)
        If SendMessageTimeout(hWnd, WM_COPYDATA, GetCurrentProcessId(), pcds, SMTO_ABORTIFHUNG, 500, dw) > 0 Then
            ' /* worked! */
            uSend = dw

        Else
            ' /* timed-out or failed */
            uSend = &H8000000A         '// M_TIMED_OUT

        End If

    Else
        uSend = &H80000008             '// M_FAILED

    End If

End Function

Public Function uToUTF8(ByVal str As String) As String
Dim stBuffer As String
Dim cwch As Long
Dim pwz As Long
Dim pwzBuffer As Long

    On Error GoTo ex

    If str = "" Then _
        Exit Function

    pwz = StrPtr(str)
    cwch = WideCharToMultiByte(CP_UTF8, 0, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = WideCharToMultiByte(CP_UTF8, 0, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    uToUTF8 = Left$(stBuffer, cwch - 1)

ex:
End Function



'Private Sub lemon_handle(ByRef Data As SOSSTRUCT)
'Dim var() As String
'Dim c As Long
'Dim i As Long
'
'    Select Case Data.Cmd
'    Case LEMON_CONNECTED
'        lemonConOut vbCrLf & "Snarl " & g_Version() & " Monitor 2.0", , True
'        lemonConOut vbCrLf & "Type <help> for command list, <quit> to quit" & vbCrLf
'        lemon_prompt
'
'    Case LEMON_ARGS_RECEIVED
'        c = lemonGetArgs(Data, var)
'        If c Then
'            Select Case var(1)
'            Case "help"
'                lemonConOut vbCrLf & vbCrLf & "Snarl Monitor Commands" & vbCrLf, , , True
'
'                lemonConOut "apps", , True
'                lemonConOut " - Lists registered applications" & vbCrLf
'
'                lemonConOut "info", , True
'                lemonConOut " - Displays Snarl information" & vbCrLf
'
'                lemonConOut "quit", , True
'                lemonConOut " - Quits the monitor" & vbCrLf
'
'                lemonConOut "test", , True
'                lemonConOut " - Displays a Snarl message with the specified information" & vbCrLf
'
'
'            Case "quit"
'                lemonConOut vbCrLf & "bye" & vbCrLf
'                lemonDisconnect
'                Exit Sub
'
'            Case "test"
'                uCmdTest c, var
'
'            Case "info"
'                lemonConOut vbCrLf & vbCrLf & "Snarl Information" & vbCrLf, , , True
'                lemonConOut "Snarl window: 0x" & g_HexStr(mhWnd) & vbCrLf
'                lemonConOut "SNARL_GLOBAL_MSG: 0x" & g_HexStr(snGetGlobalMsg(), 4) & vbCrLf
'                lemonConOut "Connected apps: " & g_AppRoster.CountApps() & vbCrLf
'
'            Case "apps"
'                If Not (g_AppRoster Is Nothing) Then
'                    With g_AppRoster
'                        For i = 1 To .CountApps
'                            With .AppAt(i)
'                                lemonConOut vbCrLf & g_HexStr(.hWnd) & IIf(IsWindow(.hWnd) = 0, "*", " ") & g_RightPad(CStr(.pid), 5) & " " & .Name
'
'                            End With
'                        Next i
'                        lemonConOut vbCrLf & CStr(.CountApps) & " listed"
'
'                    End With
'                Else
'                    lemonConOut vbCrLf & "App Roster not running", , True
'
'                End If
'
'            Case Else
'                lemonConOut vbCrLf & "unknown command '" & var(1) & "'"
'
'            End Select
'
'        End If
'
'        lemon_prompt
'
'    End Select
'
'End Sub
'
'Private Sub lemon_prompt()
'
'    lemonConOut vbCrLf & "s> ", , True
'
'End Sub
'
'Private Sub uCmdTest(ByVal Args As Long, Arg() As String)
'Dim szTitle As String
'Dim szText As String
'Dim dwTimeout As Long
'Dim szIcon As String
'
'    On Error Resume Next            ' // just in case we get a bad timeout value...
'
'    dwTimeout = 10
'
'    If Args > 1 Then
'        szTitle = Arg(2)
'        If Args > 2 Then
'            szText = Arg(3)
'            If Args > 3 Then
'                If Val(Arg(4)) <> 0 Then _
'                    dwTimeout = Val(Arg(4))
'
'                If Args > 4 Then _
'                    szIcon = Arg(5)
'
'            End If
'        End If
'    End If
'
'Dim pInfo As T_NOTIFICATION_INFO
'
'    With pInfo
'        .Title = szTitle
'        .Text = szText
'        .Timeout = dwTimeout
'        .IconPath = szIcon
'
'    End With
'
'    g_NotificationRoster.Add New TAlert, pInfo
'    lemonConOut vbCrLf & "Okay"
'
'End Sub

Public Sub snOldRebootStyleEngine(ByVal StyleEngine As String)
Dim pss As SNARLSTRUCT

    pss.Text = uToUTF8(StyleEngine)

    ' /* stop */
    pss.Cmd = SNARL_UNLOAD_STYLE_ENGINE
    uSend pss

    ' /* start */
    pss.Cmd = SNARL_LOAD_STYLE_ENGINE
    uSend pss

End Sub
