Attribute VB_Name = "mSnarl42_i"
Option Explicit

    ' /*
    '
    '   mSnarl42_i.bas -- Snarl Visual Basic 5/6 header file
    '
    '   Include this module to let your VB5 or VB6 application to talk to Snarl R2.4 or later.
    '
    '   © 2004-2010 full phat products.  All Rights Reserved.
    '
    '        Version: 42 (R2.4)
    '       Revision: 3
    '        Created: 6-Dec-2004
    '   Last Updated: 30-Dec-2010
    '         Author: full phat products
    '        Licence: Simplified BSD License (http://www.opensource.org/licenses/bsd-license.php)
    '
    '   Notes
    '   -----
    '
    '   This include file can be used in conjunction with the Snarl API documentation
    '   (http://www.fullphat.net/dev/api.htm) when porting Snarl support to a different
    '   programming language.  This include file is always the most up-to-date of any
    '   of them as Snarl is natively coded in Visual Basic 6.  (An important point to
    '   note here is that if you're reading this because you downloaded this file as
    '   part of the Snarl CVS then you should use the include file that came with the
    '   latest release of Snarl as that will be the current *supported* one).
    '
    '   As best as possible all functions are documented here, including the local and
    '   supporting ones which are VB-specific.
    '
    '
    '   Revision History
    '   ----------------
    '
    '
    ' */


    ' /* some win32 declares */

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

Private Const SMTO_ABORTIFHUNG = 2
Private Const CP_UTF8 = 65001
Private Const WM_COPYDATA = &H4A

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long

End Type

    ' /* status codes */

Public Enum SNARL_STATUS_CODE
    SNARL_SUCCESS = 0

    ' /* Win32 callbacks (renamed under V42) */

    SNARL_CALLBACK_R_CLICK = 32             '// Deprecated as of V42, ex. SNARL_NOTIFICATION_CLICKED/SNARL_NOTIFICATION_CANCELLED
    SNARL_CALLBACK_TIMED_OUT
    SNARL_CALLBACK_INVOKED                  '// left clicked and no default callback assigned
    SNARL_CALLBACK_MENU_SELECTED            '// HIWORD(wParam) contains 1-based menu item index
    SNARL_CALLBACK_M_CLICK                  '// Deprecated as of V42
    SNARL_CALLBACK_CLOSED                   '//

    ' /* critical errors */

    SNARL_ERROR_FAILED = 101                '// miscellaneous failure
    SNARL_ERROR_UNKNOWN_COMMAND             '// specified command not recognised
    SNARL_ERROR_TIMED_OUT                   '// Snarl took too long to respond
    '//104 gen critical #4
    '//105 gen critical #5
    SNARL_ERROR_BAD_SOCKET = 106            '// invalid socket (or some other socket-related error)
    SNARL_ERROR_BAD_PACKET = 107            '// badly formed request
    SNARL_ERROR_INVALID_ARG = 108           '// R2.4B4: arg supplied was invalid
    SNARL_ERROR_ARG_MISSING = 109           '// required argument missing
    SNARL_ERROR_SYSTEM                      '// internal system error
    '//120 libsnarl critical block
    SNARL_ERROR_ACCESS_DENIED = 121         '// libsnarl only
    '//130 SNP/3.0-specific
    SNARL_ERROR_UNSUPPORTED_VERSION = 131   '// requested SNP version is not supported
    SNARL_ERROR_NO_ACTIONS_PROVIDED         '// empty request
    SNARL_ERROR_UNSUPPORTED_ENCRYPTION      '// requested encryption type is not supported
    SNARL_ERROR_UNSUPPORTED_HASHING         '// requested message hashing type is not supported

    ' /* warnings */

    SNARL_ERROR_NOT_RUNNING = 201           '// Snarl handling window not found
    SNARL_ERROR_NOT_REGISTERED
    SNARL_ERROR_ALREADY_REGISTERED          '// not used yet; sn41RegisterApp() returns existing token
    SNARL_ERROR_CLASS_ALREADY_EXISTS        '// not used yet
    SNARL_ERROR_CLASS_BLOCKED
    SNARL_ERROR_CLASS_NOT_FOUND
    SNARL_ERROR_NOTIFICATION_NOT_FOUND
    SNARL_ERROR_FLOODING                    '// notification generated by same class within quantum
    SNARL_ERROR_DO_NOT_DISTURB              '// DnD mode is in effect was not logged as missed
    SNARL_ERROR_COULD_NOT_DISPLAY           '// not enough space on-screen to display notification
    SNARL_ERROR_AUTH_FAILURE                '// password mismatch
    ' /* R2.4.2 */
    SNARL_ERROR_DISCARDED                   '// discarded for some reason, e.g. foreground app match
    SNARL_ERROR_NOT_SUBSCRIBED              '// 2.4.2 DR3: subscriber not found

    ' /* informational */

    '// code 250 reserved for future use
    SNARL_WAS_MERGED = 251                  '// notification was merged, returned token is the one we merged with

    ' /* callbacks */

    '// code 300 reserved for future use
    SNARL_NOTIFY_GONE = 301                 '// reserved for future use

    ' /* the following are currently specific to SNP 2.0 and are effectively the
    '    Win32 SNARL_CALLBACK_nnn constants with 270 added to them */

'    SNARL_NOTIFY_CLICK = 302              '// indicates notification was right-clicked (deprecated as of V42)
    SNARL_NOTIFY_EXPIRED = 303
    SNARL_NOTIFY_INVOKED = 304              '// note this was "ACK" in a previous life
    SNARL_NOTIFY_MENU                       '// indicates an item was selected from user-defined menu (deprecated as of V42)
'    SNARL_NOTIFY_EX_CLICK                 '// user clicked the middle mouse button (deprecated as of V42)
    SNARL_NOTIFY_CLOSED = 307               '// user clicked the notification's close gadget (GNTP only)

    ' /* the following is generic to SNP and the Win32 API */

    SNARL_NOTIFY_ACTION = 308               '// user picked an action from the list, the data value will indicate which one

End Enum

    ' /* local variables */


    ' /*
    '
    '   RegisterWindowMessage() constant
    '
    ' */

Private Const SNARL_GLOBAL_MSG = "SnarlGlobalEvent"

    ' /*
    '
    '   Global event identifiers
    '   these values appear in wParam.
    '
    ' */

Public Enum SNARL_GLOBAL_EVENTS
    SNARL_BROADCAST_LAUNCHED = 1       ' // Snarl has just started running
    SNARL_BROADCAST_QUIT = 2           ' // Snarl is about to stop running
    SNARL_BROADCAST_STOPPED = 3        ' // sent when stopped by user
    SNARL_BROADCAST_STARTED = 4        ' // sent when started by user
    ' /* R2.4 DR8 */
    SNARL_BROADCAST_USER_AWAY          ' // away mode was enabled
    SNARL_BROADCAST_USER_BACK          ' // away mode was disabled

End Enum

    ' /* application flags */

Public Enum SNARLAPP_FLAGS
    SNARLAPP_HAS_PREFS = 1                      '// application has a UI which Snarl can display
    SNARLAPP_HAS_ABOUT = 2                      '// application has its own About... dialog
    SNARLAPP_IS_WINDOWLESS = &H8000&            '// deprecated

End Enum

    ' /*
    '
    '   application registered message constant
    '
    ' */

Private Const SNARLAPP_MSG = "SnarlAppMessage"

    ' /*
    '
    '   application requests - these values appear in wParam
    '
    ' */

Public Const SNARLAPP_DO_PREFS = 1              '// application should launch its settings UI
Public Const SNARLAPP_DO_ABOUT = 2              '// application should show its About... dialog


' /****************************************************************************************
' /*
' /*
' /*                                Public Win32 API
' /*
' /*
' /****************************************************************************************/

Public Function snDoRequest(ByVal Request As String, Optional ByVal ReplyTimeout As Long = 1000) As Long
Dim hWnd As Long

    ' /* returns zero or a positive value on success, negative value on failure
    '    in the case of failure, the return value will be a negated member of
    '    the SNARL_STATUS_CODE enum - thus ABS(ReturnValue) is required to
    '    correctly identify the error code */

    hWnd = FindWindow("w>Snarl", "Snarl")
    If IsWindow(hWnd) = 0 Then
        snDoRequest = -SNARL_ERROR_NOT_RUNNING
        Exit Function

    End If

    ' /* convert to UTF8 */

    Request = uToUTF8(Request)

    ' /* wrap the request into a COPYDATASTRUCT */

Dim pcds As COPYDATASTRUCT
Dim dw As Long

    With pcds
        .dwData = &H534E4C03            ' // "SNL",3
        .cbData = LenB(Request)
        .lpData = StrPtr(Request)

    End With

    ' /* return zero on failure */

    If SendMessageTimeout(hWnd, WM_COPYDATA, GetCurrentProcessId(), pcds, SMTO_ABORTIFHUNG, ReplyTimeout, dw) <> 0 Then
        snDoRequest = dw

    Else
        ' /* pseudo error: timed out (note the negation of the error code) */
        snDoRequest = -SNARL_ERROR_TIMED_OUT

    End If

End Function

Public Function snBroadcastMsg() As Long

    snBroadcastMsg = RegisterWindowMessage(SNARL_GLOBAL_MSG)

End Function

Public Function snAppMsg() As Long

    snAppMsg = RegisterWindowMessage(SNARLAPP_MSG)

End Function

' /****************************************************************************************
' /*
' /*
' /*                                Internal helper functions
' /*
' /*
' /****************************************************************************************/

Public Function uToUTF8(ByVal str As String) As String

    On Error GoTo ex

    If str = "" Then _
        Exit Function

Dim stBuffer As String
Dim cwch As Long
Dim pwz As Long
Dim pwzBuffer As Long

    pwz = StrPtr(str)
    cwch = WideCharToMultiByte(CP_UTF8, 0, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = WideCharToMultiByte(CP_UTF8, 0, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    uToUTF8 = Left$(stBuffer, cwch - 1)
ex:
End Function

Public Function snIsSnarlRunning() As Boolean

    snIsSnarlRunning = (IsWindow(FindWindow("w>Snarl", "Snarl")) <> 0)

End Function

' /*
'   sn41GetConfigPath() -- Returns a path to Snarl's config folder  (V41)
'
'   Inputs
'       None
'
'   Results
'       Snarl Application registered message.
'
' */

Public Function snGetConfigPath(ByRef Path As String) As Boolean
Dim h As Long

    h = FindWindow("w>Snarl", "Snarl")
    If h = 0 Then _
        Exit Function

Dim lAtom As Long

    lAtom = GetProp(h, "_config_path")
    If lAtom = 0 Then _
        Exit Function

Dim sz As String

    sz = String$(1024, 0)
    h = GetClipboardFormatName(lAtom, sz, Len(sz))
    If h > 0 Then _
        Path = Left$(sz, h) & "etc\"

    snGetConfigPath = (Path <> "")

End Function



' /****************************************************************************************
' /*
' /*
' /*                                Public helper functions
' /*
' /*
' /****************************************************************************************/

Public Function snarl_register(ByVal Signature As String, ByVal Name As String, ByVal Icon As String, Optional ByVal Password As String, Optional ByVal ReplyTo As Long, Optional ByVal Reply As Long, Optional ByVal Flags As SNARLAPP_FLAGS) As Long

    snarl_register = snDoRequest("register?app-sig=" & Signature & "&title=" & Name & "&icon=" & Icon & _
                                 IIf(Password <> "", "&password=" & Password, "") & _
                                 IIf(ReplyTo <> 0, "&reply-to=" & CStr(ReplyTo), "") & _
                                 IIf(Reply <> 0, "&reply=" & CStr(Reply), "") & _
                                 IIf(Flags <> 0, "&flags=" & CStr(Flags), ""))

End Function

Public Function snarl_unregister(ByVal TokenOrSignature As Variant, Optional ByVal Password As String) As Long
Dim sz As String

    sz = "unregister?"

    If VarType(TokenOrSignature) = vbLong Then
        sz = sz & "token=" & CStr(TokenOrSignature)

    ElseIf VarType(TokenOrSignature) = vbString Then
        sz = sz & "app-sig=" & CStr(TokenOrSignature)

    Else
        snarl_unregister = SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    If Password <> "" Then _
        sz = sz & "&password=" & Password

    snarl_unregister = snDoRequest(sz)
    Debug.Print "snarl_unregister: " & snarl_unregister

End Function

Public Function snarl_version() As Long

    snarl_version = snDoRequest("version")

End Function

Public Function snarl_ez_notify(ByVal Signature As String, ByVal Class As String, ByVal Title As String, ByVal Text As String, Optional ByVal Icon As String, Optional ByVal Priority As Long, Optional ByVal Duration As Long = -1, Optional ByVal Password As String) As Long

    snarl_ez_notify = snDoRequest("notify?app-sig=" & Signature & _
                                  "&id=" & Class & _
                                  "&title=" & Title & "&text=" & Text & _
                                  "&priority=" & CStr(Priority) & _
                                  IIf(Duration > -1, "&timeout=" & CStr(Duration), "") & _
                                  IIf(Icon <> "", "&icon=" & Icon, "") & _
                                  IIf(Password <> "", "&password=" & Password, "") _
                                  )

End Function



