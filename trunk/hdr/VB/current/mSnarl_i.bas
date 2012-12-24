Attribute VB_Name = "mSnarl_i"
Option Explicit

    ' /*
    '
    '   mSnarl_i.bas -- Snarl Visual Basic 5/6 header file
    '
    '   Include this module to let your VB5 or VB6 application to talk to Snarl R2.4 or later.
    '
    '   © 2004-2011 full phat products.  All Rights Reserved.
    '
    '        Version: 43 (R2.5)
    '       Revision: 37
    '        Created: 6-Dec-2004
    '   Last Updated: 19-Sep-2011
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
    SNARL_ERROR_BUSY = 105
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
    SNARL_ERROR_ALREADY_SUBSCRIBED          '//
    ' /* R2.5.1 */
    SNARL_ERROR_ADDON_NOT_FOUND

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

    ' /* other events */

    '//reserved app event 320
    SNARL_NOTIFY_APP_DO_ABOUT = 321
    SNARL_NOTIFY_APP_DO_PREFS
    SNARL_NOTIFY_APP_ACTIVATED
    SNARL_NOTIFY_APP_QUIT

End Enum


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

    ' /* system flags - introduced in V43 - use with snGetSystemFlags() */

Public Enum SNARL_SYSTEM_FLAGS
    SNARL_SF_USER_AWAY = 1
    SNARL_SF_USER_BUSY = 2

    SNARL_SF_DEBUG_MODE = &H80000000

End Enum

    ' /*
    '
    '   application requests - these values appear in wParam
    '
    ' */

Public Const SNARLAPP_DO_PREFS = 1              '// application should launch its settings UI
Public Const SNARLAPP_DO_ABOUT = 2              '// application should show its About... dialog
Public Const SNARLAPP_ACTIVATED = 3             '// V43
Public Const SNARLAPP_QUIT_REQUESTED = 4        '// V43


' /****************************************************************************************
' /*
' /*
' /*                                Public Win32 API
' /*
' /*
' /****************************************************************************************/




' /****************************************************************************************
' /*
' /*                                    Base functions
' /*
' /****************************************************************************************/

' /*
'   snAppMsg() -- Returns Snarl application message  (V41)
'
'   Inputs
'       None
'
'   Results
'       Returns the Snarl application Windows registered message
'
' */
Public Function snAppMsg() As Long

    snAppMsg = RegisterWindowMessage("SnarlAppMessage")

End Function

' /*
'   snDoRequest() -- Send a request to Snarl  (V41)
'
'   Inputs
'       Request: complete request
'       ReplyTimeout: the amount of time (in milliseconds) to wait for a response
'
'   Results
'       Returns whatever Snarl returns or:
'           SNARL_ERROR_NOT_RUNNING:    if Snarl's message handling window isn't found
'           SNARL_ERROR_TIMED_OUT:      if SendMessageTimeout() times out
'
'   Note
'       When returning an error, ensure the value is negated.
'
' */
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

' /*
'   snGetConfigPath() -- Returns a path to Snarl's config folder  (V41)
'
'   Inputs
'       Path is a string to contain the returned path.
'
'   Results
'       TRUE on success, FALSE otherwise.
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

' /*
'   snGetSystemFlags() -- Returns system information  (V43)
'
'   Inputs
'       None
'
'   Results
'       Series of flags from the SNARL_SYSTEM_FLAGS enum or zero if Snarl isn't running.
'
' */
Public Function snGetSystemFlags() As SNARL_SYSTEM_FLAGS
Dim hWnd As Long

    hWnd = FindWindow("w>Snarl", "Snarl")
    If IsWindow(hWnd) <> 0 Then _
        snGetSystemFlags = GetProp(hWnd, "_flags")

End Function

' /*
'   snIsSnarlRunning() -- Determines Snarl state  (V41)
'
'   Inputs
'       None
'
'   Results
'       Returns TRUE if Snarl is running, FALSE otherwise.
'
' */
Public Function snIsSnarlRunning() As Boolean

    snIsSnarlRunning = (IsWindow(FindWindow("w>Snarl", "Snarl")) <> 0)

End Function

' /*
'   snSysMsg() -- Return Snarl's system broadcast message
'
'   Inputs
'       None
'
'   Results
'       Returns Snarl's registered system message
'
' */
Public Function snSysMsg() As Long

    snSysMsg = RegisterWindowMessage("SnarlGlobalEvent")

End Function




' /****************************************************************************************
' /*
' /*                                    Helper functions
' /*
' /****************************************************************************************/

' /*
'   snarl_add_class() -- Add a notification class
'
'   Inputs
'       Signature
'       Id
'       Name
'       Enabled
'       Password
'
'   Results
'       Status code
'
'   Notes
'       Wraps the "addclass" command
'
' */
Public Function snarl_add_class(ByVal Signature As String, ByVal Id As String, ByVal Name As String, Optional ByVal Enabled As Boolean = True, Optional ByVal Password As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Icon As String, Optional ByVal Duration As Long = -1, Optional ByVal Sound As String, Optional ByVal Callback As String) As Long
Dim sz As String

    sz = "addclass?app-sig=" & Signature & "&id=" & Id & "&name=" & Name & "&enabled=" & IIf(Enabled, "1", "0")

    If Password <> "" Then _
        sz = sz & "&password=" & Password

    If Title <> "" Then _
        sz = sz & "&title=" & Title

    If Text <> "" Then _
        sz = sz & "&text=" & Text

    If Icon <> "" Then _
        sz = sz & "&icon=" & Icon

    If Duration <> -1 Then _
        sz = sz & "&duration=" & CStr(Duration)

    If Sound <> "" Then _
        sz = sz & "&sound=" & Sound

    If Callback <> "" Then _
        sz = sz & "&callback=" & Callback

    snarl_add_class = snDoRequest(sz)

End Function

' /*
'   snarl_hide_notification() -- Hide a notification
'
'   Inputs
'       Signature
'       UID
'       Password
'
'   Results
'       Status code
'
'   Notes
'       Wraps the "hide" command
'
' */
Public Function snarl_hide_notification(ByVal Signature As String, ByVal uID As String, Optional ByVal Password As String) As Long

    snarl_hide_notification = snDoRequest("hide?app-sig=" & Signature & "&uid=" & uID & IIf(Password <> "", "&password=" & Password, ""))

End Function

' /*
'   snarl_is_user_away() -- Determines if user is away
'
'   Inputs
'       None
'
'   Result
'       TRUE if user is away, FALSE otherwise
'
'   Notes
'
' */
Public Function snarl_is_user_away() As Boolean

    snarl_is_user_away = ((snGetSystemFlags() And SNARL_SF_USER_AWAY) <> 0)

End Function

' /*
'   snarl_is_user_busy() -- Determines if user is busy
'
'   Inputs
'       None
'
'   Result
'       TRUE if user is busy, FALSE otherwise
'
'   Notes
'
' */
Public Function snarl_is_user_busy() As Boolean

    snarl_is_user_busy = ((snGetSystemFlags() And SNARL_SF_USER_BUSY) <> 0)

End Function

' /*
'   snarl_notify() -- Show a notification
'
'   Inputs
'       Signature
'       Class
'       UID
'       Password
'       Title
'       Text
'       Icon
'       Priority
'       Duration
'       Callback
'       PercentValue
'       CustomData
'
'   Result
'       Status code
'
'   Notes
'       Wraps the "notify" command
'
' */
Public Function snarl_notify(ByVal Signature As String, ByVal Class As String, ByVal uID As String, Optional ByVal Password As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Icon As String, Optional ByVal Priority As Long, Optional ByVal Duration As Long = -1, Optional ByVal Callback As String, Optional ByVal PercentValue As Long = -1, Optional ByVal CustomData As String) As Long
Dim sz As String

    Title = Replace$(Title, "&", "&&")
    Title = Replace$(Title, "=", "==")

    Text = Replace$(Text, "&", "&&")
    Text = Replace$(Text, "=", "==")

    sz = "notify?app-sig=" & Signature & _
         "&id=" & Class & _
         "&title=" & Title & "&text=" & Text & _
         "&priority=" & CStr(Priority) & _
         IIf(Duration > -1, "&timeout=" & CStr(Duration), "") & _
         IIf(Icon <> "", "&icon=" & Icon, "") & _
         IIf(Password <> "", "&password=" & Password, "") & _
         IIf(uID <> "", "&uid=" & uID, "") & _
         IIf(Callback <> "", "&callback=" & Callback, "")

    If (PercentValue >= 0) And (PercentValue <= 100) Then _
        sz = sz & "&value-percent=" & CStr(PercentValue)

    If CustomData <> "" Then _
        sz = sz & "&" & CustomData

    snarl_notify = snDoRequest(sz)

End Function



' /*
'   snarl_register() -- Registers an application
'
'   Inputs
'       Signature
'       Name
'       Icon
'       Password
'       ReplyTo
'       ReplyWith
'       IsDaemon
'
'   Results
'       Status code
'
'   Notes
'       Wraps the "register" command
'
' */
Public Function snarl_register(ByVal Signature As String, ByVal Name As String, ByVal Icon As String, Optional ByVal Password As String, Optional ByVal ReplyTo As Long, Optional ByVal ReplyWith As Long, Optional ByVal IsDaemon As Boolean = False, Optional ByVal Hint As String) As Long

    snarl_register = snDoRequest("register?app-sig=" & Signature & "&title=" & Name & "&icon=" & Icon & _
                                 IIf(Password <> "", "&password=" & Password, "") & _
                                 IIf(ReplyTo <> 0, "&reply-to=" & CStr(ReplyTo), "") & _
                                 IIf(ReplyWith <> 0, "&reply-with=" & CStr(ReplyWith), "") & _
                                 IIf(IsDaemon, "&app-daemon=1", "") & _
                                 IIf(Hint <> "", "&hint=" & Hint, ""))

End Function

' /*
'   snarl_rem_class() -- Removes a notification class
'
'   Inputs
'       Signature
'       Id
'       Name
'       Enabled
'       Password
'
'   Results
'       Status code
'
'   Notes
'       Wraps the "addclass" command
'
' */
Public Function snarl_rem_class(ByVal Signature As String, ByVal Id As String, Optional ByVal Password As String) As Long
Dim sz As String

    sz = "remclass?app-sig=" & Signature & "&id=" & Id
    If Password <> "" Then _
        sz = sz & "&password=" & Password

    snarl_rem_class = snDoRequest(sz)

End Function

' /*
'   snarl_unregister() -- Unregisters an application
'
'   Inputs
'       TokenOrSignature
'       Password
'
'   Results
'       Status code
'
'   Notes
'       Wraps the "unregister" command
'
' */
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
'    Debug.Print "snarl_unregister: " & snarl_unregister

End Function

' /*
'   snarl_version() -- Returns the version of Snarl
'
'   Inputs
'       None
'
'   Results
'       Status code or version of Snarl
'
'   Notes
'       Wraps the "version" command
'
' */
Public Function snarl_version() As Long

    snarl_version = snDoRequest("version")

End Function



' /****************************************************************************************
' /*
' /*                                  VB-specific functions
' /*
' /****************************************************************************************/


' /*
'   create_password() -- Create a password  (V43)
'
'   Inputs
'       Length of password in characters
'
'   Results
'       Returns computed password
'
' */
Public Function create_password(Optional ByVal Chars As Integer = 32) As String
Dim i As Integer

    If Chars > 1 Then
        For i = 1 To Chars
            Randomize Timer
            create_password = create_password & Chr$(Rnd * (255 - 48) + 48)

        Next i

    End If

End Function

' /*
'   make_path() -- Path validator  (V43)
'
'   Inputs
'       Filesystem path
'
'   Results
'       Ensures the provided path ends with a backslash
'
' */
Public Function make_path(ByVal Path As String) As String

    If (Path = "") Then _
        Exit Function

    If Right$(Path, 1) <> "\" Then
        make_path = Path & "\"
    Else
        make_path = Path
    End If

End Function


' /*
'   uToUTF8() -- Convert a string to UTF8
'
'   Inputs
'       str: string to convert
'
'   Results
'       UTF8-encoded string
'
' */
Private Function uToUTF8(ByVal str As String) As String

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




Public Function snarl_is_notification_visible(ByVal Signature As String, ByVal uID As String, Optional ByVal Password As String) As Boolean

    snarl_is_notification_visible = (snDoRequest("isvisible?app-sig=" & Signature & "&uid=" & uID & IIf(Password <> "", "&password=" & Password, "")) = 0)

End Function

Public Function snarl_ez_notify(ByVal Signature As String, ByVal Class As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Icon As String, Optional ByVal Priority As Long, Optional ByVal Duration As Long = -1, Optional ByVal Password As String, Optional ByVal uID As String, Optional ByVal Callback As String, Optional ByVal Percent As Long = -1) As Long
Dim sz As String

    sz = "notify?app-sig=" & Signature & _
         "&id=" & Class & _
         "&title=" & Title & "&text=" & Text & _
         "&priority=" & CStr(Priority) & _
         IIf(Duration > -1, "&timeout=" & CStr(Duration), "") & _
         IIf(Icon <> "", "&icon=" & Icon, "") & _
         IIf(Password <> "", "&password=" & Password, "") & _
         IIf(uID <> "", "&uid=" & uID, "") & _
         IIf(Callback <> "", "&callback=" & Callback, "")

    If (Percent >= 0) And (Percent <= 100) Then _
        sz = sz & "&value-percent=" & CStr(Percent)

    snarl_ez_notify = snDoRequest(sz)

End Function


