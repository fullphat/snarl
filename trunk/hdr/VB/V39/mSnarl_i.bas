Attribute VB_Name = "mSnarl_i"
Option Explicit

    ' /*
    '
    '   mSnarl_i.bas -- Snarl Visual Basic 5/6 include
    '
    '   © 2004-2008 full phat products.  All Rights Reserved.
    '
    '   Include this module if you want your Visual Basic 5 or 6 application to talk to Snarl.
    '
    '        Version: 39 (R2.1)
    '       Revision: 20
    '        Created: 6-Dec-2004
    '   Last Updated: 17-Dec-2008
    '         Author: C. Peel (aka Cheekiemunkie)
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
    '   39.20 (17-Dec-08)
    '       - Changed licence to simplified BSD.
    '
    '   39.6 (3-Dec-08)
    '       - Added missing SNARL_GET_REVISION command
    '
    '   39.5 (28-Nov-08)
    '       - Incorporated new V39 API functions and enums.
    '
    '   39.1 (18-nov-08)
    '       - Bumped to V39
    '
    '   38.3 (10-Nov-08)
    '       - Added snSetAsSnarlApp().
    '
    '   38.2 (11-Apr-08)
    '       - Clarified function return values.
    '
    '   38.1 (19-Mar-08)
    '       - Fixed bug in snGetVersion() which would return True even if uSend() failed.
    '
    '   39 (7-Apr-07)
    '       - Final R1.6 release
    '
    '   36 (14-Mar-07)
    '       - VB-friendly UTF8 string conversions finally sorted out.  This change *only*
    '         affects the VB include (this one) and should *not* impact on existing VB-based
    '         applications.
    '
    '   35 (13-Mar-07)
    '       - snGetAppPath() reworked completely.  Now Snarl itself provides the path through
    '         a 'Static' class child window within the Dispatcher window.
    '
    '   34 (28-Feb-07)
    '       - Added string length constants (taken from Tresni's C++ includes)
    '       - Moved out all unused code/constants/types
    '       - Added new snSetTimeout() function and associated command
    '       - uSend() timeout increased from 100ms to 500ms
    '       - uSend() now returns M_FAILED if Snarl window not found or M_TIMED_OUT if sending timed out
    '
    '   33 (21-Feb-07)
    '       - More commenting
    '
    '   32 (13-Feb-07)
    '       - Added snGetSnarlWindow() -- standardised way of retrieving message handling window
    '       - Added WM_SNARLTEST constant
    '
    '   31
    '       - snRegisterAlert() -- registers an alert for a specific application
    '       - snRegisterConfig2() -- allows external image to be used rather than config window icon
    '       - We now use SendMessageTimeout() instead of SendMessage()
    '       - snRegisterConfig() and snRegisterConfig2() both set a local variable (m_hWndFrom)
    '         with the hWnd parameter passed to them
    '       - snRevokeConfig() clears the m_hwndFrom local variable
    '       - uSend() now passes m_hWndFrom to Snarl in wParam
    '
    ' */



                        ' /* constants */




    ' /* SNARLSTRUCT and SNARLSTRUCTEX string length maximums */

Public Const SNARL_STRING_LENGTH = 1024
Public Const SNARL_UNICODE_LENGTH = SNARL_STRING_LENGTH / 2


    ' /* WM_SNARLTEST message
    '
    '   This can be used for development and test purposes.  When received,
    '   Snarl will display a simple message in order to show that
    '   communication has been established okay.
    '
    '   Parameters for SendMessage() are:
    '
    '       hWnd - Snarl Dispatcher Window (use snGetSnarlWindow())
    '       wParam - Command - see table below
    '       lParam - Depends on wParam
    '
    '       +--------+------------------------------------------------------+
    '       | wParam | lParam                                               |
    '       +--------+------------------------------------------------------+
    '       |    0   | Not used                                             |
    '       |    1   | Value is displayed by the Snarl message              |
    '       +--------+------------------------------------------------------+
    '
    '   Possible return values are:
    '      -1                   - Succeeded
    '       0                   - SendMessage() failed
    '       M_NOT_IMPLEMENTED   - Bad wParam value
    '
    ' */
Public Const WM_SNARLTEST = &H400 + 237     ' // note hardcoded WM_USER value!


    ' /*
    '
    '   Global event identifiers
    '
    '   Identifiers marked with a '*' are sent by Snarl in two ways:
    '       1. As a broadcast message (uMsg = 'SNARL_GLOBAL_MSG')
    '       2. To the window registered in snRegisterConfig() or snRegisterConfig2()
    '          (uMsg = reply message specified at the time of registering)
    '
    '   In both cases these values appear in wParam.
    '
    '   Identifiers not marked are not broadcast; they are simply sent to the application's
    '   registered window.
    '
    ' */

Public Const SNARL_LAUNCHED = 1         ' // Snarl has just started running*
Public Const SNARL_QUIT = 2             ' // Snarl is about to stop running*
Public Const SNARL_ASK_APPLET_VER = 3   ' // (R1.5) Reserved for future use
Public Const SNARL_SHOW_APP_UI = 4      ' // (R1.6) Application should show its UI


    ' /*
    '
    '   Message event identifiers
    '
    '   These are sent by Snarl to the window specified in snShowMessage() when the
    '   Snarl Notification raised times out or the user clicks on it.
    '
    ' */

Public Const SNARL_NOTIFICATION_CLICKED = 32        ' // notification was right-clicked by user
Public Const SNARL_NOTIFICATION_TIMED_OUT = 33
Public Const SNARL_NOTIFICATION_ACK = 34            ' // notification was left-clicked by user
Public Const SNARL_NOTIFICATION_MENU = 35           ' // V39 - menu item selected
Public Const SNARL_NOTIFICATION_MIDDLE_BUTTON = 36  ' // V39 - notification middle-clicked by user
Public Const SNARL_NOTIFICATION_CLOSED = 37         ' // V39 - user clicked the close gadget

    ' /* Added in V37 (R1.6) -- same value, just improved the meaning of it */

Public Const SNARL_NOTIFICATION_CANCELLED = SNARL_NOTIFICATION_CLICKED


    ' /*
    '
    '   SNARL_COMMANDS Enumeration
    '
    ' */

Public Enum SNARL_COMMANDS
    
    ' /* -------------------------------------------------------------------
    '
    '    Standard commands -- all use a SNARLSTRUCT
    '
    '    -----------------------------------------------------------------*/
    
    SNARL_SHOW = 1
    SNARL_HIDE
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
    SNARL_ADD_CLASS

    ' /* -------------------------------------------------------------------
    '
    '    Extended commands -- all use a SNARLSTRUCTEX
    '
    '    -----------------------------------------------------------------*/

    SNARL_EX_SHOW = &H20
    SNARL_SHOW_NOTIFICATION                '// V39

End Enum

Public Enum SNARL_APP_FLAGS
    SNARL_APP_HAS_PREFS = 1
    SNARL_APP_HAS_ABOUT = 2

End Enum

Public Enum SNARL_APP_COMMANDS
    SNARL_APP_SHOW_PREFS = 1
    SNARL_APP_SHOW_ABOUT = 2

End Enum

    ' /* --------------- V39 additions --------------- */


    ' /* snAddClass() flags */
Public Enum SNARL_CLASS_FLAGS
    SNARL_CLASS_ENABLED = 0
    SNARL_CLASS_DISABLED = 1
    SNARL_CLASS_NO_DUPLICATES = 2           '// means Snarl will suppress duplicate notifications
    SNARL_CLASS_DELAY_DUPLICATES = 4        '// means Snarl will suppress duplicate notifications within a pre-set time period

End Enum

    ' /* Class attributes */
Public Enum SNARL_ATTRIBUTES
    SNARL_ATTRIBUTE_TITLE = 1
    SNARL_ATTRIBUTE_TEXT
    SNARL_ATTRIBUTE_ICON
    SNARL_ATTRIBUTE_TIMEOUT
    SNARL_ATTRIBUTE_SOUND
    SNARL_ATTRIBUTE_ACK               '// file to run on ACK
    SNARL_ATTRIBUTE_MENU

End Enum

    ' /* ------------------- end of ------------------ */
    ' /* --------------- V39 additions --------------- */



                        ' /* structures */



    ' /*
    '
    '   SNARLSTRUCT
    '   This is an internal structure used to pass information between Snarl and
    '   registered applications - do not attempt to craft your own messages
    '   using this structure; use the standard sn... api methods instead.
    '
    ' */


Public Type SNARLSTRUCT
    Cmd As SNARL_COMMANDS       ' // what to do...
    Id As Long                  ' // snarl message id (returned by snShowMessage())
    Timeout As Long             ' // timeout in seconds (0=sticky)
    LngData2 As Long            ' // reserved
    Title As String * 512
    Text As String * 512        ' // VB defines these as wide so they're actually 1024 bytes
    Icon As String * 512

End Type

Public Type SNARLSTRUCTEX
    Cmd As SNARL_COMMANDS       ' // what to do...
    Id As Long                  ' // snarl message id (returned by snShowMessage())
    Timeout As Long             ' // timeout in seconds (0=sticky)
    LngData2 As Long            ' // reserved
    Title As String * 512
    Text As String * 512        ' // VB defines these as wide so they're actually 1024 bytes
    Icon As String * 512
    Class As String * 512
    Extra As String * 512
    Extra2 As String * 512
    Reserved1 As Long
    Reserved2 As Long

End Type



    ' /* internal declares */

Private Const WM_COPYDATA = &H4A
Private Const SNARL_GLOBAL_MSG = "SnarlGlobalEvent"
Private Const SNARL_APP_MSG = "SnarlAppMessage"

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long

End Type

Dim m_hwndFrom As Long      ' // set during snRegisterConfig() or snRegisterConfig2()

Private Declare Function lstrcpyW2 Lib "kernel32" Alias "lstrcpyW" (ByVal str1 As Long, ByVal str2 As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Const SMTO_ABORTIFHUNG = 2
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal cbBytes As Long)
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Private Const CP_UTF8 = 65001
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Function snShowMessage(ByVal Title As String, ByVal Text As String, Optional ByVal Timeout As Long, Optional ByVal IconPath As String, Optional ByVal hWndReply As Long, Optional ByVal uReplyMsg As Long) As Long
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

    snShowMessage = uSend(pss)

End Function

Public Function snHideMessage(ByVal Id As Long) As Boolean
Dim pss As SNARLSTRUCT

    pss.Cmd = SNARL_HIDE
    pss.Id = Id
    snHideMessage = CBool(uSend(pss))

End Function

Public Function snIsMessageVisible(ByVal Id As Long) As Boolean
Dim pss As SNARLSTRUCT

    pss.Cmd = SNARL_IS_VISIBLE
    pss.Id = Id
    snIsMessageVisible = CBool(uSend(pss))

End Function

Public Function snUpdateMessage(ByVal Id As Long, ByVal Title As String, ByVal Text As String, Optional ByVal IconPath As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_UPDATE
        .Id = Id
        .Title = uToUTF8(Title)
        .Text = uToUTF8(Text)
        ' /* 1.6 Beta 4 */
        .Icon = uToUTF8(IconPath)

    End With

    snUpdateMessage = uSend(pss)

End Function

Public Function snRegisterConfig(ByVal hWnd As Long, ByVal AppName As String, ByVal ReplyMsg As Long) As Long
Dim pss  As SNARLSTRUCT

    m_hwndFrom = hWnd

    With pss
        .Cmd = SNARL_REGISTER_CONFIG_WINDOW
        .LngData2 = hWnd
        .Id = ReplyMsg
        .Title = uToUTF8(AppName)

    End With

    snRegisterConfig = uSend(pss)

End Function

Public Function snRevokeConfig(ByVal hWnd As Long) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_REVOKE_CONFIG_WINDOW
        .LngData2 = hWnd

    End With

    snRevokeConfig = uSend(pss)
    m_hwndFrom = 0

End Function

Private Function uSend(pss As SNARLSTRUCT) As Long
Dim hWnd As Long
Dim pcds As COPYDATASTRUCT
Dim dw As Long

    hWnd = snGetSnarlWindow()
    If IsWindow(hWnd) <> 0 Then
        pcds.dwData = 2                 '// SNARLSTRUCT
        pcds.cbData = LenB(pss)
        pcds.lpData = VarPtr(pss)
        If SendMessageTimeout(hWnd, WM_COPYDATA, m_hwndFrom, pcds, SMTO_ABORTIFHUNG, 500, dw) > 0 Then
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

Public Function snGetVersion(ByRef Major As Integer, ByRef Minor As Integer) As Boolean
Dim pss As SNARLSTRUCT
Dim hr As Long

    pss.Cmd = SNARL_GET_VERSION
    hr = uSend(pss)
    If (hr And &H80000000) = 0 Then         ' // FIXED: uSend() returns an M_RESULT on error
        Major = uHIWORD(hr)
        Minor = uLOWORD(hr)
        snGetVersion = True

    End If

End Function

' /*
'   snGetGlobalMsg() -- returns the value of Snarl's global registered message
'
'   Synopsis
'
'       int32 snGetGlobalMsg()
'
'   Inputs
'       None
'
'   Results
'       A 16-bit value (translated to 32-bit) which is the registered Windows
'       message for Snarl.
'
'   Notes
'       Snarl registers SNARL_GLOBAL_MSG during startup which it then uses to
'       communicate with all running applications through a Windows broadcast
'       message.  This function can only fail if for some reason the Windows
'       RegisterWindowMessage() function fails - given this, this function
'       *cannnot* be used to test for the presence of Snarl.
'
' */
Public Function snGetGlobalMsg() As Long

    snGetGlobalMsg = RegisterWindowMessage(SNARL_GLOBAL_MSG)

End Function

Private Function uLOWORD(ByVal dw As Long) As Integer
Dim i As Integer

    CopyMemory i, ByVal VarPtr(dw), 2
    uLOWORD = i

End Function

Private Function uHIWORD(ByVal dw As Long) As Integer
Dim i As Integer

    CopyMemory i, ByVal VarPtr(dw) + 2, 2
    uHIWORD = i

End Function




    ' /*
    '
    '       V37 (R1.6) Additions
    '
    ' */


' /*
'   snRegisterAlert() -- registers a specific application notification  (V37)
'
'   Inputs
'       AppName - Application name, same as that used in snRegisterConfig() or snRegisterConfig2()
'       Class - Alert class name
'
'   Results
'       Returns M_OK if the alert registered okay, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_NOT_FOUND - Application not found in Snarl's roster
'           M_ALREADY_EXISTS - Alert is already registered
'
' */

Public Function snRegisterAlert(ByVal AppName As String, ByVal Class As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_REGISTER_ALERT
        .Title = uToUTF8(AppName)
        .Text = uToUTF8(Class)

    End With

    snRegisterAlert = uSend(pss)

End Function

' /*
'   snRegisterConfig2() -- registers a configuration handler with Snarl  (V37)
'
'   Inputs
'       hWnd - Application name, same as that used in snRegisterConfig() or snRegisterConfig2()
'       AppName - Name of application to register
'       ReplyMsg - Message Snarl will send to hWnd to notify it of something
'       Icon - Path to PNG icon to use
'
'   Results
'       Returns M_OK if the handler registered okay, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_ALREADY_EXISTS - Application is already registered
'           M_ABORTED - Internal problem registering the handler
'
' */

Public Function snRegisterConfig2(ByVal hWnd As Long, ByVal AppName As String, ByVal ReplyMsg As Long, ByVal Icon As String, Optional ByVal LargeIcon As String) As Long
Dim pss As SNARLSTRUCT

    m_hwndFrom = hWnd

    With pss
        .Cmd = SNARL_REGISTER_CONFIG_WINDOW_2
        .LngData2 = hWnd
        .Id = ReplyMsg
        .Title = uToUTF8(AppName)
        .Icon = uToUTF8(Icon)
        ' /* added for R2.2 (V39.50) although R2.1 does support this as well */
        .Text = uToUTF8(LargeIcon)

    End With

    snRegisterConfig2 = uSend(pss)

End Function

' /*
'   snGetIconsPath() -- returns a path to Snarl's default icon location  (V37)
'
'   Synopsis
'
'       str snGetIconsPath()
'
'   Inputs
'       None
'
'   Results
'       A fully-qualified path to Snarl's default icon location
'
'   Notes
'       The easiest way to create this function is as below; simply return the
'       result of snGetAppPath() and tag "etc\icons\" to the end of it.
'
'       Starting with R2.0 (V38) Snarl now makes better use of per-user settings
'       by storing configuration data in %APPDATA%.  Consequently the use of this
'       function is now very limited.
'
' */

Public Function snGetIconsPath() As String

    snGetIconsPath = snGetAppPath() & "etc\icons\"

End Function

' /*
'   snGetAppPath() -- returns a path to Snarl's installed location  (V37)
'
'   Synopsis
'
'       str snGetAppPath()
'
'   Inputs
'       None
'
'   Results
'       A fully-qualified path to the location of the *running* instance of Snarl
'       or an empty string if an error occurred (mostly likely Snarl isn't running)
'
'   Notes
'       Snarl creates a static control within its dispatcher window, the label of
'       which is set to the path the executable is run from.  This function simply
'       finds the dispatcher window, then finds the static control and retrieves
'       the control's title.
'
'       Starting with R2.0 (V38) Snarl now makes better use of per-user settings
'       by storing configuration data in %APPDATA%.  Consequently the use of this
'       function is now very limited.
'
' */
Public Function snGetAppPath() As String
Dim hWnd As Long
Dim hWndPath As Long
Dim sz As String
Dim i As Long

    hWnd = snGetSnarlWindow()
    If hWnd <> 0 Then
        hWndPath = FindWindowEx(hWnd, 0, "static", vbNullString)
        If hWndPath <> 0 Then
            sz = String$(1024, 0)
            i = GetWindowText(hWndPath, sz, Len(sz))
            If i > 0 Then _
                snGetAppPath = Left$(sz, i)

        End If
    End If

End Function



' /*
'   snShowMessageEx() -- displays a Snarl notification using registered class  (V37)
'
'   Inputs
'       Class - Notification class, same as that specified in snRegisterAlert()
'       Title - Text to display in title
'       Text - Text to display in body
'       Timeout - Number of seconds to display notification or zero for indefinite (sticky)
'       IconPath - Path to PNG icon to use
'       hWndReply - Handle of window for Snarl to send replies to
'       uReplyMsg - Message for Snarl to send to hWndReply
'       SoundFile - Path to WAV file to play
'
'   Results
'       Returns an M_RESULT indicating success or failure
'
' */

Public Function snShowMessageEx(ByVal Class As String, ByVal Title As String, ByVal Text As String, Optional ByVal Timeout As Long, Optional ByVal IconPath As String, Optional ByVal hWndReply As Long, Optional ByVal uReplyMsg As Long, Optional ByVal SoundFile As String) As Long
Dim pss As SNARLSTRUCTEX

    With pss
        .Cmd = SNARL_EX_SHOW
        .Title = uToUTF8(Title)
        .Text = uToUTF8(Text)
        .Icon = uToUTF8(IconPath)
        .Timeout = Timeout
        .LngData2 = hWndReply
        .Id = uReplyMsg
        .Extra = uToUTF8(SoundFile)
        .Class = uToUTF8(Class)

    End With

    snShowMessageEx = uSendEx(pss)

End Function

Private Function uSendEx(pss As SNARLSTRUCTEX) As Long
Dim hWnd As Long
Dim pcds As COPYDATASTRUCT
Dim dw As Long

    hWnd = snGetSnarlWindow()
    If IsWindow(hWnd) <> 0 Then
        pcds.dwData = 2
        pcds.cbData = LenB(pss)
        pcds.lpData = VarPtr(pss)
        If SendMessageTimeout(hWnd, WM_COPYDATA, m_hwndFrom, pcds, SMTO_ABORTIFHUNG, 500, dw) > 0 Then
            ' /* worked! */
            uSendEx = dw

        Else
            ' /* timed-out or failed */
            #If USE_LEMON Then
'            g_Debug "uSendEx(): failed (" & g_ApiError() & ")", LEMON_LEVEL_WARNING
            #End If
            uSendEx = &H8000000A        '// M_TIMED_OUT

        End If

    Else
        #If USE_LEMON Then
        g_Debug "uSendEx(): Snarl window not found", LEMON_LEVEL_WARNING
        #End If
        uSendEx = &H80000008            '// M_FAILED

    End If

End Function

' /*
'   snGetSnarlWindow() -- returns a handle to the Snarl Dispatcher window  (V37)
'
'   Synopsis
'
'       int32 snGetSnarlWindow()
'
'   Inputs
'       None
'
'   Results
'       Returns handle to Snarl Dispatcher window, or zero if it's not found
'
'   Notes
'       This is now the preferred way to test if Snarl is actually running
'
' */
Public Function snGetSnarlWindow() As Long

    snGetSnarlWindow = FindWindow("w>Snarl", "Snarl")
    If snGetSnarlWindow = 0 Then _
        snGetSnarlWindow = FindWindow(vbNullString, "Snarl")

End Function

' /*
'   snGetVersionEx() -- returns the Snarl system version number  (V37)
'
'   Synopsis
'
'   int32 snGetVersionEx()
'
'   Inputs
'       None
'
'   Results
'       Returns Snarl system version number or one of the following:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           0 or M_NOT_IMPLEMENTED - Pre-V37 version of Snarl
'
' */
Public Function snGetVersionEx() As Long
Dim pss As SNARLSTRUCT

    pss.Cmd = SNARL_GET_VERSION_EX
    snGetVersionEx = uSend(pss)

End Function

' /*
'   snSetTimeout() -- changes the timeout of an active notification  (V37)
'
'   Inputs
'       Id - Notification identifier, returned after a successful snShowMessage() or snShowMessageEx()
'       Timeout - Updated timeout in seconds, zero means display indefinately
'
'   Results
'       M_OK - Succeeded
'       M_FAILED - Snarl not running
'       M_TIMED_OUT - Message sending timed out
'       M_NOT_FOUND - Notification wasn't found
'
'   Notes
'       Timeout cannot be less than zero or greater than 65535 (18.2 hours)
'
' */
Public Function snSetTimeout(ByVal Id As Long, ByVal Timeout As Long) As Long
Dim pss As SNARLSTRUCT

    pss.Cmd = SNARL_SET_TIMEOUT
    pss.Id = Id
    pss.LngData2 = Timeout
    snSetTimeout = CBool(uSend(pss))

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

    ' /* THIS DOES NOT NEED TO BE RECREATED - IT'S PURELY FOR INTERNAL SNARL TESTING */
Public Sub bsnSet(ByVal l As Long)
    m_hwndFrom = l
End Sub

' /*
'   snSetAsSnarlApp() -- identifies an application as a Snarl App.  (V39)
'
'   Inputs
'       hWndOwner - the window to be used when registering
'       Flags - features this app supports
'
'   Results
'       No return value.
'
' */

Public Sub snSetAsSnarlApp(ByVal hWndOwner As Long, Optional ByVal Flags As SNARL_APP_FLAGS = SNARL_APP_HAS_ABOUT Or SNARL_APP_HAS_PREFS)

    If IsWindow(hWndOwner) <> 0 Then
        SetProp hWndOwner, "snarl_app", 1
        SetProp hWndOwner, "snarl_app_flags", Flags

    End If

End Sub

' /*
'   snGetAppMsg() -- Returns the global Snarl Application message  (V39)
'
'   Inputs
'       None
'
'   Results
'       Snarl Application registered message.
'
' */

Public Function snGetAppMsg() As Long

    snGetAppMsg = RegisterWindowMessage(SNARL_APP_MSG)

End Function


' /*
'   snRegisterApp() -- registers an application with Snarl  (V39)
'
'   Inputs
'       Application - Name of application to register
'       ReplyMsg - Message Snarl will send to hWnd to notify it of something
'       SmallIcon - Path to PNG icon to use in Snarl's preferences
'       LargeIcon - Path to PNG icon to use in Registered/Unregistered notifications
'
'   Results
'       Returns M_OK if the handler registered okay, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_ALREADY_EXISTS - Application is already registered
'           M_ABORTED - Internal problem registering the handler
'
' */

Public Function snRegisterApp(ByVal Application As String, ByVal SmallIcon As String, ByVal LargeIcon As String, Optional ByVal hWnd As Long, Optional ByVal ReplyMsg As Long) As Long
Dim pss As SNARLSTRUCT

    m_hwndFrom = hWnd

    With pss
        .Cmd = SNARL_REGISTER_APP
        .Title = uToUTF8(Application)
        .Icon = uToUTF8(SmallIcon)
        .Text = uToUTF8(LargeIcon)
        .LngData2 = hWnd
        .Id = ReplyMsg
        .Timeout = GetCurrentProcessId()

    End With

    snRegisterApp = uSend(pss)

End Function


' /*
'   snUnregisterApp() -- unregisters an application with Snarl  (V39)
'
'   PRIVATE FUNCTION: due for documentation in V39.  For now should only be used
'   under direct guidance from application developers.
'
'   Inputs
'       None
'
'   Results
'       Returns M_OK if the handler registered okay, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_ALREADY_EXISTS - Application is already registered
'           M_ABORTED - Internal problem registering the handler
'
' */

Public Function snUnregisterApp() As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_UNREGISTER_APP
        .LngData2 = GetCurrentProcessId()

    End With

    snUnregisterApp = uSend(pss)
    m_hwndFrom = 0

End Function


' /*
'   snShowNotification() -- displays a Snarl notification using registered class  (V39)
'
'   Inputs
'       Application - name of application
'       Class - Class, same as that specified in snRegisterAlert()
'       Title - Text to display in title
'       Text - Text to display in body
'       Timeout - Number of seconds to display notification or zero for indefinite (sticky)
'       IconPath - Path to PNG icon to use
'       hWndReply - Handle of window for Snarl to send replies to
'       uReplyMsg - Message for Snarl to send to hWndReply
'       SoundFile - Path to WAV file to play
'
'   Results
'       Returns handle to Snarl Dispatcher window, or zero if it's not found
'
' */

Public Function snShowNotification(ByVal Class As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Timeout As Long, Optional ByVal Icon As String, Optional ByVal hWndReply As Long, Optional ByVal uReplyMsg As Long, Optional ByVal Sound As String) As Long
Dim pss As SNARLSTRUCTEX

    With pss
        .Cmd = SNARL_SHOW_NOTIFICATION
        .Title = uToUTF8(Title)
        .Text = uToUTF8(Text)
        .Icon = uToUTF8(Icon)
        .Timeout = Timeout
        .LngData2 = hWndReply
        .Id = uReplyMsg
        .Extra = uToUTF8(Sound)
        .Class = uToUTF8(Class)
        .Reserved1 = GetCurrentProcessId()

    End With

    snShowNotification = uSendEx(pss)

End Function

Public Function snChangeAttribute(ByVal Id As Long, ByVal Attr As SNARL_ATTRIBUTES, ByVal Value As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_CHANGE_ATTR
        .Id = Id
        .LngData2 = Attr
        .Text = uToUTF8(Value)

    End With

    snChangeAttribute = uSend(pss)

End Function

' /*
'   snSetClassDefault() -- sets the default value for an alert class  (V39)
'
'   PRIVATE FUNCTION: due for documentation in V39.  For now should only be used
'   under direct guidance from the application developers.
'
'   Inputs
'       Application - Application name, same as that used in snRegisterConfig(), snRegisterConfig2() or snRegisterApp()
'       Class - Class name
'       Attr - Class element to change
'       Value - New value
'
'   Results
'       Returns M_OK if the alert registered okay, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_NOT_FOUND - Application or Alert Class not found in Snarl's roster
'           M_INVALID_ARGS - Invalid argument specified
'
' */

Public Function snSetClassDefault(ByVal Class As String, ByVal Attr As SNARL_ATTRIBUTES, ByVal Value As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_SET_CLASS_DEFAULT
        .Text = uToUTF8(Class)
        .LngData2 = Attr
        .Icon = uToUTF8(Value)
        .Timeout = GetCurrentProcessId()

    End With

    snSetClassDefault = uSend(pss)

End Function

' /*
'   snGetRevision() -- gets the current Snarl revision (build) number  (V39)
'
'   Inputs
'       None
'
'   Results
'       Returns the build version number, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'
' */

Public Function snGetRevision() As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_GET_REVISION
        .LngData2 = &HFFFE&             ' // COPWAIT ;)

    End With

    snGetRevision = uSend(pss)

End Function

Public Function snAddClass(ByVal Class As String, Optional ByVal Description As String, Optional ByVal Flags As SNARL_CLASS_FLAGS, Optional ByVal DefaultTitle As String, Optional ByVal DefaultIcon As String, Optional ByVal DefaultTimeout As Long) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_ADD_CLASS
        .Text = uToUTF8(Class)
        .Title = uToUTF8(Description)
        .LngData2 = Flags
        .Timeout = GetCurrentProcessId()

    End With

    snAddClass = uSend(pss)

    If snAddClass = 0 Then
        ' /* succeeded */
        snSetClassDefault Class, SNARL_ATTRIBUTE_TITLE, DefaultTitle
        snSetClassDefault Class, SNARL_ATTRIBUTE_ICON, DefaultIcon
        If DefaultTimeout > 0 Then _
            snSetClassDefault Class, SNARL_ATTRIBUTE_TIMEOUT, CStr(DefaultTimeout)

    End If

End Function

