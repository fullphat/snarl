Attribute VB_Name = "mSnarlPrivate_i"
Option Explicit

    ' /*
    '
    '   mSnarlPrivate_i.bas -- Private Snarl functions
    '
    '   Copyright (C) 2004-2008 full phat products
    '             All Rights Reserved.
    '
    '   NOT FOR GENERAL USE.  Refer to documentation for more information.  If in any
    '   doubt please consult development team for guidance.
    '
    ' */
    
    
    ' /*
    '
    '   SNARLSTRUCT
    '   This is an internal structure used to pass information between Snarl and
    '   registered applications - do not attempt to craft your own messages
    '   using this structure; use the standard sn... api methods instead.
    '
    ' */


Public Type SNARLSTRUCT
    Cmd As Long       ' // what to do...
    Id As Long                  ' // snarl message id (returned by snShowMessage())
    Timeout As Long             ' // timeout in seconds (0=sticky)
    LngData2 As Long            ' // reserved
    Title As String * 512
    Text As String * 512        ' // VB defines these as wide so they're actually 1024 bytes
    Icon As String * 512

End Type

Public Type SNARLSTRUCTEX
    Cmd As Long       ' // what to do...
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


    
    

    ' /* ================= Constants and Enums ================= */

    ' /* WM_MANAGE_SNARL - reserved for future use */
Public Const WM_MANAGE_SNARL = &H400 + 238  ' // note hardcoded WM_USER value!

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


    ' /* snPrivateUpdateMessage() flags */
Public Enum E_SNARL_UPDATE_FLAGS
    E_SNARL_CREATE_IF_NEEDED = 1        '// snarl will create a new notification if requested one has gone

End Enum


Public Enum PRIVATE_SNARL_COMMANDS
    SNARL_PRIV_LAST_PUBLIC = 17         '// (SNARL_ADD_CLASS)
    SNARL_LOAD_EXTENSION                '// private (V38.107)
    SNARL_UNLOAD_EXTENSION              '// private (V38.107)
    SNARL_COUNT_NOTIFICATIONS           '// private (V39.12)
    SNARL_PREVIEW_SCHEME                '// private (V39.66)
    SNARL_LOAD_STYLE_ENGINE             '// private (V39.77)
    SNARL_UNLOAD_STYLE_ENGINE           '// private (V39.77)

End Enum

'Public Const SNARL_GET_REVISION = 9     '// private (V38.128) NOTE: Same as unused SNARL_REVOKE_ALERT

Public Enum SNRL_NOTIFICATION_FLAGS
    SNRL_NOTIFICATION_REMOTE = &H80000000
    SNRL_NOTIFICATION_SECURE = &H40000000

End Enum

    ' /* internal declares */

Private Const WM_COPYDATA = &H4A
Private Const SNARL_GLOBAL_MSG = "SnarlGlobalEvent"

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

Private Const CP_UTF8 = 65001
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long


                        ' /* functions */

    ' /* ----------------------- Private ----------------------- */







' /*
'   snPrivateUnloadExtension() -- asks Snarl to unload a specific extension  (V39)
'
'   PRIVATE FUNCTION: due for documentation in V39.  For now should only be used
'   under direct guidance from application developers.
'
'   Inputs
'       Extension - Name of extension to unload
'
'   Results
'       Returns M_OK if the handler registered okay, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_ALREADY_EXISTS - Application is already registered
'           M_ABORTED - Internal problem registering the handler
'
' */

Public Function snPrivateUnloadExtension(ByVal Extension As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_UNLOAD_EXTENSION
        .Text = uToUTF8(Extension)

    End With

    snPrivateUnloadExtension = uSend(pss)

End Function


' /*
'   snPrivateLoadExtension() -- asks Snarl to load a new extension  (V39)
'
'   PRIVATE FUNCTION: due for documentation in V39.  For now should only be used
'   under direct guidance from application developers.
'
'   Inputs
'       Extension - Name and path to extension to load
'
'   Results
'       Returns M_OK if the handler registered okay, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_ALREADY_EXISTS - Application is already registered
'           M_ABORTED - Internal problem registering the handler
'
' */

Public Function snPrivateLoadExtension(ByVal Extension As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_LOAD_EXTENSION
        .Text = uToUTF8(Extension)

    End With

    snPrivateLoadExtension = uSend(pss, 1000)       ' // give it a bit longer to complete

End Function

' /*
'   snPrivateCountNotifications() -- asks Snarl to load a new extension  (V40)
'
'   PRIVATE FUNCTION: due for documentation in V40.  For now should only be used
'   under direct guidance from application developers.
'
'   Inputs
'       ByClass - count notification for a particular class, or for the whole application
'       Class - only required if ByClass is TRUE.  Name of class to count
'
'   Results
'       Returns the number of notifications for the app or class or one of the following:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_ABORTED - Internal problem
'
' */

Public Function snPrivateCountNotifications(ByVal ByClass As Boolean, ByVal Class As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_COUNT_NOTIFICATIONS
        .LngData2 = IIf(ByClass, 1, 0)
        .Text = uToUTF8(Class)
        .Timeout = GetCurrentProcessId()

    End With

    snPrivateCountNotifications = uSend(pss)

End Function


    ' /* ==================== End Of Private ==================== */






'Public Function snPrivateAddClass(ByVal Class As String, Optional ByVal Description As String, Optional ByVal Flags As SNARL_CLASS_FLAGS) As Long ', Optional ByVal DefaultTitle As String, Optional ByVal DefaultIcon As String, Optional ByVal DefaultTimeout As Long) As Long
'Dim pss As SNARLSTRUCT
'
'    With pss
'        .Cmd = SNARL_ADD_CLASS
'        .Text = uToUTF8(Class)
'        .Title = uToUTF8(Description)
'        .LngData2 = Flags
'        .Timeout = GetCurrentProcessId()
'
'    End With
'
'    snPrivateAddClass = uSend(pss)
''    If snPrivateAddClass = 0 Then
''        ' /* succeeded */
''        snPrivateSetClassDefault Class, SNARL_ATTRIBUTE_TITLE, DefaultTitle
''        snPrivateSetClassDefault Class, SNARL_ATTRIBUTE_ICON, DefaultIcon
''        If DefaultTimeout > 0 Then _
''            snPrivateSetClassDefault Class, SNARL_ATTRIBUTE_TIMEOUT, CStr(DefaultTimeout)
''
''    End If
'
'End Function

'Public Function snPrivateUpdateMessage(ByVal Id As Long, ByVal Title As String, ByVal Text As String, Optional ByVal IconPath As String, Optional ByVal Flags As E_SNARL_UPDATE_FLAGS) As Long
'Dim pss As SNARLSTRUCT
'
'    ' /* this is currently the same as snUpdateMessage() except that it allows for an extra parameter - Flags - to be
'    '    specified.  At the moment only a single flag - E_SNARL_CREATE_IF_NEEDED - if defined and this isn't actually
'    '    implemented as yet.  Note that the command is still the original SNARL_UPDATE; we just make use of the
'    '    previously reserved "LngData2" value. */
'
'    With pss
'        .Cmd = SNARL_UPDATE
'        .Id = Id
'        .Title = uToUTF8(Title)
'        .Text = uToUTF8(Text)
'        ' /* 1.6 Beta 4 */
'        .Icon = uToUTF8(IconPath)
'        ' /* V39 */
'        .LngData2 = Flags
'
'    End With
'
'    snPrivateUpdateMessage = uSend(pss)
'
'End Function

' /*
'   snPrivateLoadStyleEngine() -- asks Snarl to reload an EXISTING style engine  (V39)
'
'   PRIVATE FUNCTION: should only be used under direct guidance from the Snarl dev team.
'
'   Inputs
'       Engine - Object name of style engine to load (i.e. <engine>.styleengine)
'
'   Results
'       Returns M_OK on success, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_NOT_FOUND - Error unloading the engine
'
' */

Public Function snPrivateLoadStyleEngine(ByVal Engine As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_LOAD_STYLE_ENGINE
        .Text = uToUTF8(Engine)

    End With

    snPrivateLoadStyleEngine = uSend(pss)

End Function

' /*
'   snPrivateUnloadStyleEngine() -- asks Snarl to unload an existing style engine  (V39)
'
'   PRIVATE FUNCTION: should only be used under direct guidance from the Snarl dev team.
'
'   Inputs
'       Engine - Object name of style engine to unload (i.e. <engine>.styleengine)
'
'   Results
'       Returns M_OK on success, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_NOT_FOUND - Error unloading the engine
'
' */

Public Function snPrivateUnloadStyleEngine(ByVal Engine As String) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_UNLOAD_STYLE_ENGINE
        .Text = uToUTF8(Engine)

    End With

    snPrivateUnloadStyleEngine = uSend(pss)

End Function








    ' /* =================== Local Functions =================== */

Private Function uSend(pss As SNARLSTRUCT, Optional ByVal Timeout As Long = 500) As Long
Dim hWnd As Long
Dim pcds As COPYDATASTRUCT
Dim dw As Long

    hWnd = FindWindow("w>Snarl", "Snarl")
    If IsWindow(hWnd) <> 0 Then
        pcds.dwData = 2                 '// SNARLSTRUCT
        pcds.cbData = LenB(pss)
        pcds.lpData = VarPtr(pss)
        If SendMessageTimeout(hWnd, WM_COPYDATA, IIf(m_hwndFrom = 0, GetCurrentProcessId(), m_hwndFrom), pcds, SMTO_ABORTIFHUNG, Timeout, dw) > 0 Then
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

Private Function uSendEx(pss As SNARLSTRUCTEX) As Long
Dim hWnd As Long
Dim pcds As COPYDATASTRUCT
Dim dw As Long

    hWnd = FindWindow(vbNullString, "Snarl")
    If IsWindow(hWnd) <> 0 Then
        pcds.dwData = 2
        pcds.cbData = LenB(pss)
        pcds.lpData = VarPtr(pss)
        If SendMessageTimeout(hWnd, WM_COPYDATA, m_hwndFrom, pcds, SMTO_ABORTIFHUNG, 500, dw) > 0 Then
            ' /* worked! */
            uSendEx = dw

        Else
            ' /* timed-out or failed */
            uSendEx = &H8000000A        '// M_TIMED_OUT

        End If

    Else
        uSendEx = &H80000008            '// M_FAILED

    End If

End Function

Private Function uToUTF8(ByVal str As String) As String
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

Public Function snManageSnarl(ByVal Command As E_MANAGE_SNARL, Optional ByVal Data As Long) As Long
Dim hWnd As Long

    snManageSnarl = &H80000008            '// M_FAILED
    hWnd = FindWindow(vbNullString, "Snarl")
    If IsWindow(hWnd) = 0 Then _
        Exit Function

Dim dw As Long

    snManageSnarl = &H8000000A        '// M_TIMED_OUT
    If SendMessageTimeout(hWnd, WM_MANAGE_SNARL, Command, ByVal Data, SMTO_ABORTIFHUNG, 500, dw) > 0 Then _
        snManageSnarl = dw

End Function

Public Function snPrivateOpenStyleRoster() As Long

    snPrivateOpenStyleRoster = snManageSnarl(E_MANAGE_STYLE_ROSTER, 1)

End Function

Public Function snPrivateCloseStyleRoster() As Long

    snPrivateCloseStyleRoster = snManageSnarl(E_MANAGE_STYLE_ROSTER, 2)

End Function

Public Function snCommandStr(ByVal Command As Long) As String

'    Select Case Command
'    Case SNARL_SHOW
'        snCommandStr = "SNARL_SHOW"
'
'    Case SNARL_HIDE
'        snCommandStr = "SNARL_HIDE"
'
'    Case SNARL_UPDATE
'        snCommandStr = "SNARL_UPDATE"
'
'    Case SNARL_IS_VISIBLE
'        snCommandStr = "SNARL_IS_VISIBLE"
'
'    Case SNARL_GET_VERSION
'        snCommandStr = "SNARL_GET_VERSION"
'
'    Case SNARL_REGISTER_CONFIG_WINDOW
'        snCommandStr = "SNARL_REGISTER_CONFIG_WINDOW"
'
'    Case SNARL_REVOKE_CONFIG_WINDOW
'        snCommandStr = "SNARL_REVOKE_CONFIG_WINDOW"
'
'    ' /* R1.6 onwards */
'    Case SNARL_REGISTER_ALERT
'        snCommandStr = "SNARL_REGISTER_ALERT"
'
'    Case SNARL_REVOKE_ALERT                  '// for future use
'        snCommandStr = "SNARL_REVOKE_ALERT"
'
'    Case SNARL_REGISTER_CONFIG_WINDOW_2
'        snCommandStr = "SNARL_REGISTER_CONFIG_WINDOW_2"
'
'    Case Else
'        snCommandStr = "unknown: " & CStr(Command)
'
'    End Select

End Function

' /*
'   snPrivateLoadExtension() -- asks Snarl to load a new extension  (V39)
'
'   PRIVATE FUNCTION: due for documentation in V39.  For now should only be used
'   under direct guidance from application developers.
'
'   Inputs
'       Extension - Name and path to extension to load
'
'   Results
'       Returns M_OK if the handler registered okay, or one of the following if it didn't:
'           M_FAILED - Snarl not running
'           M_TIMED_OUT - Message sending timed out
'           M_ALREADY_EXISTS - Application is already registered
'           M_ABORTED - Internal problem registering the handler
'
' */

Public Function snPrivatePreviewScheme(ByVal StyleName As String, ByVal SchemeName As String, Optional ByVal Priority As Boolean, Optional ByVal Timeout As Long, Optional ByVal Percent As Integer) As Long
Dim pss As SNARLSTRUCT

    With pss
        .Cmd = SNARL_PREVIEW_SCHEME
        .Title = uToUTF8(StyleName)
        .Text = uToUTF8(SchemeName)
        .Timeout = Timeout
        .LngData2 = IIf(Priority, 1, 0)
        .Id = Percent

    End With

    snPrivatePreviewScheme = uSend(pss)

End Function
