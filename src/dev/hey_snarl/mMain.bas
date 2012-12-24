Attribute VB_Name = "mMain"
Option Explicit

' Constants that will be used in the API functions
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&

' Declare the needed API functions
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal bsName As String, ByVal buff As String, ByVal ch As Long) As Long

'Private Const WM_USER = &H400
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public Sub Main()
'Dim sz() As String
Dim hr As Long

    If Command$ = "" Then
        uHint SNARL_ERROR_FAILED

    ElseIf Command$ = "?" Then
        uHint SNARL_ERROR_FAILED

    Else
        uEnd (snDoRequest(g_RemoveQuotes(Command$)))

'        sz() = Split(Command$, " ")
'        uWrite UBound(sz)
'
'        If UBound(sz) <> 0 Then
'            ' /* not the right number of args */
'            uHint SNARL_ERROR_ARG_MISSING
'
'        Else
'            uEnd (snDoRequest(sz(0)))
'
'        End If

    End If

End Sub

Private Sub uHint(ByVal ExitCode As SNARL_STATUS_CODE)

    uWrite "HeySnarl " & CStr(App.Major) & "." & CStr(App.Minor) & " Build " & CStr(App.Revision) & vbCrLf & App.LegalCopyright & vbCrLf
    uWrite "Usage: heysnarl <request>"
    uWrite "<request> should be enclosed in quotes if it includes spaces"
    ExitProcess ExitCode

End Sub

'======================
' Send output to STDOUT
'======================
'
Private Sub uWrite(ByVal s As String)
Dim llResult As Long

    s = s & vbCrLf
    WriteFile GetStdHandle(STD_OUTPUT_HANDLE), s, Len(s), llResult, ByVal 0&

End Sub

Private Sub uEnd(ByVal ExitCode As SNARL_STATUS_CODE)

    If ExitCode > 0 Then
        uWrite "Ok: " & CStr(ExitCode)

    ElseIf ExitCode = 0 Then
        uWrite "Ok"

    Else
        uWrite "Failed: " & CStr(Abs(ExitCode)) & " (" & uError(Abs(ExitCode)) & ")"

    End If

    ExitProcess ExitCode

End Sub

''============================
'' Get the CGI data from STDIN
''============================
'' Data is collected as a single string. We will read it 1024 bytes at a time.
''
'Sub GetCGIpostData()
'
'   ' Read the standard input handle
'   llStdIn = GetStdHandle(STD_INPUT_HANDLE)
'   ' Get POSTed CGI data from STDIN
'   Do
'      lsBuff = String(1024, 0)    ' Create a buffer big enough to hold the 1024 bytes
'      llBytesRead = 1024          ' Tell it we want at least 1024 bytes
'      If ReadFile(llStdIn, ByVal lsBuff, 1024, llBytesRead, ByVal 0&) Then
'         ' Read the data
'         ' Add the data to our string
'         postData = postData & Left(lsBuff, llBytesRead)
'         If llBytesRead < 1024 Then Exit Do
'      Else
'         Exit Do
'      End If
'   Loop
'
'End Sub


Private Function uError(ByVal Error As Long) As String

    Select Case Error
    Case SNARL_SUCCESS:                     uError = "SUCCESS"
    Case SNARL_CALLBACK_R_CLICK:            uError = "CALLBACK_R_CLICK"
    Case SNARL_CALLBACK_TIMED_OUT:          uError = "CALLBACK_TIMED_OUT"
    Case SNARL_CALLBACK_INVOKED:            uError = "CALLBACK_INVOKED"
    Case SNARL_CALLBACK_MENU_SELECTED:      uError = "CALLBACK_MENU_SELECTED"
    Case SNARL_CALLBACK_M_CLICK:            uError = "CALLBACK_M_CLICK"
    Case SNARL_CALLBACK_CLOSED:             uError = "CALLBACK_CLOSED"
    Case SNARL_ERROR_FAILED:                uError = "ERROR_FAILED"
    Case SNARL_ERROR_UNKNOWN_COMMAND:       uError = "ERROR_UNKNOWN_COMMAND"
    Case SNARL_ERROR_TIMED_OUT:             uError = "ERROR_TIMED_OUT"
    Case SNARL_ERROR_BAD_SOCKET:            uError = "ERROR_BAD_SOCKET"
    Case SNARL_ERROR_BAD_PACKET:            uError = "ERROR_BAD_PACKET"
    Case SNARL_ERROR_INVALID_ARG:           uError = "ERROR_INVALID_ARG"
    Case SNARL_ERROR_ARG_MISSING:           uError = "ERROR_ARG_MISSING"
    Case SNARL_ERROR_SYSTEM:                uError = "ERROR_SYSTEM"
    Case SNARL_ERROR_ACCESS_DENIED:         uError = "ERROR_ACCESS_DENIED"
    Case SNARL_ERROR_UNSUPPORTED_VERSION:   uError = "ERROR_UNSUPPORTED_VERSION"
    Case SNARL_ERROR_NO_ACTIONS_PROVIDED:   uError = "ERROR_NO_ACTIONS_PROVIDED"
    Case SNARL_ERROR_UNSUPPORTED_ENCRYPTION:    uError = "ERROR_UNSUPPORTED_ENCRYPTION"
    Case SNARL_ERROR_UNSUPPORTED_HASHING:   uError = "ERROR_UNSUPPORTED_HASHING"
    Case SNARL_ERROR_NOT_RUNNING:           uError = "ERROR_NOT_RUNNING"
    Case SNARL_ERROR_NOT_REGISTERED:        uError = "ERROR_NOT_REGISTERED"
    Case SNARL_ERROR_ALREADY_REGISTERED:    uError = "ERROR_ALREADY_REGISTERED"
    Case SNARL_ERROR_CLASS_ALREADY_EXISTS:  uError = "ERROR_CLASS_ALREADY_EXISTS"
    Case SNARL_ERROR_CLASS_BLOCKED:         uError = "ERROR_CLASS_BLOCKED"
    Case SNARL_ERROR_CLASS_NOT_FOUND:       uError = "ERROR_CLASS_NOT_FOUND"
    Case SNARL_ERROR_NOTIFICATION_NOT_FOUND:    uError = "ERROR_NOTIFICATION_NOT_FOUND"
    Case SNARL_ERROR_FLOODING:              uError = "ERROR_FLOODING"
    Case SNARL_ERROR_DO_NOT_DISTURB:        uError = "ERROR_DO_NOT_DISTURB"
    Case SNARL_ERROR_COULD_NOT_DISPLAY:     uError = "ERROR_COULD_NOT_DISPLAY"
    Case SNARL_ERROR_AUTH_FAILURE:          uError = "ERROR_AUTH_FAILURE"
    Case SNARL_ERROR_DISCARDED:             uError = "ERROR_DISCARDED"
    Case SNARL_ERROR_NOT_SUBSCRIBED:        uError = "ERROR_NOT_SUBSCRIBED"
    Case SNARL_WAS_MERGED:                  uError = "WAS_MERGED"

    Case SNARL_NOTIFY_GONE:                 uError = "NOTIFY_GONE"
    Case 302:                               uError = "NOTIFY_CLICK"
    Case SNARL_NOTIFY_EXPIRED:              uError = "NOTIFY_EXPIRED"
    Case SNARL_NOTIFY_INVOKED:              uError = "NOTIFY_INVOKED"
    Case SNARL_NOTIFY_MENU:                 uError = "NOTIFY_MENU"
    Case 306:                               uError = "NOTIFY_EX_CLICK"
    Case SNARL_NOTIFY_CLOSED:               uError = "NOTIFY_CLOSED"
    Case SNARL_NOTIFY_ACTION:               uError = "NOTIFY_ACTION"
    Case SNARL_NOTIFY_APP_DO_ABOUT:         uError = "NOTIFY_APP_DO_ABOUT"
    Case SNARL_NOTIFY_APP_DO_PREFS:         uError = "NOTIFY_APP_DO_PREFS"

    Case Else:                              uError = "Undefined error"

    End Select

End Function
