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

Private Const WM_USER = &H400
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public Sub Main()
Dim sz() As String

    If Command$ = "" Then
        uHint TMINUS_BAD_ARGS

    Else
        sz() = Split(Command$, " ")
        If UBound(sz) <> 1 Then
            ' /* not the right number of args */
            uHint TMINUS_BAD_ARGS

        Else
            Select Case sz(0)
            Case "-t"
                uDoTime sz(1)

            Case "-s"
                If Not g_IsNumeric(sz(1)) Then
                    uWrite "<seconds> must be a numeric value"
                    ExitProcess TMINUS_BAD_ARGS

                Else
                    uDoSeconds g_SafeLong(sz(1))

                End If

            Case Else
                uWrite "unknown argument '" & sz(0) & "'"
                uHint TMINUS_BAD_ARGS

            End Select

        End If

    End If

End Sub

Private Sub uDoTime(ByVal Time As String)
Dim s() As String

    s = Split(Time, ":")
    If UBound(s) <> 1 Then
        uWrite "<time> must be formatted as hh:mm"
        uEnd TMINUS_BAD_ARGS

    End If

Dim h As Long
Dim t As Long

    h = FindWindow(TMINUS_CLASS_NAME, "")
    If h = 0 Then
        uWrite "TMinus is not running"
        uEnd TMINUS_NOT_RUNNING

    Else
        t = MAKELONG(MAKEWORD(Val(s(1)), Val(s(0))), 0)

    End If

    uEnd SendMessage(h, WM_USER + 4, 1, ByVal t)

End Sub

Private Sub uDoSeconds(ByVal Seconds As Long)
Dim h As Long
Dim r As Long

    h = FindWindow(TMINUS_CLASS_NAME, "")
    If h = 0 Then
        uWrite "TMinus is not running"
        uEnd TMINUS_NOT_RUNNING

    Else
        uEnd SendMessage(h, WM_USER + 4, 0, ByVal Seconds)

    End If

End Sub

Private Sub uHint(ByVal ExitCode As TMINUS_STATUS_CODES)

    uWrite "TMinus CLI " & CStr(App.Major) & "." & CStr(App.Minor) & " Build " & CStr(App.Revision) & " " & App.LegalCopyright
    uWrite "Usage: tminus -s <seconds> or tminus -t <time>"
    uWrite "<time> should be formatted as hh:mm and should use the 24 hour clock"
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

Private Sub uEnd(ByVal ExitCode As TMINUS_STATUS_CODES)

    If ExitCode = TMINUS_SUCCESS Then
        uWrite "Ok"

    Else
        uWrite "Failed: " & CStr(ExitCode)

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


