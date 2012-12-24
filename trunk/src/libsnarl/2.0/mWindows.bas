Attribute VB_Name = "mWindows"
Option Explicit

Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const CS_DBLCLKS = &H8
Private Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hBrBackground As Long
    lpszMenuName As String
    lpszClassName As String

End Type

Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As Any, ByVal hInstance As Long) As Long

Private Const WM_NCCREATE = &H81

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal g_hMenu As Long, ByVal hInstance As Long, lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long


'Private Const WM_SETICON = &H80
'Private Const ICON_SMALL = 0
'Private Const ICON_BIG = 1

Dim mClass() As String
Dim mClasses As Long

Private Type T_WINDOW
    hWnd As Long
    Handler As IWndProc

End Type

Dim mWindows As Long
Dim mWindow() As T_WINDOW

Public Function EZRegisterClass(ByVal Name As String, Optional ByVal Styles As Long = CS_DBLCLKS) As Boolean
Dim wc As WNDCLASS

    If (Name = "") Then _
        Exit Function

    With wc
        .hInstance = App.hInstance
        .lpfnwndproc = ADDROF(AddressOf uEZClassWndProc)
        .lpszClassName = Name
        .hCursor = LoadCursor(0, IDC_ARROW)
        .Style = Styles

    End With

    If RegisterClass(wc) <> 0 Then
        g_Debug "EZRegisterClass(): class '" & Name & "' registered okay"

        mClasses = mClasses + 1
        ReDim Preserve mClass(mClasses)
        mClass(mClasses) = Name

        EZRegisterClass = True

    Else
        g_Debug "EZRegisterClass(): class '" & Name & "' not registered (" & Err.LastDllError & ")", LEMON_LEVEL_CRITICAL

    End If

End Function

Public Function EZUnregisterClass(ByVal Name As String) As Boolean
Dim i As Long
Dim j As Long

    i = uFindClass(Name)
    If i Then
        g_Debug "EZUnregisterClass(): '" & Name & "' found"

        j = UnregisterClass(Name, App.hInstance)
        g_Debug "_UnregisterClass() returned " & CStr(j)
        If j = 0 Then _
            Exit Function                           ' // unregisterclass() failed

        If i < mClasses Then
            For j = i To (mClasses - 1)
                mClass(j) = mClass(j + 1)

            Next j

        End If

        mClasses = mClasses - 1
        ReDim Preserve mClass(mClasses)
        g_Debug "EZUnregisterClass(): '" & Name & "' unregistered okay"
        EZUnregisterClass = True

    Else
        g_Debug "EZUnregisterClass(): '" & Name & "' not found", LEMON_LEVEL_WARNING

    End If

End Function

Private Function uFindClass(ByVal Name As String) As Long
Dim i As Long

    If mClasses Then
        For i = 1 To mClasses
            If mClass(i) = Name Then
                uFindClass = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Function uEZClassWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Debug.Print g_HexStr(uMsg)

    If uMsg = WM_NCCREATE Then _
        mWindow(mWindows).hWnd = hWnd           ' // fix the reference

Static i As Long
Dim r As Long

    i = uFindWindow(hWnd)
    If i = 0 Then
        Debug.Print "uEZClassWndProc(): " & g_HexStr(hWnd) & " was not found"
        uEZClassWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
        Exit Function

    End If

    If (mWindow(i).Handler Is Nothing) Then
        Debug.Print "uEZClassWndProc(): " & g_HexStr(hWnd) & " has no handler"
        uEZClassWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
        Exit Function

    End If

    If mWindow(i).Handler.WndProc(hWnd, uMsg, wParam, lParam, 0, r) Then
        uEZClassWndProc = r

    Else
        uEZClassWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    
    End If

End Function

Public Function EZ4AddWindow(ByVal Class As String, ByRef Handler As IWndProc, Optional ByVal Title As String, Optional ByVal Styles As Long = &H80000000, Optional ByVal ExStyles As Long = &H80, Optional ByVal hWndParent As Long) As Long
Static hr As Long

'    If (Handler Is Nothing) Or (Class = "") Then
'        g_Debug "EZ4AddWindow(): bad args", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End If

    ' /* add the handler _before_ we create the window so we can capture all messages */

    mWindows = mWindows + 1
    ReDim Preserve mWindow(mWindows)
    Set mWindow(mWindows).Handler = Handler

    hr = CreateWindowEx(ExStyles, Class, Title, Styles, 0, 0, 1, 1, hWndParent, 0, App.hInstance, ByVal 0&)
    If hr = 0 Then
        g_Debug "EZ4AddWindow(): couldn't create window [" & uApiError() & "]", LEMON_LEVEL_CRITICAL
        Set mWindow(mWindows).Handler = Nothing
        mWindows = mWindows - 1
        Exit Function

'    Else
'        Debug.Print "OK: " & g_HexStr(hr)

    End If

    mWindow(mWindows).hWnd = hr
    EZ4AddWindow = hr

End Function

Public Function EZ4RemoveWindow(ByVal hWnd As Long) As Boolean
Static i As Long
Static j As Long

    i = uFindWindow(hWnd)
    If i = 0 Then
        g_Debug "EZ4RemoveWindow(): '" & g_HexStr(hWnd) & "' not found", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    j = DestroyWindow(hWnd)
    If j = 0 Then
        g_Debug "EZ4RemoveWindow(): _DestoryWindow() failed (" & Err.LastDllError & ")", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    Set mWindow(i).Handler = Nothing

    If i < mWindows Then
        For j = i To (mWindows - 1)
            LSet mWindow(j) = mWindow(j + 1)

        Next j

    End If

    mWindows = mWindows - 1
    ReDim Preserve mWindow(mWindows)
    g_Debug "EZ4RemoveWindow(): '" & g_HexStr(hWnd) & "' removed ok"
    EZ4RemoveWindow = True

End Function

Private Function uFindWindow(ByVal hWnd As Long) As Long
Static i As Long

    If mWindows Then
        For i = 1 To mWindows
            If mWindow(i).hWnd = hWnd Then
                uFindWindow = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Function uApiError(Optional ByVal lError As Long = -1, Optional ByVal AddErrorCode As Boolean = True) As String
Dim sz  As String
Dim hr   As Long

    If lError = -1 Then _
        lError = Err.LastDllError

    sz = String(256, 0)
    hr = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, lError, 0, sz, Len(sz), ByVal 0)
    If hr > 2 Then
        uApiError = Left$(sz, hr - 2)      ' // trim off CR/LF...
        If AddErrorCode Then _
            uApiError = uApiError & " (" & CStr(lError) & ")"

    End If

End Function

