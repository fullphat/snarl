Attribute VB_Name = "mMain"
Option Explicit

Public Const WM_TEST = &H400 + 1
'Public Const WM_CLOSE = &H10

Private Const CLASS_NAME = "w>snarlmail"

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub Main()
Dim hWndExisting As Long
Dim bQuit As Boolean
Dim sz As String

    If g_GetSystemFolder(CSIDL_APPDATA, sz) Then
        sz = g_MakePath(sz) & "full phat\SnarlMail\"

    Else
        sz = g_MakePath(App.Path)

    End If

    l3OpenLog sz & "snarlmail.log"

    bQuit = (InStr(Command$, "-quit") <> 0)
    hWndExisting = FindWindow(CLASS_NAME, CLASS_NAME)

    g_Debug "main: launched: hWndExisting=0x" & g_HexStr(hWndExisting) & " -quit:" & bQuit

    If (hWndExisting) Or (bQuit) Then
        g_Debug "main: existing instance detected (this one will now close)", LEMON_LEVEL_INFO
        If bQuit Then _
            PostMessage hWndExisting, WM_CLOSE, 0, ByVal 0&

        Exit Sub

    End If

    If Not uGotMiscResource() Then
        g_Debug "main: no misc.resource..."
        MsgBox "misc.resource missing or damaged" & vbCrLf & vbCrLf & "This can happen if melon is uninstalled - try reinstalling melon in the first instance", vbCritical Or vbOKOnly, App.Title
        Exit Sub

    End If

    If g_IsIDE Then _
        Form1.Show

Dim hWnd As Long

    EZRegisterClass CLASS_NAME
    hWnd = EZ4AddWindow(CLASS_NAME, New TWindow, CLASS_NAME)

    Form1.Add "Handler is " & g_HexStr(hWnd)
    Form1.Tag = CStr(hWnd)

    g_Debug "main: started"

    With New BMsgLooper
        .Run

    End With

    g_Debug "main: ended"

    EZ4RemoveWindow hWnd
    EZUnregisterClass CLASS_NAME

    Unload Form1

End Sub

Private Function uGotMiscResource() As Boolean

    On Error Resume Next

Dim i As Long

    err.Clear
    i = processor_count()
    uGotMiscResource = (err.Number = 0)

End Function
