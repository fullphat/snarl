Attribute VB_Name = "mMain"
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_CLOSE = &H10
Public Const WM_TEST = &H400 + 1
Public Const WM_NOTIFICATION = &H400 + 2
Public Const WM_RELOAD = &H400 + 3
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public gDebugMode As Boolean

Private Const CLASS_NAME = "w>snarlsnooper"

Public Sub Main()
Dim hWndExisting As Long

    hWndExisting = FindWindow(CLASS_NAME, CLASS_NAME)

    If InStr(Command$, "-quit") Then
        ' /* quit any existing instance (but don't run this one) */
        If IsWindow(hWndExisting) <> 0 Then _
            SendMessage hWndExisting, WM_CLOSE, 0, ByVal 0&

    ElseIf InStr(Command$, "-reload") Then
        ' /* if an existing instance is running, tell it to reload tasks */
        If IsWindow(hWndExisting) <> 0 Then
            SendMessage hWndExisting, WM_RELOAD, 0, ByVal 0&

        End If

    End If

    If hWndExisting <> 0 Then _
        Exit Sub

    gDebugMode = (InStr(Command$, "-debug") <> 0)

    If Not uGotMiscResource() Then
        MsgBox "misc.resource missing or damaged" & vbCrLf & vbCrLf & "This can happen if melon is uninstalled - try reinstalling melon in the first instance", vbCritical Or vbOKOnly, App.Title
        Exit Sub

    End If

    If gDebugMode Then _
        Form1.Show

Dim hwnd As Long

    EZRegisterClass CLASS_NAME
    hwnd = EZAddWindow(CLASS_NAME, New TWindow, CLASS_NAME)

    Form1.List1.AddItem "window: " & g_HexStr(hwnd)
    Form1.Tag = CStr(hwnd)

    With New BMsgLooper
        .Run

    End With

    EZRemoveWindow hwnd
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
