Attribute VB_Name = "mMain"
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WM_CLOSE = &H10
Public Const WM_NOTIFICATION = &H400 + 2
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public gDebugMode As Boolean
Public gVerboseMode As Boolean

Private Const CLASS_NAME = "w>snarlware"

Public Sub Main()
Dim hWndExisting As Long

    hWndExisting = FindWindow(CLASS_NAME, CLASS_NAME)

    If InStr(Command$, "-quit") Then
        ' /* quit any existing instance (but don't run this one) */
        If IsWindow(hWndExisting) <> 0 Then _
            SendMessage hWndExisting, WM_CLOSE, 0, ByVal 0&

        Exit Sub

'    ElseIf InStr(Command$, "-reload") Then
'        ' /* if an existing instance is running, tell it to reload tasks */
'        If IsWindow(hWndExisting) <> 0 Then
'            SendMessage hWndExisting, WM_RELOAD, 0, ByVal 0&
'
'        End If
    End If

    If hWndExisting <> 0 Then _
        Exit Sub

    If Not uGotMiscResource() Then
        MsgBox "misc.resource missing or damaged" & vbCrLf & vbCrLf & "This can happen if melon is uninstalled - try reinstalling melon in the first instance", vbCritical Or vbOKOnly, App.Title
        Exit Sub

    End If

    gDebugMode = (InStr(Command$, "-debug") <> 0)
    gVerboseMode = (InStr(Command$, "-verbose") <> 0)

    If gDebugMode Then _
        Form1.Show

Dim hWnd As Long

    l3OpenLog "%APPDATA%\snarlware.log"

    EZRegisterClass CLASS_NAME
    hWnd = EZAddWindow(CLASS_NAME, New TWindow, CLASS_NAME)

    Form1.Add "window: " & g_HexStr(hWnd)
    Form1.Tag = CStr(hWnd)

    With New BMsgLooper
        .Run

    End With

    EZRemoveWindow hWnd
    EZUnregisterClass CLASS_NAME
    Unload Form1

End Sub

Private Function uGotMiscResource() As Boolean

    On Error Resume Next

Dim i As Long

    Err.Clear
    i = processor_count()
    uGotMiscResource = (Err.Number = 0)

End Function
