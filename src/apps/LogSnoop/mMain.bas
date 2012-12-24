Attribute VB_Name = "mMain"
Option Explicit

Public ghWnd As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim lp As Long

    GetWindowThreadProcessId hwnd, lp
    If lp = lParam Then
        If (g_ClassName(hwnd) = "ThunderRT6Main") And (g_WindowText(hwnd) = "LogSnoopHelper") Then
            Debug.Print "EnumWindowsProc(): found helper window"
            ghWnd = hwnd
            Exit Function

        End If

    End If

    EnumWindowsProc = -1

End Function


