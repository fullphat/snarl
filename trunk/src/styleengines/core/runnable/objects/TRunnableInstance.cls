VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRunnableInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRunnableInstance.cls
    '/
    '/  Description:    Represents a single on-screen instance of a runnable style
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Const SEE_MASK_NOCLOSEPROCESS = &H40

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long

End Type

Private Declare Function ShellExecuteEx Lib "shell32" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim mTimeout As Long
Dim mContent As MImage
Dim mStyle As TRunnableStyle

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mContent

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    On Error GoTo er

    g_Debug "TRunnableInstance.UpdateContent()", LEMON_LEVEL_PROC_ENTER

    g_Debug "deleting existing image..."
    DeleteFile mStyle.Path & "content.png"

Dim szPath As String
Dim lPid As Long

    ' /* launch the style provider passing the V42 args as a single param */

'Debug.Print "running '" & mStyle.Executable & "'..."
''MsgBox Chr$(34) & mStyle.Executable & Chr$(34) & " " & Chr$(34) & NotificationInfo.Text & "#?scheme::" & NotificationInfo.Scheme & Chr$(34)
'Debug.Print "--"
'Debug.Print NotificationInfo.Text
'Debug.Print "--"

    g_Debug "launching style process " & g_Quote(mStyle.Executable) & "..."

Dim pei As SHELLEXECUTEINFO
Dim hr As Long
Dim t As Long

    With pei
        .cbSize = LenB(pei)
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .lpFile = mStyle.Executable
        .lpParameters = Chr$(34) & NotificationInfo.Text & "#?scheme::" & NotificationInfo.Scheme & Chr$(34)
        .lpDirectory = mStyle.Path

    End With

    t = GetTickCount()

    If ShellExecuteEx(pei) = 0 Then
        g_Debug "ShellExecuteEx() failed: " & g_ApiError(), LEMON_LEVEL_CRITICAL

    Else
        ' /* wait for it to finish */
        hr = WaitForSingleObject(pei.hProcess, mTimeout)
        CloseHandle pei.hProcess
        
        Select Case hr
        Case WAIT_TIMEOUT
            g_Debug "process didn't create content within time limit", LEMON_LEVEL_CRITICAL

        Case WAIT_OBJECT_0
            g_Debug "process completed", LEMON_LEVEL_INFO

        Case Else
            g_Debug "WaitForSingleObject: " & g_HexStr(hr)

        End Select

    End If

    g_Debug "took " & CStr(GetTickCount() - t) & " ms"

Dim pb As mfxBitmap

    If mStyle.IsRedirect Then
        g_Debug "redirect style - assume completed ok"

    Else
        ' /* style should have created a PNG for us... */
        g_Debug "getting content..."
        Set pb = load_image_obj(mStyle.Path & "content.png")
        If is_valid_image(pb) Then
            Set mContent = pb

        Else
            g_Debug "style returned no/unusable content", LEMON_LEVEL_CRITICAL

        End If

    End If

    ' /* zap it now we have it */

    g_Debug "deleting content..."
    DeleteFile mStyle.Path & "content.png"
    Exit Sub

er:
    Debug.Print "xx: " & Err.Description
    Resume Next

End Sub

Public Sub Init(ByRef Style As TRunnableStyle, ByVal Timeout As Long)

    Set mStyle = Style
    mTimeout = Timeout

End Sub
