VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mGuid As String
Dim mURL As String
Dim mEnabled As Boolean
Dim mChecking As Boolean
Dim mInt As Long
Dim mLastCheckWasGood As Boolean

Dim WithEvents theCheck As CHTTPRequest3
Attribute theCheck.VB_VarHelpID = -1
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Dim mLastSuccess As Date
Dim mLastTry As Date

'Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mGuid

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mURL

End Function

Public Function Init(ByVal URL As String, ByVal Guid As String, Optional ByVal Enabled As Boolean = True) As Boolean

    mURL = URL
    mGuid = Guid

    If mGuid = "" Then _
        mGuid = g_CreateGUID(True)

    Set theCheck = New CHTTPRequest3

    Init = Check()

Const OFFSET = 5

    ' /* on success, start the timer with a +/- N second offset */

    Randomize Timer
    mInt = (60 + (CLng(Rnd * (OFFSET * 2)) - OFFSET)) * 1000
    frmMain.DebugOutput g_SafeRightStr(mURL, 80, True)
    frmMain.DebugOutput "  interval timer is " & CStr(mInt \ 1000) & " secs"

    Me.SetEnabled Enabled

'        g_Debug "TSite.Init(): bad URL '" & URL & "'", LEMON_LEVEL_CRITICAL

End Function

Public Sub Quit()

    g_Debug "TSite.Quit()", LEMON_LEVEL_PROC

    Set theTimer = Nothing

    If Not (theCheck Is Nothing) Then
        theCheck.Abort
        Set theCheck = Nothing

    End If

End Sub

Public Function Guid() As String

    Guid = mGuid

End Function

Public Function URL() As String

    URL = mURL

End Function

Private Sub theTimer_Pulse()

    g_Debug "theTimer/Pulse: " & mURL
    Check

End Sub

Public Function Check() As Boolean

    g_Debug "TSite.Check()", LEMON_LEVEL_PROC_ENTER

    If mChecking Then
        g_Debug "busy...", LEMON_LEVEL_WARNING

    Else
        g_Debug "testing '" & mURL & "'..."
        Check = theCheck.GetContent(mURL)
        mLastTry = Now()
        g_Debug "timestamp: " & CStr(mLastTry)
        mChecking = True
        frmMain.UpdateList

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub theCheck_Completed(ByVal WasSuccessful As Boolean)

    Debug.Print "done: " & mURL & ": " & WasSuccessful & " : " & (theCheck.ResponseBody <> "")
'    Debug.Print Replace$(theCheck.ResponseHeaders, vbCrLf, " ")
    Debug.Print theCheck.ResponseBody
    Debug.Print "--"

    If (Not WasSuccessful) Or (theCheck.ResponseBody = "") Then
        ' /* failed or nothing returned */
        mLastCheckWasGood = False
        g_Debug "theCheck/Completed(): failed: notifying..."
        snarl_notify App.ProductName, "", "", frmMain.Tag, _
                     "Site unreachable", _
                     mURL & " is not responding." & vbCrLf & "Last seen: " & IIf(Year(mLastSuccess) < 1900, "never", CStr(mLastSuccess)), _
                     g_MakePath(App.Path) & "failed-big.png", , , , , _
                     "action=Open link," & mURL

'        snDoRequest "notify?app-sig=" & App.ProductName & _
                    "&icon=" & g_MakePath(App.Path) & "failed.png" & _
                    "&title=Site uncontactable" & _
                    "&text=" & mURL & vbCrLf & "Last seen: " & IIf(Year(mLastSuccess) < 1900, "never", CStr(mLastSuccess)) & _
                    "&action="

    Else
        mLastCheckWasGood = True
        mLastSuccess = Now()
        g_Debug "theCheck/Completed(): success [" & CStr(mLastSuccess) & "]"

    End If

    mChecking = False
    frmMain.UpdateList

End Sub

Public Function LastSeen() As String

    If Year(mLastSuccess) < 1900 Then
        LastSeen = "never"

    ElseIf g_IsToday(mLastSuccess) Then
        LastSeen = Format$(mLastSuccess, "short time")

    Else
        LastSeen = CStr(mLastSuccess)

    End If

End Function

Public Function LastCheck() As String

    If Year(mLastTry) < 1900 Then
        LastCheck = "never"

    ElseIf g_IsToday(mLastTry) Then
        LastCheck = Format$(mLastTry, "short time")

    Else
        LastCheck = CStr(mLastTry)

    End If

End Function

Public Function IsChecking() As Boolean

    IsChecking = mChecking

End Function

Public Sub SetEnabled(ByVal Enabled As Boolean)

    mEnabled = Enabled

    If Enabled Then
        Set theTimer = new_BTimer(mInt)

    Else
        Set theTimer = Nothing

    End If

End Sub

Public Function IsEnabled() As Boolean

    IsEnabled = mEnabled

End Function

Public Function LastCheckWasGood() As Boolean

    LastCheckWasGood = mLastCheckWasGood
    
End Function
