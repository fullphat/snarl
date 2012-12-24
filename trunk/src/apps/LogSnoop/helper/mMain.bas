Attribute VB_Name = "mMain"
Option Explicit

Dim mLog As String

Public Sub Main()

    On Error Resume Next

    l3OpenLog g_MakePath(App.Path) & "logsnoophelper.txt"

Dim hAppEventLog As Long
Dim hEvent As Long

    hAppEventLog = OpenEventLog(vbNullString, Command$)
    If hAppEventLog = 0 Then
        g_Debug "main: couldn't open log '" & Command$ & "'"
        l3CloseLog
        Exit Sub

    End If

    mLog = Command$
    g_Debug "main: watching '" & Command$ & "'..."

    hEvent = CreateEvent(0, True, False, vbNullString)
    If hEvent = 0 Then
        g_Debug "main: failed to create event"
        l3CloseLog
        Exit Sub

    End If

    g_Debug "main: most recent event is:"
    uGetMostRecentEvent hAppEventLog

    g_Debug "main: adding change notifier..."
    NotifyChangeEventLog hAppEventLog, hEvent

Dim hr As Long

    Do
        hr = WaitForSingleObject(hEvent, 50)

        Select Case hr
        Case WAIT_TIMEOUT
            DoEvents

        Case 0
            ' /* signalled */
            g_Debug "main: event notified:"
            uGetMostRecentEvent hAppEventLog

        Case Else
            g_Debug "main: WaitForSingleObject() error " & CStr(hr)

        End Select

    Loop

    g_Debug "main: ending..."

    CloseHandle hEvent
    CloseEventLog hAppEventLog

    l3CloseLog

End Sub

Private Sub uGetMostRecentEvent(ByVal hEventLog As Long)
Dim c As Long

    ' /* some error testing... */

    If GetNumberOfEventLogRecords(hEventLog, c) = 0 Then _
        Exit Sub

Dim rx As EVENTLOGRECORD
Dim cbNeeded As Long
Dim cbRead As Long
Dim b() As Byte

    ReDim b(2048)

    If ReadEventLog(hEventLog, EVENTLOG_SEQUENTIAL_READ Or EVENTLOG_BACKWARDS_READ, 0, b(0), 2048, cbRead, cbNeeded) = 0 Then
        Debug.Print "failed: needed = " & cbNeeded
        Exit Sub

    End If

    Debug.Print "read: " & cbRead
    CopyMemory rx, b(0), Len(rx)

Dim ox As Long
Dim cb As Long

Dim szIcon As String
Dim szTitle As String
Dim szText As String
Dim szSource As String

    With rx
        g_Debug "uGetMostRecentEvent(): details..."
        g_Debug "length=" & CStr(.Length) & " eventid=" & CStr(.EventID) & " type=" & CStr(.EventType) & _
                " numstrings=" & CStr(.NumStrings) & " tgen=" & CStr(.TimeGenerated) & " twri=" & CStr(.TimeWritten)

        Select Case .EventType

        Case EVENTLOG_ERROR_TYPE
            szIcon = "!system-critical"

        Case EVENTLOG_WARNING_TYPE
            szIcon = "!system-warning"

        Case EVENTLOG_INFORMATION_TYPE
            szIcon = "!system-info"

        Case EVENTLOG_AUDIT_SUCCESS
            szIcon = "!system-yes"

        Case EVENTLOG_AUDIT_FAILURE
            szIcon = "!system-no"

        End Select

        ox = VarPtr(b(0)) + Len(rx)
        szSource = g_CopyStrA(ox, cb)

        ox = ox + cb + 1
        cb = 0
        szTitle = g_CopyStrA(ox, cb)
        szText = g_CopyStrA(VarPtr(b(0)) + .StringOffset)

        g_Debug "source=" & szSource & " description=" & szText & " computer=" & szTitle & _
                " numstrings=" & CStr(.NumStrings) & " tgen=" & CStr(.TimeGenerated) & " twri=" & CStr(.TimeWritten)

        szTitle = szSource & " on " & szTitle
        szText = szText & vbCrLf & _
                          "Event: " & .EventID & vbCrLf & _
                          "Category: " & .EventCategory & vbCrLf & _
                          "Generated: " & g_UnixTimeToDate(.TimeGenerated) & vbCrLf & _
                          "Logged: " & g_UnixTimeToDate(.TimeWritten)

    End With

    ' /* we assume Snarl is around... */

    snDoRequest "notify?app-sig=" & App.ProductName & _
                "&class=" & "" & _
                "&title=" & szTitle & _
                "&icon=" & szIcon & _
                "&text=" & szText & _
                "&label-subtext=" & mLog & " log"

'Dim i As Long

'    For i = 0 To cbRead
'        Select Case b(i)
'        Case Is >= 32
'            Debug.Print Chr$(b(i));
'
'        Case 13, 10
'            Debug.Print
'
'        Case Else
'            Debug.Print "."
'
'        End Select
'
'    Next i

End Sub

