Attribute VB_Name = "mForwardingSupport"
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           mForwardingSupport.bas
    '/
    '/  Description:    Notification forwarding via SNP2.0 support routines
    '/
    '/  © 2010 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public gRemoteComputers As ConfigSection

Dim mDest() As TForwarder
Dim mCount As Long
Dim mUID As Long

Public Sub g_ForwardInit()

    mUID = &HE0

End Sub

Public Sub g_ForwardNotification(ByRef Details As notification_info)
Static i As Long

    If (gRemoteComputers Is Nothing) Then
        g_Debug "g_ForwardNotification(): remote computers list is invalid", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    If gRemoteComputers.CountEntries = 0 Then
        g_Debug "g_ForwardNotification(): remote computers list is empty", LEMON_LEVEL_WARNING
        Exit Sub

    End If

    With gRemoteComputers

        For i = 1 To .CountEntries
            Debug.Print .EntryAt(i).Name & " / " & .EntryAt(i).Value
            uAddSender .EntryAt(i).Name, Details

        Next i

    End With

End Sub

Private Sub uAddSender(ByVal Destination As String, ByRef Details As notification_info)

    mCount = mCount + 1
    ReDim Preserve mDest(mCount)

    Set mDest(mCount) = New TForwarder
    mDest(mCount).Go Destination, Details, mUID

    mUID = mUID + 4

End Sub

Public Sub g_ForwardRemove(ByVal UID As Long)

End Sub

