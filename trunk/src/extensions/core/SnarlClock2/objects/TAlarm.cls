VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TAlarm.cls
    '/
    '/  Description:    Represents a single alarm
    '/
    '/  Created:        6-Sep-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

                             ' // Bit  M | T | W | T | F | S | S
'Dim mRepeat As Integer      ' //      6   5   4   3   2   1   0
Dim mConfig As CConfSection
Dim mDone As Boolean
Dim mDue As Date

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mConfig.Name

End Function

Private Function BTagItem_Value() As String
End Function

Private Sub Class_Initialize()
Dim m As Integer
Dim n As Integer

    Set mConfig = New CConfSection
    With mConfig
        .SetName g_CreateGUID(True)

        ' /* round up to nearest 5 minutes */

        m = Minute(Now)
        n = Round((m \ 5) + 1) * 5
        n = n - m

        mDue = DateAdd("n", n, Now)
        Debug.Print "new alarm: calculated time is " & mDue

        .Add "hour", CStr(Hour(mDue))
        .Add "minute", CStr(Minute(mDue))
        .Add "repeat", "0"
        .Add "enabled", "1"

    End With

End Sub

Public Sub Update(ByVal Label As String, ByVal Due As Date, ByVal Repeat As Integer, ByVal Style As String, ByVal Icon As String, ByVal Priority As Integer)

    With mConfig
        .Update "hour", CStr(Hour(Due))
        .Update "minute", CStr(Minute(Due))
        .Update "repeat", Repeat
        .Update "label", Label
        .Update "priority", CStr(Priority)

    End With

    mDue = Due

End Sub

Public Function Label() As String

    Label = mConfig.GetValueWithDefault("label")

End Function

Public Function Config() As CConfSection

    Set Config = mConfig

End Function

Public Function SetFrom(ByRef Config As CConfSection) As Boolean

    If (Config Is Nothing) Then _
        Exit Function

Dim n As Long

    Set mConfig = Config
    If (mConfig.Name = "") Then _
        Exit Function

    mDue = TimeSerial(CInt(mConfig.GetValueWithDefault("hour")), CInt(mConfig.GetValueWithDefault("minute")), 0)
    SetFrom = True

End Function

Public Function Details() As String
Dim i As Integer

    i = Me.Repeat

    If i = 0 Then
        Details = "Today"
'        Details = CStr(Me.HourVal) & ":" & Format$(Me.MinuteVal, "00")
'        Exit Function

    ElseIf i = &H7F Then
        Details = "Every day"

    ElseIf (i And &H7F) = 3 Then
        Details = "Weekends"

    ElseIf (i And &H7F) = &H7C Then
        Details = "Weekdays"

    Else
        Details = "Every "
        If (i And &H40) Then _
            Details = Details & "Monday, "

        If (i And &H20) Then _
            Details = Details & "Tuesday, "

        If (i And &H10) Then _
            Details = Details & "Wednesday, "

        If (i And &H8) Then _
            Details = Details & "Thursday, "

        If (i And &H4) Then _
            Details = Details & "Friday, "

        If (i And &H2) Then _
            Details = Details & "Saturday, "

        If (i And &H1) Then _
            Details = Details & "Sunday, "

        Details = g_SafeLeftStr(Details, Len(Details) - 2)

    End If

    Details = Details & " at " & CStr(Me.HourVal) & ":" & Format$(Me.MinuteVal, "00")

End Function

Public Function IsDue() As Boolean

    If (Not Me.IsEnabled) Or (mDone) Then _
        Exit Function

Dim bRightDay As Boolean
Dim i As Integer

    i = Me.Repeat

    If i = 0 Then
        ' /* one-off */
        IsDue = (Hour(Now) = Hour(mDue) And Minute(Now) = Minute(mDue))

    Else
        Select Case Weekday(Now, vbMonday)
        Case 1
            ' /* monday */
            bRightDay = (i And &H40)

        Case 2
            ' /* tuesday */
            bRightDay = (i And &H20)

        Case 3
            ' /* wednesday */
            bRightDay = (i And &H10)

        Case 4
            ' /* thursday */
            bRightDay = (i And &H8)

        Case 5
            ' /* friday */
            bRightDay = (i And &H4)

        Case 6
            ' /* saturday */
            bRightDay = (i And &H2)

        Case 7
            ' /* friday */
            bRightDay = (i And &H1)

        Case Else
            Debug.Print "weekday is " & Weekday(Now, vbMonday)

        End Select

        IsDue = (bRightDay) And (Hour(Now) = Hour(mDue) And Minute(Now) = Minute(mDue))

    End If

End Function

Public Function Style() As String

    Style = mConfig.GetValueWithDefault("style")

End Function

Public Function HourVal() As Integer

    HourVal = Hour(mDue)

End Function

Public Function MinuteVal() As Integer

    MinuteVal = Minute(mDue)

End Function

Public Function IsEnabled() As Boolean

    If Not (mConfig Is Nothing) Then _
        IsEnabled = (mConfig.GetValueWithDefault("enabled") = "1")

End Function

Public Sub SetEnabled(ByVal Enabled As Boolean)

    If Not (mConfig Is Nothing) Then _
        mConfig.Update "enabled", IIf(Enabled, "1", "0")

End Sub

Public Function Repeat() As Integer

    If Not (mConfig Is Nothing) Then _
        Repeat = CInt(mConfig.GetValueWithDefault("repeat"))

End Function

Public Sub SetAsDone()

    If Me.Repeat = 0 Then
        ' /* if there's no repeat set, then disable it */
        Me.SetEnabled False

    Else
        ' /* otherwise just set transient marker */
        mDone = True

    End If

End Sub

Public Function Priority() As Integer

    On Error Resume Next

    If Not (mConfig Is Nothing) Then _
        Priority = CInt(mConfig.GetValueWithDefault("priority"))

End Function


