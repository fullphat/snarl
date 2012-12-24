VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TReminder.cls
    '/
    '/  Description:    Represents a single reminder
    '/
    '/  Created:        5-Sep-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mDue As Date
Dim mConfig As CConfSection
Dim mHour As Integer
Dim mMin As Integer

'Dim mRepeatType As Integer  ' // daily = 1, weekly = 2, fortnightly = 3, monthly = 4, yearly = 5
Dim mDisplayed As Boolean
Dim mNextDue As Date

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
        Debug.Print "calculated time is " & mDue
        mNextDue = mDue

        .Add "date", mDue
        .Add "desc", ""
        .Add "repeat_type", "0"

    End With

End Sub

Public Sub Update(ByVal Description As String, ByVal Due As Date, ByVal RepeatType As Integer, ByVal Style As String, ByVal Icon As String, ByVal Priority As Integer)

    With mConfig
        .Update "desc", Description
        .Update "date", CStr(Due)
        .Update "icon", Icon
        .Update "style", Style
        .Update "repeat_type", RepeatType
        .Update "priority", CStr(Priority)

    End With

    mDue = Due
    CalcNextDue

End Sub

Public Function Description() As String

    Description = mConfig.GetValueWithDefault("desc")

End Function

Public Function Config() As CConfSection

    Set Config = mConfig

End Function

Public Function SetFrom(ByRef Config As CConfSection) As Boolean

    If (Config Is Nothing) Then _
        Exit Function

    Set mConfig = Config
    If (mConfig.Name = "") Then _
        Exit Function

    mDue = Config.GetValueWithDefault("date", CStr(mDue))
    CalcNextDue
    SetFrom = True

End Function

Public Function Details() As String

    Select Case Me.RepeatType
    Case 0
        ' /* none */
        Details = Format$(mDue, "d mmm yyyy")

    Case 1
        ' /* daily */
        Details = "Daily"

    Case 2
        ' /* weekly */
        Details = "Every " & uWeeklyRepeatDayName()

    Case 3
        ' /* fortnightly */
        Details = "Every other " & uWeeklyRepeatDayName()

    Case 4
        ' /* monthly */
        Details = "On the " & CStr(Day(mDue)) & " of every month"

    Case 5
        ' /* yearly */
        Details = "Every " & Format$(mDue, "d mmm")

    End Select

    ' /* add time (to all) */
    Details = Details & " at " & Format$(mDue, "short time")

    ' /* add "from nnn" to all except static */
    If Me.RepeatType > 0 Then _
        Details = Details & " from " & Format$(mDue, "d mmm yyyy")

'    If mRepeat = 0 Then
'        Details = "Today"
'
'    ElseIf mRepeat = &H7F Then
'        Details = "Every day"
'
'    ElseIf (mRepeat And &H7F) = 3 Then
'        Details = "Weekends"
'
'    ElseIf (mRepeat And &H7F) = &H7C Then
'        Details = "Weekdays"
'
'    Else
'        Details = "Every "
'        If (mRepeat And &H40) Then _
'            Details = Details & "Monday, "
'
'        If (mRepeat And &H20) Then _
'            Details = Details & "Tuesday, "
'
'        If (mRepeat And &H10) Then _
'            Details = Details & "Wednesday, "
'
'        If (mRepeat And &H8) Then _
'            Details = Details & "Thursday, "
'
'        If (mRepeat And &H4) Then _
'            Details = Details & "Friday, "
'
'        If (mRepeat And &H2) Then _
'            Details = Details & "Saturday, "
'
'        If (mRepeat And &H1) Then _
'            Details = Details & "Sunday, "
'
'        Details = g_SafeLeftStr(Details, Len(Details) - 2)
'
'    End If
'
'    Details = Details & " at " & CStr(Me.HourVal) & ":" & Format$(Me.MinuteVal, "00")

End Function

Public Function IsDue(ByVal GlobalOffsetType As Long) As Boolean
Dim d As Date

    ' /* apply the global offset delta */

    Select Case GlobalOffsetType
    Case 0
        ' /* on event */
        d = mNextDue

    Case 1
        ' /* -5 mins */
        d = DateAdd("n", -5, mNextDue)

    Case 2
        ' /* -10 mins */
        d = DateAdd("n", -10, mNextDue)

    Case 3
        ' /* -30 mins */
        d = DateAdd("n", -30, mNextDue)

    Case 4
        ' /* -1 hour */
        d = DateAdd("h", -1, mNextDue)

    Case 5
        ' /* -1 day */
        d = DateAdd("d", -1, mNextDue)

    End Select

    If (mDisplayed) Or (Now() < d) Then
'        Debug.Print "TReminder.IsDue(): '" & Me.Description & "' already done/not due yet"

    Else
        IsDue = (DateDiff("n", Now, d) = 0)

    End If

End Function

Public Function Style() As String

    Style = mConfig.GetValueWithDefault("style")

End Function

Public Function Icon(ByVal UseDefaultIfNoIcon As Boolean) As String

    Icon = mConfig.GetValueWithDefault("icon")
    If ((Icon = "") Or (Not g_Exists(Icon))) And (UseDefaultIfNoIcon) Then _
        Icon = g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "reminder.png"

End Function

Public Function Due() As Date

    Due = mDue

End Function

Public Function RepeatType() As Integer
Dim n As Long

    On Error Resume Next

    n = Val(mConfig.GetValueWithDefault("repeat_type", "0"))
    n = n And &H7F&
    RepeatType = CInt(n)

End Function

Private Function uWeeklyRepeatDayName() As String
Dim i As Integer

    On Error Resume Next

    Err.Clear
    i = Weekday(mDue)
    If (Err.Number = 0) And (i <> 0) Then _
        uWeeklyRepeatDayName = WeekdayName(i, , vbSunday)

End Function

Public Function DueDelta() As String
Dim i As Integer

    i = DateDiff("n", Now, mNextDue)
    If i = 0 Then
        DueDelta = "Now"

    Else
        DueDelta = "In " & CStr(i) & " minute" & IIf(i = 1, "", "s")

    End If

End Function

Public Sub SetWasDisplayed()

    mDisplayed = True
    CalcNextDue

End Sub

Public Function Priority() As Integer

    On Error Resume Next

    If Not (mConfig Is Nothing) Then _
        Priority = CInt(mConfig.GetValueWithDefault("priority"))

End Function

Public Sub CalcNextDue()

    ' /* if it's in the future then the due date is the start date */

    If Now() < mDue Then
        mNextDue = mDue

    Else
        ' /* starting point */
    
        mNextDue = mDue
    
        Select Case Me.RepeatType
        Case 0
            ' /* none (will have been missed) */
    
        Case 1
            ' /* daily */
            Do
                mNextDue = DateAdd("d", 1, mNextDue)
    
            Loop While Now() > mNextDue
    
        Case 2
            ' /* weekly */
            Do
                mNextDue = DateAdd("ww", 1, mNextDue)
    
            Loop While Now() > mNextDue
    
        Case 3
            ' /* fortnightly */
            Do
                mNextDue = DateAdd("ww", 2, mNextDue)
    
            Loop While Now() > mNextDue
    
        Case 4
            ' /* monthly */
            Do
                mNextDue = DateAdd("m", 1, mNextDue)
    
            Loop While Now() > mNextDue
    
        Case 5
            ' /* yearly */
            Do
                mNextDue = DateAdd("yyyy", 1, mNextDue)
    
            Loop While Now() > mNextDue
    
        End Select
    End If

    g_Debug "TReminder.CalcNextDue(): '" & Me.Description & "' due " & Format$(mDue, "d mmm yyyy hh:mm") & ", next due " & Format$(mNextDue, "d mmm yyyy hh:mm")

End Sub

Public Function NextDue() As Date

    NextDue = mNextDue

End Function


