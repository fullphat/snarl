VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TReminderPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IsAlarmPanel As Boolean

Dim mPanel As BPrefsPanel
Dim mhWndOwner As Long
Dim mItem As BTagItem

Dim mDate As Date
Dim mHour As Integer
Dim mMinute As Integer
Dim mDay As Integer
Dim mMonth As Integer
Dim mYear As Integer

Public Event Closed()
Public Event Save(Item As BTagItem)

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim fUpdateDate As Boolean

    Select Case Control.GetName()
    Case "time"
        fUpdateDate = True

        Select Case Value
        Case 1
            ' /* hour */
            If g_IsPressed(VK_SHIFT) Then
                mHour = mHour - 1
                If mHour < 0 Then _
                    mHour = 23

            Else
                mHour = mHour + 1
                If mHour > 23 Then _
                    mHour = 0

            End If

            prefskit_SetItem Control, Val(Value), , Format$(mHour, "00")

        Case 2
            ' /* minute */
            If g_IsPressed(VK_SHIFT) Then
                mMinute = mMinute - 5
                If mMinute < 0 Then _
                    mMinute = 55

            Else
                mMinute = mMinute + 5
                If mMinute > 55 Then _
                    mMinute = 0

            End If

            prefskit_SetItem Control, Val(Value), , Format$(mMinute, "00")

        End Select


    Case "date"
        fUpdateDate = True

        Select Case Value
        Case 1
            ' /* day */
            If g_IsPressed(VK_SHIFT) Then
                mDay = mDay - 1
                If mDay < 1 Then _
                    mDay = g_DaysInMonth(DateSerial(mYear, mMonth, 1))

            Else
                mDay = mDay + 1
                If mDay > g_DaysInMonth(DateSerial(mYear, mMonth, 1)) Then _
                    mDay = 1

            End If

'            prefskit_SetItem Control, Val(Value), , Format$(mDay, "ddd dd")

        Case 2
            ' /* month */
            If g_IsPressed(VK_SHIFT) Then
                mMonth = mMonth - 1
                If mMonth < 1 Then _
                    mMonth = 12

            Else
                mMonth = mMonth + 1
                If mMonth > 12 Then _
                    mMonth = 1

            End If

            prefskit_SetItem Control, Val(Value), , MonthName(mMonth, True)

        Case 3
            ' /* year */
            If g_IsPressed(VK_SHIFT) Then
                mYear = mYear - 1
                If mYear < Year(Now) Then _
                    mYear = Year(Now) + 20

            Else
                mYear = mYear + 1
                If mYear > Year(Now) + 20 Then _
                    mYear = Year(Now)

            End If

            prefskit_SetItem Control, Val(Value), , CStr(mYear)

        End Select


    Case "repeat_toolbar"
        uSetAlarmRepeat Val(Value)

    End Select

    If fUpdateDate Then
        mDate = DateSerial(mYear, mMonth, mDay) + TimeSerial(mHour, mMinute, 0)
        prefskit_SetItemEx mPanel, "date", , 1, Format$(mDate, "ddd dd")

    End If

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim wRepeat As Integer
Dim pa As TAlarm
Dim pr As TReminder
Dim dDue As Date
Dim n As Integer
Dim i As Integer

    Select Case Control.GetName()
    Case "save"

        If Me.IsAlarmPanel Then

            ' /* build repeat */

            n = 1
            For i = 7 To 1 Step -1
                If prefskit_GetItemEx(mPanel, "repeat", "checked", i) = "1" Then _
                    wRepeat = wRepeat Or n
    
                n = n * 2
    
            Next i

            ' /* build due time */

            dDue = TimeSerial(mHour, mMinute, 0)

            Set pa = mItem
            pa.Update prefskit_SafeGetText(mPanel, "desc"), dDue, _
                      wRepeat, prefskit_SafeGetText(mPanel, "style"), prefskit_GetValue(mPanel, "icon"), Val(prefskit_GetValue(mPanel, "priority")) - 2

        Else
            ' /* reminder */

            If prefskit_SafeGetText(mPanel, "desc") = "" Then
                MsgBox "Must have a description", vbExclamation Or vbOKOnly, App.Title
                Exit Sub
    
            End If

            ' /* build due date/time */

            dDue = DateSerial(mYear, mMonth, mDay) + TimeSerial(mHour, mMinute, 0)

            ' /* day repeat is used for weekly and fortnightly repeats */

            wRepeat = CInt(prefskit_GetValue(mPanel, "repeat_type")) - 1

            Set pr = mItem
            pr.Update prefskit_SafeGetText(mPanel, "desc"), dDue, _
                      wRepeat, prefskit_SafeGetText(mPanel, "style"), prefskit_GetValue(mPanel, "icon"), Val(prefskit_GetValue(mPanel, "priority")) - 2

        End If

        RaiseEvent Save(mItem)
        mPanel.Quit

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    EnableWindow mhWndOwner, -1
    g_ShowWindow mhWndOwner, True, True
    RaiseEvent Closed

End Sub

Private Sub KPrefsPanel_Ready()
Dim pc As BControl
Dim pa As TAlarm
Dim n As Long
Dim i As Long

    If (mPanel.Find("repeat", pc)) And (Me.IsAlarmPanel) Then
        ' /* set the selected repeat */
        Set pa = mItem
        If pa.Repeat <> 0 Then
            n = 1
            For i = 7 To 1 Step -1
                If (pa.Repeat And n) Then _
                    prefskit_SetItem pc, i, "checked", 1&
    
                n = n * 2
    
            Next i
        End If

    End If

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long, ByRef Item As BTagItem, Optional ByVal IsNew As Boolean)

    mhWndOwner = hWndPanel
    Set mItem = Item

Dim pr As TReminder
Dim pa As TAlarm

    If Me.IsAlarmPanel Then
        Set pa = Item
        mHour = pa.HourVal
        mMinute = pa.MinuteVal

    Else
        ' /* reminder */
        Set pr = Item
        mDate = pr.Due
        mHour = Hour(mDate)
        mMinute = Minute(mDate)
        mDay = Day(mDate)
        mMonth = Month(mDate)
        mYear = Year(mDate)

    End If

Dim pp As BPrefsPage
Dim pm As CTempMsg
Dim pc As BControl
Dim sz As String
Dim i As Long

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle IIf(Me.IsAlarmPanel, "Alarm Details", "Reminder Details")
        .SetWidth 460

        ' /* page */

        Set pp = new_BPrefsPage("", , Me)

        With pp
            .SetMargin 90

            If Me.IsAlarmPanel Then

                ' /* alarm controls */

                .Add new_BPrefsControl("fancytoolbar", "time", Format$(mHour, "00") & "|" & Format$(mMinute, "00"), "Time:")

                Set pm = New CTempMsg
                pm.Add "checkboxes", 1&
                Set pc = new_BPrefsControl("listbox", "repeat", "Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday", "Repeat:", , pm)
                pc.SizeTo 0, 158
                .Add pc

                .Add new_BPrefsControl("fancytoolbar", "repeat_toolbar", "None|Daily|Weekdays|Weekends")

                .Add new_BPrefsControl("fancyedit", "desc", pa.Label, "Label:")
                i = pa.Priority

            Else

                ' /* reminder controls */

                .Add new_BPrefsControl("fancytoolbar", "date", Format$(mDate, "ddd dd") & "|" & MonthName(mMonth, True) & "|" & CStr(mYear), "Date:")
                .Add new_BPrefsControl("fancytoolbar", "time", Format$(mHour, "00") & "|" & Format$(mMinute, "00"), "Time:")
                .Add new_BPrefsControl("fancycycle", "repeat_type", "Never|Daily|Weekly|Fortnightly|Monthly|Yearly", "Repeat:", CStr(pr.RepeatType + 1))
                .Add new_BPrefsControl("fancyedit", "desc", pr.Description, "Description:")

                Set pm = New CTempMsg
                sz = pr.Icon(False)
                pm.Add "path", IIf(sz = "", g_GetSystemFolderStr(CSIDL_MYPICTURES), sz)
                .Add new_BPrefsControl("fancyfilepicker", "icon", sz, "Icon:", , pm)
                .Add new_BPrefsControl("fancyedit", "style", pr.Style, "Style:")
                i = pr.Priority

            End If

            .Add new_BPrefsControl("fancycycle", "priority", "Low|Normal|High", "Priority:", CStr(i + 2))

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "save", "Save")

        End With

        .AddPage pp

        ' /* page */

        .Go
        g_SetWindowIconToAppResourceIcon .hWnd
        SetWindowLong .hWnd, GWL_HWNDPARENT, hWndPanel
        EnableWindow hWndPanel, 0

Dim rcOwner As RECT
Dim rcThis As RECT

        GetWindowRect hWndPanel, rcOwner
        GetWindowRect .hWnd, rcThis
        g_RectNormalise rcThis

        g_MoveWindow .hWnd, _
                     rcOwner.Left + Fix(((rcOwner.Right - rcOwner.Left) - rcThis.Right) / 2), _
                     rcOwner.Top + Fix(((rcOwner.Bottom - rcOwner.Top) - rcThis.Bottom) / 2)

        g_ShowWindow .hWnd, True, True

    End With

End Sub

Private Function uIndexFromMinutes(ByVal Minutes As Integer) As String

    If (Minutes < 0) Or (Minutes > 55) Then _
        Minutes = 0

    uIndexFromMinutes = (Minutes \ 5) + 1

End Function

Private Sub uSetAlarmRepeat(ByVal RepeatType As Long)
Dim pc As BControl
Dim i As Long

    If Not mPanel.Find("repeat", pc) Then _
        Exit Sub

    Select Case RepeatType
    Case 1
        ' /* none */
        For i = 1 To 7
            prefskit_SetItem pc, i, "checked", 0&

        Next i

    Case 2
        ' /* daily */
        For i = 1 To 7
            prefskit_SetItem pc, i, "checked", 1&

        Next i
        
    Case 3
        ' /* weekdays */
        For i = 1 To 7
            prefskit_SetItem pc, i, "checked", IIf(i < 6, 1&, 0&)

        Next i
    
    Case 4
        ' /* weekends */
        For i = 1 To 7
            prefskit_SetItem pc, i, "checked", IIf(i > 5, 1&, 0&)

        Next i
    
    End Select

End Sub
