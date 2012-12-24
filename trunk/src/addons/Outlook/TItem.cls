VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mObj As Object
Dim mToken As Long
Dim mPassword As String

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mToken)

End Function

Private Function BTagItem_Value() As String
End Function

Public Sub SetTo(ByRef Item As Object, ByVal Token As Long, ByVal Password As String)

    Set mObj = Item
    mToken = Token
    mPassword = Password

End Sub

Public Sub DoAction(ByVal Action As Integer, ByVal MarkAsRead As Boolean)
Dim pa As AppointmentItem
Dim pt As MeetingItem

    On Error GoTo eh

    ' /* if it's a MeetingItem, try to get the associated appointment */

    If TypeOf mObj Is MeetingItem Then
        Set pt = mObj
        Set pa = pt.GetAssociatedAppointment(False)

    End If

    Select Case Action

    Case 1
        ' /* mark as read */

    Case 2
        mObj.Reply.Display

    Case 3
        mObj.ReplyAll.Display

    Case 4
        mObj.Forward.Display

    Case 9
        mObj.Display

    Case 12
        ' /* accept */
        pa.Respond(olMeetingAccepted).Send
        mObj.Delete

    Case 13
        ' /* tentative */
        pa.Respond(olMeetingTentative).Display
        mObj.Delete

    Case 14
        ' /* decline */
        pa.Respond(olMeetingDeclined).Display
        mObj.Delete

    Case 99
        ' /* open appointment */
        Set pa = mObj
        pa.Display

    End Select

    If MarkAsRead Then _
        mObj.UnRead = False

    snDoRequest "hide?token=" & mToken & "&password=" & mPassword
    Exit Sub

eh:
    g_Debug "TItem.DoAction(): " & Err.Description, LEMON_LEVEL_CRITICAL
    Resume Next

End Sub

Public Sub DoClicked()

    mObj.Display

End Sub

Public Sub MarkAsRead()

    mObj.UnRead = False

End Sub
