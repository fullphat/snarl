VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNCHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mApp As TApp
Dim mNotifications As BTagList
Public Frame As BRect
Public Expanded As Boolean
Public CloseEnabled As Boolean

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mApp.Signature

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = "<header>"

End Function

Public Sub SetTo(ByRef App As TApp)

    Set mApp = App
    Set mNotifications = new_BTagList()
    Me.Expanded = True

End Sub

Public Sub Add(ByVal Title As String, ByVal Text As String, ByRef Icon As MImage, ByVal Token As Long)
Dim pp As TPopUpItem

    Set pp = New TPopUpItem
    pp.SetTo Title, Text, Icon
    pp.Token = Token
    mNotifications.Add pp

End Sub

Public Function App() As TApp

    Set App = mApp

End Function

Public Function Notifications() As BTagList

    Set Notifications = mNotifications

End Function

Public Function InCloseButton(ByRef Point As BPoint) As Boolean
Dim pr As BRect

    Set pr = Me.Frame.InsetByCopy(0, 2)
    pr.Right = pr.Right - 3
    pr.Left = pr.Right - 16
    InCloseButton = (pr.Contains(Point))

End Function

Public Sub RemoveAll()
Dim pp As TPopUpItem

    With mNotifications
        .Rewind
        Do While .GetNextTag(pp) = B_OK
            Debug.Print pp.Token & " >> " & g_NotificationRoster.MissedList.IndexOf(CStr(pp.Token))
            g_NotificationRoster.MissedList.Remove g_NotificationRoster.MissedList.IndexOf(CStr(pp.Token))

        Loop

    End With

    mNotifications.MakeEmpty

End Sub
