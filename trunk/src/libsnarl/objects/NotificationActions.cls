VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotificationActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_ACTION
    Label As String
    Command As String

End Type

Dim mAction() As T_ACTION
Dim mCount As Long

Public Sub Add(ByVal Label As String, ByVal Command As String)

    If (Label = "") Or (Command = "") Then _
        Exit Sub

    mCount = mCount + 1
    ReDim Preserve mAction(mCount)
    With mAction(mCount)
        .Label = Label
        .Command = Command

    End With

End Sub

Public Function Count() As Long

    Count = mCount

End Function

Public Function AsRequest(ByVal Index As Long, ByVal AppToken As Long, ByVal Password As String) As String

    If (Index > 0) And (Index <= mCount) Then _
        AsRequest = "addaction?token=" & CStr(AppToken) & "&label=" & mAction(Index).Label & "&cmd=" & mAction(Index).Command & _
                    IIf(Password <> "", "&password=" & Password, "")

End Function

Public Sub MakeEmpty()

    ReDim mAction(0)
    mCount = 0

End Sub

