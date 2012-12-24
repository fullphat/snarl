VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Actions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_ACTION
    Label As String
    Command As String
    WinAPIUID As Long

End Type

Dim mAction() As T_ACTION
Dim mCount As Long
Dim mNextWinUID As Long

Public Sub Add(ByVal Label As String, ByVal Command As String)

    If (Label = "") Or (Command = "") Then _
        Exit Sub

    mNextWinUID = mNextWinUID + 1

    mCount = mCount + 1
    ReDim Preserve mAction(mCount)
    With mAction(mCount)
        .Label = Label
        .Command = Command
        .WinAPIUID = mNextWinUID

    End With

End Sub

Public Function Count() As Long

    Count = mCount

End Function

Public Sub Remove(ByVal Index As Long)

    If (Index < 1) Or (Index > mCount) Then _
        Exit Sub

Static j As Long

    If Index < mCount Then
        For j = Index To mCount - 1
            LSet mAction(j) = mAction(j + 1)
    
        Next j

    End If
    
    mCount = mCount - 1
    ReDim Preserve mAction(mCount)

End Sub

Friend Function bAsWin32Request() As String
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            bAsWin32Request = bAsWin32Request & "&action=" & mAction(i).Label & ",@" & CStr(mAction(i).WinAPIUID)

        Next i

    End If

'    If (Index > 0) And (Index <= mCount) Then _
'        AsRequest = "addaction?token=" & CStr(AppToken) & "&label=" & mAction(Index).Label & "&cmd=" & mAction(Index).Command & _
'                    IIf(Password <> "", "&password=" & Password, "")

End Function

Public Sub MakeEmpty()

    ReDim mAction(0)
    mCount = 0

End Sub

Friend Function bLookUp(ByVal WinId As Long) As String
Static i As Long

    If mCount Then
        For i = 1 To mCount
            If mAction(i).WinAPIUID = WinId Then
                bLookUp = mAction(i).Command
                Exit Function

            End If
        Next i

        Debug.Print "Actions.bLookUp(): action #" & CStr(WinId) & " not found"

    Else
        Debug.Print "Actions.bLookUp(): no actions"

    End If

End Function
