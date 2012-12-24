VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Destinations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mDest() As String
Dim mCount As Long
Dim mSendLocal As Boolean

Private Sub Class_Initialize()

    mSendLocal = True

End Sub

Public Sub Add(ByVal IPAddress As String)
Static i As Long
Static j As Long

    i = Me.IndexOf(IPAddress)
    If i Then _
        Exit Sub

    mCount = mCount + 1
    ReDim Preserve mDest(mCount)
    mDest(mCount) = IPAddress

End Sub

Public Sub Remove(ByVal IPAddress As String)
Static i As Long
Static j As Long

    i = Me.IndexOf(IPAddress)
    If i Then
        If i < mCount Then
            For j = i To (mCount - 1)
                mDest(j) = mDest(j + 1)

            Next j

        End If

        mCount = mCount - 1
        ReDim Preserve mDest(mCount)

    End If

End Sub

Public Sub MakeEmpty()

    mCount = 0
    ReDim mDest(0)

End Sub

Public Property Get SendToLocalInstance() As Boolean

    SendToLocalInstance = mSendLocal

End Property

Public Property Let SendToLocalInstance(ByVal vNewValue As Boolean)

    mSendLocal = vNewValue

End Property

Public Function IndexOf(ByVal IPAddress As String) As Long
Static i As Long

    If mCount Then
        For i = 1 To mCount
            If mDest(i) = IPAddress Then
                IndexOf = i
                Exit Function

            End If
        Next i
    End If

End Function

Friend Function bDestination(ByVal Index As Long) As String

    If (Index > 0) And (Index <= mCount) Then _
        bDestination = mDest(Index)

End Function

Public Function Count() As Long

    Count = mCount

End Function
