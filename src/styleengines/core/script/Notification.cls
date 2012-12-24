VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type T_ENTRY
    Name As String
    Value As String

End Type

Dim mItemSplit As String
Dim mValSplit As String

Dim mEntry() As T_ENTRY
Dim mCount As Long
Dim mIndex As Long

Friend Function bSetTo(ByVal PackedString As String, Optional ByVal ItemDelimiter As String = "#?", Optional ByVal ValueDelimiter As String = "::") As Boolean
Dim a() As String
Dim i As Long

    On Error Resume Next

    ReDim mEntry(0)
    mCount = 0

    mItemSplit = ItemDelimiter
    mValSplit = ValueDelimiter

    Err.Clear
    a() = Split(PackedString, ItemDelimiter)
    If Err.Number <> 0 Then _
        Exit Function

    If UBound(a()) >= 0 Then
        For i = 0 To UBound(a())
            uAdd a(i)

        Next i

    End If

    mIndex = 1
    bSetTo = True

End Function

Private Sub uAdd(ByVal PackedEntry As String)

    If mValSplit = "" Then
        ' /* entries are not pairs, so just add as a named entry with no value */
        bAdd PackedEntry, ""
        Exit Sub

    End If

Dim i As Long

    i = InStr(PackedEntry, mValSplit)
    If i = 0 Then
        ' /* no marker: add as named value with no data */
        bAdd PackedEntry, ""

    Else
        bAdd g_SafeLeftStr(PackedEntry, i - 1), g_SafeRightStr(PackedEntry, Len(PackedEntry) - i - (Len(mValSplit) - 1))

    End If

End Sub

Friend Sub bAdd(ByVal Name As String, ByVal Data As String)

    mCount = mCount + 1
    ReDim Preserve mEntry(mCount)
    With mEntry(mCount)
        .Name = Name
        .Value = Data

    End With

End Sub

'Public Function AsString() As String
'Dim i As Long
'
'    If mCount Then
'        For i = 1 To mCount
'            If mEntry(i).Name <> "" Then
'                AsString = AsString & mEntry(i).Name & mValSplit & mEntry(i).Value & IIf(i < mCount, mItemSplit, "")
'
'            End If
'        Next i
'    End If
'
'End Function

Public Function Find(ByVal Name As String, ByRef Value As String) As Boolean
Dim i As Long

    i = IndexOf(Name)
    If i <> 0 Then
        Value = mEntry(i).Value
        Find = True

    End If

End Function

Public Function ValueOf(ByVal Name As String) As String
Dim sz As String

    If Find(Name, sz) Then _
        ValueOf = sz

End Function

Public Function Exists(ByVal Name As String) As Boolean

    Exists = (IndexOf(Name) <> 0)

End Function

Public Sub Rewind()

    mIndex = 1

End Sub

Public Function GetNextItem(ByRef Name As String, ByRef Value As String) As Boolean

    If (mIndex < 1) Or (mIndex > mCount) Then _
        Exit Function

    Name = mEntry(mIndex).Name
    Value = mEntry(mIndex).Value
    mIndex = mIndex + 1
    GetNextItem = True

End Function

Public Function Count() As Long

    Count = mCount

End Function

Public Function EntryAt(ByVal Index As Long, ByRef Name As String, ByRef Value As String) As Boolean

    If (Index > 0) And (Index <= mCount) Then
        Name = mEntry(Index).Name
        Value = mEntry(Index).Value
        EntryAt = True

    End If

End Function


' // V49 //
'
'Public Sub Update(ByVal Name As String, ByVal Data As String, Optional ByVal AddIfMissing As Boolean = False)
'Dim i As Long
'
'    i = IndexOf(Name)
'    If (i = 0) And (AddIfMissing) Then
'        Add Name, Data
'
'    ElseIf i <> 0 Then
'        mEntry(i).Value = Data
'
'    End If
'
'End Sub

Public Function IndexOf(ByVal Name As String) As Long
Static i As Long

    If mCount Then
        Name = LCase$(Name)

        For i = 1 To mCount
            If LCase$(mEntry(i).Name) = Name Then
                IndexOf = i
                Exit Function

            End If
        Next i

    End If

End Function

Friend Sub bRemove(ByVal Index As Long)
Dim i As Long

    If (Index > 0) And (Index <= mCount) Then
        If Index < mCount Then
            For i = Index To (mCount - 1)
                LSet mEntry(i) = mEntry(i + 1)

            Next i
        End If

        mCount = mCount - 1
        ReDim Preserve mEntry(mCount)

    End If

End Sub

Public Function Hint() As String
Dim sz As String
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            sz = sz & mEntry(i).Name & IIf(i < mCount, ",", "")

        Next i
    End If

    Hint = sz

End Function
