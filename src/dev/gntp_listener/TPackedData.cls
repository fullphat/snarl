VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TPackedData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           BPackedData.cls
    '/
    '/  Description:    Manages a packed melon string
    '/
    '/  Created:        18-Mar-2009
    '/
    '/  Updated:         5-Oct-2010
    '
    '/  © 2009,2010 k23 productions
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* creates or decodes a packed string - defaults are:
    '       #? delimits entry pairs
    '       :: splits an entry pair into name and value
    ' */

Private Type T_ENTRY
    Name As String
    Value As String

End Type

Dim mItemSplit As String
Dim mValSplit As String

Dim mEntry() As T_ENTRY
Dim mCount As Long
Dim mIndex As Long

Implements BObject

Private Function BObject_Duplicate() As melon4.BObject
Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo Me.AsString, mItemSplit, mValSplit
    Set BObject_Duplicate = ppd

End Function

Private Function BObject_Init(Args As String) As melon4.B_STATUS

    SetTo Args

End Function

Private Function BObject_Name() As String
End Function

Private Function BObject_Ptr() As Long
End Function

Private Sub BObject_Tidyup()

    mCount = 0
    ReDim mEntry(0)
    mIndex = 1              ' // effectively invalid...

End Sub

Private Function BObject_Type() As String

'    BObject_Type = B_PACKED_DATA

End Function

Private Sub Class_Initialize()

    mItemSplit = "#?"
    mValSplit = "::"

End Sub

Public Function SetTo(ByVal PackedString As String, Optional ByVal ItemDelimiter As String = "#?", Optional ByVal ValueDelimiter As String = "::") As Boolean
Dim a() As String
Dim i As Long

    On Error Resume Next

    BObject_Tidyup

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
    SetTo = True

End Function

Private Sub uAdd(ByVal PackedEntry As String)

    On Error Resume Next

'#If GNTP_TEST = 1 Then
'    Form1.Output "TPackedData.uAdd(): '" & PackedEntry & "'"
'#End If

    If mValSplit = "" Then
        ' /* entries are not pairs, so just add as a named entry with no value */
        Add PackedEntry, ""
        Exit Sub

    End If

Dim i As Long

    i = InStr(PackedEntry, mValSplit)
    If i = 0 Then
        ' /* no marker: add as named value with no data */
        Add PackedEntry, ""

    Else
        Add g_SafeLeftStr(PackedEntry, i - 1), g_SafeRightStr(PackedEntry, Len(PackedEntry) - i - (Len(mValSplit) - 1))

    End If

End Sub

Public Sub Add(ByVal Name As String, ByVal Data As String)

'#If GNTP_TEST = 1 Then
'    Form1.Output "TPackedData.Add(): '" & Name & "' '" & Data & "'"
'#End If

    mCount = mCount + 1
    ReDim Preserve mEntry(mCount)
    With mEntry(mCount)
        .Name = Name
        .Value = Data

    End With

End Sub

Public Function AsString() As String
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If mEntry(i).Name <> "" Then
                AsString = AsString & mEntry(i).Name & mValSplit & mEntry(i).Value & IIf(i < mCount, mItemSplit, "")

            End If
        Next i
    End If

End Function

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
Dim i As Long

    If Name = "__hint" Then
        If mCount Then
            For i = 1 To mCount
                sz = sz & mEntry(i).Name & IIf(i < mCount, ",", "")

            Next i
        End If

        ValueOf = sz

    ElseIf Find(Name, sz) Then
        ValueOf = sz

    End If

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

Public Sub Update(ByVal Name As String, ByVal Data As String, Optional ByVal AddIfMissing As Boolean = False)
Dim i As Long

    i = IndexOf(Name)
    If (i = 0) And (AddIfMissing) Then
        Add Name, Data

    ElseIf i <> 0 Then
        mEntry(i).Value = Data

    End If

End Sub

Public Function IndexOf(ByVal Name As String) As Long
Dim i As Long

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

Public Sub Remove(ByVal Index As Long)
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

Public Sub MakeEmpty()

    BObject_Tidyup

End Sub


