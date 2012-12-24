VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Classes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mApp As SnarlApp

Private Type T_CLASS
    Name As String
    Description As String
    Enabled As Boolean
    Title As String
    Text As String
    Icon As String
    Callback As String
'    Priority As Long
    Duration As Long
    Sound As String

End Type

Dim mClass() As T_CLASS
Dim mCount As Long

Public Sub Add(ByVal Name As String, ByVal Description As String, Optional ByVal Enabled As Boolean = True, Optional ByVal DefaultTitle As String, Optional ByVal DefaultText As String, Optional ByVal DefaultIcon As String, Optional DefaultCallback As String, Optional ByVal DefaultDuration As Long = -1, Optional ByVal DefaultSound As String)

    mCount = mCount + 1
    ReDim Preserve mClass(mCount)
    With mClass(mCount)
        .Callback = DefaultCallback
        .Description = Description
        .Enabled = Enabled
        .Icon = DefaultIcon
        .Name = Name
        .Text = DefaultText
        .Title = DefaultTitle
        .Duration = DefaultDuration
'        .Priority = DefaultPriority
        .Sound = DefaultSound

    End With

    If NOTNULL(mApp) Then _
        mApp.bDoRequest Me.bAsRequest(mCount, mApp.Signature, mApp.bPassword)

End Sub

Public Function Count() As Long

    Count = mCount

End Function

Friend Function bAsRequest(ByVal Index As Long, ByVal AppSignature As String, ByVal Password As String) As String

    If (Index > 0) And (Index <= mCount) Then _
        bAsRequest = "addclass?app-sig=" & AppSignature & "&id=" & mClass(Index).Name & "&name=" & mClass(Index).Description & _
                     "&enabled=" & IIf(mClass(Index).Enabled, "1", "0") & "&callback=" & mClass(Index).Callback & _
                     "&icon=" & mClass(Index).Icon & "&title=" & mClass(Index).Title & "&text=" & mClass(Index).Text & _
                     IIf(Password <> "", "&password=" & Password, "") & _
                     IIf(mClass(Index).Duration <> -1, "&duration=" & CStr(mClass(Index).Duration), "") & _
                     IIf(mClass(Index).Sound <> "", "&sound=" & mClass(Index).Sound, "")

'                     IIf(mClass(Index).Priority <> 0, "&priority=" & CStr(mClass(Index).Priority), "") & _

End Function

'Friend Sub bAddToApp(ByVal Token As Long, ByVal Password As String)
'Static i As Long
'Dim r As Long
'
'    If mCount = 0 Then _
'        Exit Sub
'
'    For i = 1 To mCount
'        With mClass(i)
'
'            r = snDoRequest("addclass?token=" & CStr(Token) & "&id=" & .Name & "&name=" & .Description & _
'                            "&enabled=" & IIf(.Enabled, "1", "0") & "&callback=" & .Callback & "&icon=" & .Icon & _
'                            "&title=" & .Title & "&text=" & .Text)
'
'            MsgBox r
'
''        .DefaultSound = Data.ValueOf("sound")
''        .DefaultTimeout = Val(Data.ValueOf("duration"))
'
'        End With
'    Next i
'
'End Sub

Public Sub MakeEmpty()

    ReDim mClass(0)
    mCount = 0

End Sub

Friend Function bHasClass(ByVal Name As String) As Boolean
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If mClass(i).Name = Name Then
                bHasClass = True
                Exit Function

            End If
        Next i
    End If

End Function

Public Sub Remove(ByVal Name As String)

    uRemove uIndexOf(Name)

End Sub

Friend Sub bSet(ByRef App As SnarlApp)

    Set mApp = App

End Sub

Friend Sub bUnSet()

    Set mApp = Nothing

End Sub

Private Sub uRemove(ByVal Index As Long)

    If (Index < 1) Or (Index > mCount) Then _
        Exit Sub

    If NOTNULL(mApp) Then _
        mApp.bRemoveClass mClass(Index).Name

Static j As Long

    If Index < mCount Then
        For j = Index To mCount - 1
            LSet mClass(j) = mClass(j + 1)
    
        Next j
    End If
    
    mCount = mCount - 1
    ReDim Preserve mClass(mCount)

End Sub

Private Function uIndexOf(ByVal Name As String) As Long
Static i As Long

    If mCount Then
        For i = 1 To mCount
            If mClass(i).Name = Name Then
                uIndexOf = i
                Exit Function

            End If
        Next i
    End If

End Function
