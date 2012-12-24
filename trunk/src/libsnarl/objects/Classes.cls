VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotificationClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_CLASS
    Name As String
    Description As String
    Enabled As Boolean
    Title As String
    Text As String
    Icon As String
    Callback As String

End Type

Dim mClass() As T_CLASS
Dim mCount As Long

Public Sub Add(ByVal Name As String, Optional ByVal Description As String, Optional ByVal Enabled As Boolean = True, Optional ByVal DefaultTitle As String, Optional ByVal DefaultText As String, Optional ByVal DefaultIcon As String, Optional Callback As String)

    mCount = mCount + 1
    ReDim Preserve mClass(mCount)
    With mClass(mCount)
        .Callback = Callback
        .Description = Description
        .Enabled = Enabled
        .Icon = DefaultIcon
        .Name = Name
        .Text = DefaultText
        .Title = DefaultTitle

    End With

End Sub

Public Function Count() As Long

    Count = mCount

End Function

Public Function AsRequest(ByVal Index As Long, ByVal AppToken As Long, ByVal Password As String) As String

    If (Index > 0) And (Index <= mCount) Then _
        AsRequest = "addclass?token=" & CStr(AppToken) & "&id=" & mClass(Index).Name & "&name=" & mClass(Index).Description & _
                    "&enabled=" & IIf(mClass(Index).Enabled, "1", "0") & "&callback=" & mClass(Index).Callback & _
                    "&icon=" & mClass(Index).Icon & "&title=" & mClass(Index).Title & "&text=" & mClass(Index).Text & _
                    IIf(Password <> "", "&password=" & Password, "")

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
