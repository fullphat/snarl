VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mGuid As String         ' // used as Snarl Class identifier
Dim mTitle As String
Dim mClass As String

Implements BTagItem

Private Function BTagItem_Name() As String

End Function

Private Function BTagItem_Value() As String

End Function

Public Function SetTo(ByVal Guid As String, ByVal Title As String, ByVal Class As String) As Boolean

    If Guid = "" Then _
        Guid = g_CreateGUID()

    mGuid = Guid
    mTitle = Title
    mClass = Class

    SetTo = True

End Function

Public Function Guid() As String

    Guid = mGuid

End Function

Public Function Title() As String

    Title = mTitle

End Function

Public Function Class() As String

    Class = mClass

End Function

Public Function Detail() As String

    Detail = IIf(mTitle = "*", "Any title", "Title is '" & mTitle & "'") & " and " & IIf(mClass = "*", "any class", "class is '" & mClass & "'")

End Function
