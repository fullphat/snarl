VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mhWnd As Long
Dim mTitle As String
Dim mClass As String

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mhWnd)

End Function

Private Function BTagItem_Value() As String
End Function

Public Sub Init(ByVal hWnd As Long, ByVal Class As String, ByVal Title As String)

    mhWnd = hWnd
    mClass = Class
    mTitle = Title

End Sub

Public Function Class() As String

    Class = mClass

End Function

Public Function Title() As String

    Title = mTitle

End Function

Public Function Handle() As Long

    Handle = mhWnd

End Function
