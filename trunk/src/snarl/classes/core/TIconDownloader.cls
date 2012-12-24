VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TIconDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mGuid As String
Dim mPath As String
Dim WithEvents theRequest As BWebRequest

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mGuid

End Function

Private Function BTagItem_Value() As String
End Function

Public Sub Go(ByVal URL As String, ByVal Path As String)

    mGuid = g_CreateGUID()
    Set theRequest = New BWebRequest
    theRequest.GetFile URL, Path
    g_Debug "TIconDownloader.Go(): source='" & URL & "' dest='" & Path & "' guid=" & mGuid

End Sub

Private Sub theRequest_Completed(ByVal WasSuccessful As Boolean)

    g_Debug "TIconDownloader.Completed(): guid=" & mGuid & " success=" & WasSuccessful
    g_AppRoster.RemoveIconDownloader mGuid

End Sub
