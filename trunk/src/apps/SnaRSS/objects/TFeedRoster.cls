VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFeedRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mFeed() As TFeed
Dim mFeeds As Long
Dim mIndex As Long

Implements MRoster

Private Sub Class_Initialize()

    g_Debug "TFeedRoster.TFeedRoster()", LEMON_LEVEL_PROC
    mIndex = 1

End Sub

Private Sub Class_Terminate()
Dim i As Long

    On Error Resume Next

    g_Debug "TFeedRoster.~TFeedRoster()", LEMON_LEVEL_PROC

    If mFeeds Then
        For i = 1 To mFeeds
            mFeed(i).Quit

        Next i
    End If

    mFeeds = 0
    ReDim mFeed(0)

End Sub

Private Function MRoster_CountItems() As Long

    MRoster_CountItems = mFeeds

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long
End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.MObject

    If (Index > 0) And (Index <= mFeeds) Then _
        Set MRoster_ItemAt = mFeed(Index)

End Function

Private Function MRoster_NextItem(Item As melon.MObject) As Boolean

End Function

Private Sub MRoster_Rewind()

    mIndex = 1

End Sub

Public Function AddFeed(ByVal URL As String, ByVal Title As String, ByVal Guid As String) As Boolean
Dim sz() As String
Dim pf As TFeed

    On Error Resume Next

    If URL = "" Then _
        Exit Function

    Set pf = New TFeed
    If pf.Init(URL, Guid, Title) Then
        mFeeds = mFeeds + 1
        ReDim Preserve mFeed(mFeeds)
        Set mFeed(mFeeds) = pf
        g_Debug "TFeedRoster.AddFeed(): added ok"
        snarl_add_class App.ProductName, Guid, IIf(Title = "", URL, Title)
        AddFeed = True

    Else
        g_Debug "TFeedRoster.AddFeed(): RSS initialisation error", LEMON_LEVEL_CRITICAL

    End If

End Function

Public Function CountFeeds() As Long

    CountFeeds = MRoster_CountItems

End Function

Public Function FindFeed(ByVal UID As Long) As Long

    FindFeed = MRoster_FindItem(CStr(UID))

End Function

Public Function FeedAt(ByVal Index As Long) As TFeed

    Set FeedAt = MRoster_ItemAt(Index)

End Function

Public Sub Rewind()

    mIndex = 1

End Sub

Public Function NextFeed(ByRef Feed As TFeed) As Boolean

    If (mIndex > 0) And (mIndex <= mFeeds) Then
        Set Feed = mFeed(mIndex)
        mIndex = mIndex + 1
        NextFeed = True

    End If

End Function

Public Sub Remove(ByVal Index As Long)
Static i As Long

    If (Index < 1) Or (Index > mFeeds) Then _
        Exit Sub

    snarl_rem_class App.ProductName, mFeed(Index).Guid, frmMain.Password

    If Index < mFeeds Then
        For i = Index To (mFeeds - 1)
            Set mFeed(i) = mFeed(i + 1)

        Next i
    End If

    mFeeds = mFeeds - 1
    ReDim Preserve mFeed(mFeeds)

End Sub

Public Sub AddClasses()
Static i As Long

    If mFeeds Then
        For i = 1 To mFeeds
            snarl_add_class App.ProductName, mFeed(i).Guid, mFeed(i).TitleOrURL, , frmMain.Password

        Next i
    End If

End Sub
