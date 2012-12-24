VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theFeed As BRSSFeed
Attribute theFeed.VB_VarHelpID = -1
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Dim mURL As String
Dim mLastAttempt As Date
Dim mSuccess As Boolean
Dim mEnabled As Boolean
Dim mFirstRun As Boolean
Dim mRefreshing As Boolean
Dim mGuid As String

Dim mTitle As String

Dim mHistory As BPackedData

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Implements MObject

Private Property Get MObject_Type() As String
End Property

Public Function Init(ByVal URL As String, ByVal Guid As String, ByVal Title As String, Optional ByVal Enabled As Boolean = True) As Boolean

    mURL = URL
    mEnabled = Enabled
    mGuid = Guid
    mTitle = Title

    Set mHistory = New BPackedData

'    mFirstRun = True

    Set theFeed = New BRSSFeed
    Init = Refresh()

Const OFFSET = 5

    ' /* on success, start the timer with a +/- N second offset */

Dim i As Long

    If Init Then
        Randomize Timer
        i = (gConfig.RefreshInterval + (CInt(Rnd * (OFFSET * 2)) - OFFSET)) * 1000
        frmMain.Add g_SafeRightStr(mURL, 80, True)
        frmMain.Add "  feed interval timer is " & CStr(i \ 1000) & " secs"
        Set theTimer = new_BTimer(i)

    Else
        g_Debug "TFeed.Init(): error getting feed", LEMON_LEVEL_CRITICAL

    End If

End Function

Public Sub Quit()

    g_Debug "TFeed.Quit()", LEMON_LEVEL_PROC

    Set theTimer = Nothing

    If Not (theFeed Is Nothing) Then _
        Set theFeed = Nothing

End Sub

'Public Function WasSuccessful() As Boolean
'
'    WasSuccessful = mSuccess
'
'End Function

Private Function uFixStr(ByVal str As String) As String
Dim i As Integer
Dim c As Integer
Dim sz As String
Dim x As Integer
Dim a As String

    ' /* replace consecutive whitespace characters (TAB, CR, LF) with a single space */

    c = Len(str)
    If c Then
        For i = 1 To c
'            g_Debug Mid$(str, i, 1) & " (" & Asc(Mid$(str, i, 1)) & ")"

            Select Case Asc(Mid$(str, i, 1))
            Case 9, 10, 13
                ' /* whitespace character so increment count */
                x = x + 1

            Case Else
                ' /* non-whitespace character so add a single space */
                If x > 0 Then
                    ' /* but _only_ if the previous character added _wasn't_ a space! */
                    If a <> " " Then _
                        sz = sz & " "

                    x = 0

                End If

                a = Mid$(str, i, 1)
                sz = sz & a

            End Select

        Next i

        uFixStr = sz

    End If

End Function

Public Function IsValid() As Boolean

    If Not (theFeed Is Nothing) Then _
        IsValid = theFeed.IsValid


End Function

Private Sub theFeed_Completed(ByVal WasSuccessful As Boolean)
Dim pList As BControl

    On Error Resume Next

    g_Debug "theFeed.Completed()", LEMON_LEVEL_PROC_ENTER

    mLastAttempt = Now()
    mSuccess = WasSuccessful
    mRefreshing = False

'    Debug.Print "imgurl: " & theFeed.Channel.ImageUrl
'    Debug.Print "imglink: " & theFeed.Channel.ImageLink

    ' /* will need to update the feed list */

    frmMain.UpdateFeedList

'Static szLastEntry As String
Static szEntry As String

    frmMain.Add g_SafeRightStr(mURL, 80, True)

    If Not theFeed.IsValid Then
        frmMain.Add "  failed: feed content is not valid"

    ElseIf theFeed.CountItems < 1 Then
        frmMain.Add "  failed: feed content is valid but has no entries"

    Else
        ' /* record the title here (each time in case it's changed) */

        mTitle = Replace$(theFeed.Channel.Title, "|", ":")

        frmMain.Add "  title: " & mTitle
        frmMain.Add "  timestamp: " & theFeed.ItemAt(1).PubDate

        With theFeed.ItemAt(1)

            szEntry = .Title & " " & .Description

            If Not mHistory.Exists(szEntry) Then
                ' /* not already been displayed */

                frmMain.Add "  new item found in feed"

'                szLastEntry = szEntry
                If Not mFirstRun Then
                    ShowHeadline

                Else
                    frmMain.Add "  (first run, not showing headline)"
                    mFirstRun = False

                End If


                ' /* add to the history */

                mHistory.Add szEntry, CStr(Now)

            Else
                frmMain.Add "  already displayed " & szEntry

            End If


'            If gConfig.SuperSensitive Then
'                ' /* notify a headline change when either description or title changes */
'                szEntry = .Description & .Title
'
'            Else
'                ' /* notify a change when only the title changes */
'                szEntry = .Title
'
'            End If
'
'            If szEntry <> szLastEntry Then
'                ' /* top story has changed */
'                frmMain.Add "first entry has changed"
'                szLastEntry = szEntry
'                If Not mFirstRun Then
'                    ShowHeadline
'
'                Else
'                    g_Debug "first run, not showing headline"
'                    mFirstRun = False
'
'                End If
'            End If
        End With
    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub theTimer_Pulse()

    g_Debug "timer_pulse: " & mURL & " (" & Title() & ")"
    Refresh

End Sub

Public Function URL() As String

    URL = mURL

End Function

Private Function uEntriesStr() As String

    If (theFeed Is Nothing) Then _
        Exit Function

    Select Case theFeed.CountItems
    Case 0
        uEntriesStr = "No entries"

    Case 1
        uEntriesStr = "1 entry"

    Case Else
        uEntriesStr = CStr(theFeed.CountItems) & " entries"

    End Select

End Function

Public Sub SetEnabled(ByVal Enabled As Boolean)

    mEnabled = Enabled

    If Not (theTimer Is Nothing) Then _
        theTimer.SetEnabled Enabled

End Sub

Public Function IsEnabled() As Boolean

    IsEnabled = mEnabled

End Function

Public Sub ShowHeadline()
Static i As Long

    g_Debug "TFeed.ShowHeadline()", LEMON_LEVEL_PROC_ENTER

    If NOTNULL(theFeed) Then
        If theFeed.CountItems > 0 Then

            With theFeed.ItemAt(1)
                snarl_notify App.ProductName, mGuid, "", frmMain.Password, _
                             .Title, _
                             IIf(.Description <> "", uTidyup(.Description), "<no entry>") & vbCrLf & vbCrLf & theFeed.Channel.Title, _
                             theFeed.Channel.ImageUrl, _
                             , , _
                             IIf(gConfig.UseDefaultCallback, .Link, ""), , _
                             "action=Open Item," & .Link & "&action=Settings...,!snarl manage " & App.ProductName

'                i = sn41EZNotify(gToken, mGuid, _
'                                 .Title, _
'                                 IIf(.Description <> "", .Description, "<no entry>") & vbCrLf & vbCrLf & theFeed.Channel.Title, _
'                                 , _
'                                 theFeed.Channel.ImageUrl, , _
'                                 IIf(gConfig.UseDefaultCallback, .Link, ""), , _
'                                 IIf(Not gConfig.UseDefaultCallback, "action::Open Item," & .Link, ""))
'
'                g_Debug "sz41EZNotify() returned " & CStr(i) & " (lasterror=" & sn41GetLastError() & ")"

            End With

        Else
            g_Debug "feed has no entries", LEMON_LEVEL_CRITICAL

        End If

    Else
        g_Debug "feed isn't valid", LEMON_LEVEL_CRITICAL

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Sub ShowSummary()

    g_Debug "TFeed.ShowSummary()", LEMON_LEVEL_PROC_ENTER

Static sz As String
Static i As Long

    If NOTNULL(theFeed) Then
        If theFeed.CountItems > 0 Then

            sz = ""

            With theFeed
                For i = 1 To MIN(8, .CountItems)
                    sz = sz & "• " & .ItemAt(i).Title & IIf(i < .CountItems, vbCrLf, "")

                Next i

                snarl_notify App.ProductName, mGuid, "", frmMain.Password, _
                             .Channel.Title, _
                             sz, _
                             .Channel.ImageUrl, , , _
                             IIf(gConfig.UseDefaultCallback, .Channel.Link, ""), , _
                             IIf(Not gConfig.UseDefaultCallback, "action::Open Feed," & .Channel.Link, "")


'                i = sn41EZNotify(gToken, mGuid, _
                                 .Channel.Title, _
                                 sz, _
                                 , _
                                 .Channel.ImageUrl, , _
                                 IIf(gConfig.UseDefaultCallback, .Channel.Link, ""), , _
                                 IIf(Not gConfig.UseDefaultCallback, "action::Open Feed," & .Channel.Link, ""))

            End With

        Else
            g_Debug "feed has no entries", LEMON_LEVEL_CRITICAL

        End If

    Else
        g_Debug "bad feed or not registered with Snarl", LEMON_LEVEL_CRITICAL

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function Refresh() As Boolean

    g_Debug "TFeed.Refresh()", LEMON_LEVEL_PROC_ENTER

    If mRefreshing Then
        g_Debug "busy...", LEMON_LEVEL_WARNING

    Else
        g_Debug "asking '" & mURL & "' to refresh..."
        Refresh = (theFeed.GetContent(mURL) = B_OK)
        mRefreshing = True

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Sub FeedInfo()

    On Error Resume Next

    g_Debug "TFeed.FeedInfo()", LEMON_LEVEL_PROC_ENTER

Static sz As String
Static i As Long

    If NOTNULL(theFeed) Then
        If Not theFeed.IsValid Then
            snarl_notify App.ProductName, mGuid, mGuid, frmMain.Password, _
                         "Could not retrieve feed", _
                         "There may be a problem with your Internet connection, or the server providing the feed may be off-line - try accessing the feed from a browser", _
                         g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "failed.png"

'            i = sn41EZNotify(gToken, "", _
                             "Could not retrieve feed", _
                             "There may be a problem with your Internet connection, or the server providing the feed may be off-line - try accessing the feed from a browser", , _
                             g_MakePath(App.Path) & "failed.png")

        ElseIf (theFeed.Channel Is Nothing) Then
            snarl_notify App.ProductName, mGuid, mGuid, frmMain.Password, _
                         "Invalid RSS Feed", _
                         "Invalid RSS Feed", _
                         g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "failed.png"

'            i = sn41EZNotify(gToken, "", _
                             "Invalid RSS Feed", _
                             "Data was returned but it does not appear to be a valid RSS feed", , _
                             g_MakePath(App.Path) & "failed.png")

        Else
            With theFeed

                g_Debug .RawXML

                sz = IIf(.Channel.Description <> "", .Channel.Description & vbCrLf, "") & _
                     IIf(.Channel.Copyright <> "", .Channel.Copyright & vbCrLf, "") & _
                     IIf(.FeedVersion <> "", "Version " & .FeedVersion, "")

                If sz <> "" Then _
                    snarl_notify App.ProductName, mGuid, mGuid, frmMain.Password, _
                                 .Channel.Title, _
                                 sz, _
                                 .Channel.ImageUrl, , , _
                                 .Channel.Link

'                    i = sn41EZNotify(gToken, "", _
                                     .Channel.Title, _
                                     sz, _
                                     , _
                                     .Channel.ImageUrl, , _
                                     .Channel.Link)

            End With

        End If

    Else
        g_Debug "bad feed or not registered with Snarl", LEMON_LEVEL_CRITICAL

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function TitleOrURL() As String

    TitleOrURL = mTitle
    If TitleOrURL = "" Then _
        TitleOrURL = mURL

End Function

Public Function Title() As String

    Title = mTitle

End Function

Public Function Status() As String

    If NOTNULL(theFeed) Then
        If theFeed.IsValid Then
            If mRefreshing Then
                Status = "Refreshing..."

            Else
                Status = uEntriesStr() & ".  Last refreshed: " & mLastAttempt

            End If

        Else
            Status = "Failed.  Last attempt: " & mLastAttempt

        End If

    Else
        Status = "Error."

    End If

End Function

Public Function Guid() As String

    Guid = mGuid

End Function

Public Function CountEntries() As Long

    If NOTNULL(theFeed) Then _
        CountEntries = theFeed.CountItems

End Function

Private Function uTidyup(ByVal html As String) As String
Dim pDoc As HTMLDocument

    On Error Resume Next

    Err.Clear
    Set pDoc = New HTMLDocument
    If (Err.Number <> 0) Or (ISNULL(pDoc)) Then
        uTidyup = html

    Else
        pDoc.body.innerHTML = html
        uTidyup = pDoc.body.innerText

    End If

Dim i As Integer

    Select Case gConfig.HeadlineLength
    Case 2
        ' /* short */
        i = 128

    Case 3
        ' /* long */
        i = 512

    Case Else
        ' /* medium */
        i = 256

    End Select

    uTidyup = trim(g_SafeLeftStr(uTidyup, i, True))
    If g_BeginsWith(uTidyup, vbCrLf) Then _
        uTidyup = g_SafeRightStr(uTidyup, Len(uTidyup) - 2)

End Function
