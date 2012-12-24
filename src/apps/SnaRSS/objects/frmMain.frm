VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "snaRSS Log"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME = "w>snarss"

Public theFeedRoster As TFeedRoster
Public Password As String
Dim mDebugMode As Boolean
Dim mPanel As BPrefsPanel
Dim mhWnd As Long

Implements KPrefsPanel
Implements KPrefsPage
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    On Error Resume Next

    Select Case uMsg


    Case WM_CLOSE
        Unload Me


    Case snSysMsg()
        Select Case wParam
        Case SNARL_BROADCAST_LAUNCHED
            Me.Add "Snarl launched (V" & CStr(lParam) & ")"
            If LoWord(lParam) >= 41 Then _
                uRegisterWithSnarl

        Case SNARL_BROADCAST_QUIT
            Me.Add "Snarl quit"

        End Select


    Case snAppMsg()
        Select Case wParam
        Case SNARLAPP_DO_PREFS, SNARLAPP_ACTIVATED
            Me.Add "_APP_PREFS"
            uDoPrefs

        Case SNARLAPP_DO_ABOUT
'            sn41EZNotify gToken, "", _
                         "snaRSS " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments & " (Build " & CStr(App.Revision) & ")", _
                         App.LegalCopyright & vbCrLf & vbCrLf & "Tracks RSS feeds and displays a Snarl notification whenever the headline changes.", , _
                         g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png"

        End Select

    End Select

End Function

Private Sub Form_Load()
Dim hWndExisting As Long

    hWndExisting = FindWindow(CLASS_NAME, CLASS_NAME)

    If InStr(Command$, "-quit") Then
        ' /* quit any existing instance (but don't run this one) */
        If IsWindow(hWndExisting) <> 0 Then _
            SendMessage hWndExisting, WM_CLOSE, 0, ByVal 0&

        Unload Me
        Exit Sub

    ElseIf hWndExisting <> 0 Then
        Unload Me
        Exit Sub

    ElseIf Not uGotMiscResource() Then
        MsgBox "misc.resource missing or damaged" & vbCrLf & vbCrLf & "This can happen if melon is uninstalled - try reinstalling melon in the first instance", vbCritical Or vbOKOnly, App.Title
        Unload Me
        Exit Sub

    End If

    ' /* startup */

    l3OpenLog "%APPDATA%\snaRSS.log", True
    g_Debug "Main()", LEMON_LEVEL_PROC_ENTER

    mDebugMode = (InStr(Command$, "-debug") <> 0)
    If Not mDebugMode Then _
        Me.Hide

    ' /* create window */

    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me, CLASS_NAME)
    List1.AddItem "window: " & g_HexStr(hWnd)


Dim pFeedInfo As CConfSection
Dim sz As String
Dim i As Long

    ' /* set up */

    Set theFeedRoster = New TFeedRoster
    Password = create_password()

    ' /* defaults */

    With gConfig
        .RefreshInterval = 60
        .UseDefaultCallback = True
        .SuperSensitive = False
        .HeadlineLength = 2

    End With

    ' /* read config */

    With New CConfFile3

        .SetFile g_MakePath(App.Path) & "snaRSS.conf"

        If .Load Then
            .Rewind

            ' /* general settings */

            i = .IndexOf("settings")
            If i > 0 Then
                With .SectionAt(i)
                    gConfig.RefreshInterval = g_SafeLong(.GetValueWithDefault("RefreshInterval", "60"))
                    gConfig.UseDefaultCallback = (.GetValueWithDefault("UseDefaultCallback", "1") = "1")
                    gConfig.SuperSensitive = (.GetValueWithDefault("SuperSensitive", "0") = "1")
                    gConfig.HeadlineLength = g_SafeLong(.GetValueWithDefault("HeadlineLength", "2"))

                End With

            Else
                g_Debug "TWindow.uStart(): settings section not in config"

            End If

            With gConfig
                ' /* can't be less than 3 seconds as we could end up with a -ve or 0 timeout */
                If .RefreshInterval < 3 Then _
                    .RefreshInterval = 3

            End With

            Me.Add "refresh interval: " & gConfig.RefreshInterval

            ' /* read in stored feeds */

            Do While .GetNextSection(pFeedInfo)
                If pFeedInfo.Name = "feed" Then
                    ' /* R1.0: identified by a guid */
                    sz = pFeedInfo.GetValueWithDefault("guid")
                    If trim(sz) = "" Then
                        pFeedInfo.Update "guid", g_CreateGUID()
                        .Save

                    End If

                    theFeedRoster.AddFeed pFeedInfo.GetValueWithDefault("URL"), pFeedInfo.GetValueWithDefault("title"), pFeedInfo.GetValueWithDefault("guid")

                End If
            Loop
        End If
    End With

'        theFeedRoster.AddFeed "http://newsrss.bbc.co.uk/rss/newsonline_uk_edition/sci/tech/rss.xml", ""
'        theFeedRoster.AddFeed "http://newsrss.bbc.co.uk/rss/newsonline_uk_edition/england/rss.xml", ""
'        theFeedRoster.AddFeed "http://www.theregister.co.uk/headlines.atom", ""

    uRegisterWithSnarl
    g_Debug "startup complete"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    If UnloadMode = 0 Then _
        Unload Me

End Sub

Public Sub Add(ByVal Text As String)

    List1.AddItem Text
    List1.ListIndex = List1.ListCount - 1
    g_Debug Text

End Sub

Private Sub uRegisterWithSnarl()

    On Error Resume Next

    If snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png", Password, mhWnd, 0, True) > 0 Then
        Me.Add "Registered with Snarl"
        theFeedRoster.AddClasses

    Else
        Me.Add "Error registering with Snarl"

    End If

End Sub

Private Sub uDoPrefs()
Dim pPage As BPrefsPage
Dim pCtl As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "snaRSS Preferences"
            .SetWidth 400

            Set pPage = new_BPrefsPage("", Nothing, Me)
            With pPage
                .SetMargin 0
                Set pm = New CTempMsg
                pm.Add "height", 330
                Set pCtl = new_BPrefsControl("tabstrip", "", , , , pm)
                BTabStrip_AddPage pCtl, "Feeds", new_BPrefsPage("feeds", , New TSubPage)
                BTabStrip_AddPage pCtl, "Options", new_BPrefsPage("options", , New TSubPage)
                BTabStrip_AddPage pCtl, "About", new_BPrefsPage("about", , New TSubPage)
                .Add pCtl

                .Add new_BPrefsControl("fancybutton2", "quit_app", "Quit snaRSS")

            End With
            .AddPage pPage




            .Go
            g_SetWindowIconToAppResourceIcon .hWnd
            g_WindowToFront .hWnd, True

'            SetForegroundWindow .hWnd

            UpdateFeedList

        End With
    End If

    g_ShowWindow mPanel.hWnd, True, True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    g_Debug "stopping..."

    If NOTNULL(mPanel) Then _
        mPanel.Quit

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

    UpdateConfig

    If Password <> "" Then _
        snarl_unregister App.ProductName, Password

    Set theFeedRoster = Nothing

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    If Control.GetName = "quit_app" Then _
        Unload Me

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub UpdateFeedList()

    If ISNULL(mPanel) Then _
        Exit Sub

Dim pList As BControl

    If Not mPanel.Find("feed_list", pList) Then _
        Exit Sub

Dim szCurrent As String
Dim pFeed As TFeed
Dim sz As String

    szCurrent = pList.GetValue()
    If szCurrent = "0" Then _
        szCurrent = "1"

    With theFeedRoster
        .Rewind
        Do While .NextFeed(pFeed)
            sz = sz & g_FormattedMidStr(pFeed.TitleOrURL, 52) & "#?0#?" & pFeed.Status() & "|"

        Loop

    End With

    Debug.Print "**"
    Debug.Print sz
    Debug.Print "**"

    pList.SetText g_SafeLeftStr(sz, Len(sz) - 1)
    pList.SetValue szCurrent

Dim i As Long

    With theFeedRoster
        .Rewind
        Do While .NextFeed(pFeed)
            i = i + 1
            prefskit_SetItem pList, i, "checked", IIf(pFeed.IsEnabled, 1&, 0&)

        Loop

    End With

End Sub

'Private Function uSpecialMidStr(ByVal Text As String, ByVal MaxLen As Long) As String
'
'    If Len(Text) < MaxLen Then
'        uSpecialMidStr = Text
'        Exit Function
'
'    End If
'
'Dim i As Long
'
'    i = Fix((MaxLen - 3) / 2)
'    uSpecialMidStr = g_SafeLeftStr(Text, i) & "…" & g_SafeRightStr(Text, i)  '"…"
'
'End Function

Public Sub UpdateConfig()

    If ISNULL(theFeedRoster) Then _
        Exit Sub

Dim pFile As CConfFile3
Dim pItem As CConfSection
Dim pFeed As TFeed

    ' /* update config */

    Set pFile = New CConfFile3
    pFile.SetFile g_MakePath(App.Path) & "snaRSS.conf"

    ' /* [settings] */

    Set pItem = New CConfSection
    With pItem
        .SetName "settings"
        .Add "RefreshInterval", CStr(gConfig.RefreshInterval)
        .Add "UseDefaultCallback", IIf(gConfig.UseDefaultCallback, "1", "0")

    End With

    pFile.Add pItem

    ' /* [feed] sections */

    With theFeedRoster
        .Rewind
        Do While .NextFeed(pFeed)
            Set pItem = New CConfSection
            pItem.SetName "feed"
            pItem.Add "url", pFeed.URL
            pItem.Add "guid", pFeed.Guid
            pItem.Add "title", pFeed.Title
            pFile.Add pItem

        Loop

    End With

    pFile.Save

End Sub

Private Function uGotMiscResource() As Boolean

    On Error Resume Next

Dim i As Long

    Err.Clear
    i = processor_count()
    uGotMiscResource = (Err.Number = 0)

End Function



