VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "snaRSS Log"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Prefs"
      Height          =   555
      Left            =   1800
      TabIndex        =   2
      Top             =   2940
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Site"
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   2940
      Width           =   1455
   End
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

Private Const CLASS_NAME = "w>LOGCABIN"

Dim mhWnd As Long
Dim mConfig As CConfFile3
Dim mSites As BTagList
Dim mPanel As BPrefsPanel
Dim mPassword As String
Dim mAllowDrop As Boolean
Dim mDebugMode As Boolean

Implements KPrefsPanel
Implements KPrefsPage
Implements BWndProcSink
Implements IDropTarget

Private Sub Command1_Click()
Dim sz As String

    sz = InputBox$("Site:")
    If sz = "" Then _
        Exit Sub

    Me.AddNewSite sz

End Sub

Private Sub Command2_Click()

    Me.DoPrefs

End Sub

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

    ' /* pre-start */

    l3OpenLog "%APPDATA%\full phat\logs\sitechecker.log"
    g_Debug "Main()", LEMON_LEVEL_PROC_ENTER

    mDebugMode = (InStr(Command$, "-debug") <> 0)

    If (Not g_IsIDE()) And (Not mDebugMode) Then _
        Me.Hide

    mPassword = create_password(64)
    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me, CLASS_NAME)

    List1.AddItem "window: " & g_HexStr(mhWnd)
    Me.Tag = mPassword
    Me.Caption = App.Title

    ' /* final set up */

    Set mSites = new_BTagList
    uLoadConfig
    uRegister

End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'    If UnloadMode = 0 Then _
'        PostQuitMessage 0
'
'End Sub

Public Sub DebugOutput(ByVal Text As String)

    List1.AddItem Text
    List1.ListIndex = List1.ListCount - 1
    g_Debug Text

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not (mPanel Is Nothing) Then _
        mPanel.Quit

    If ISNULL(mSites) Then _
        Exit Sub

Dim hr As Long
Dim ps As TSite

    hr = snarl_unregister(App.ProductName, mPassword)
    g_Debug "TWindow.Quit(): unregister: " & CStr(hr)

    With mSites
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            ps.Quit

        Loop

        .MakeEmpty

    End With

    Set mSites = Nothing

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Function BWndProcSink_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    On Error Resume Next

    Select Case uMsg

    Case WM_CLOSE
        g_Debug "BWndProcSink.WndProc(): WM_CLOSE"
        Unload Me


    Case snSysMsg()
        Select Case wParam
        Case SNARL_BROADCAST_LAUNCHED
            uRegister

        End Select


    Case snAppMsg()
        Select Case wParam
        Case SNARLAPP_DO_PREFS, SNARLAPP_ACTIVATED
            Me.DebugOutput "_APP_PREFS/_ACTIVATED"
            Me.DoPrefs

        Case SNARLAPP_QUIT_REQUESTED
            g_Debug "BWndProcSink.WndProc(): SNARLAPP_QUIT_REQUESTED"
            Unload Me

        Case SNARLAPP_DO_ABOUT
            g_Debug "BWndProcSink.WndProc(): SNARLAPP_DO_ABOUT"
            snDoRequest "notify?app-sig=" & App.ProductName & _
                        "&icon=" & g_MakePath(App.Path) & "icon.png" & _
                        "&title=SiteChecker " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments & " (Build " & CStr(App.Revision) & ")" & _
                        "&text=" & App.LegalCopyright & vbCrLf & vbCrLf & "Monitors web sites" & _
                        "&uid=about&password=" & mPassword

        End Select

    End Select

End Function

Private Function uRegister() As Boolean
Dim hr As Long

    hr = snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", mPassword, mhWnd, 0, True)
    If hr > 0 Then

        ' /* add classes */

'        snDoRequest "addclass?app-sig=" & App.ProductName & _
'                    "&id=" & CLASS_NORMAL & _
'                    "&name=Normal messages"
'
'        snDoRequest "addclass?app-sig=" & App.ProductName & _
'                    "&id=" & CLASS_PERSONAL & _
'                    "&name=Personal messages"
'
'        snDoRequest "addclass?app-sig=" & App.ProductName & _
'                    "&id=" & CLASS_PRIVATE & _
'                    "&name=Private messages"
'
'        snDoRequest "addclass?app-sig=" & App.ProductName & _
'                    "&id=" & CLASS_CONFIDENTIAL & _
'                    "&name=Confidential messages"

        uRegister = True

    Else
        g_Debug "TWindow.uRegister(): failed (" & CStr(hr) & ")"

    End If

End Function

Private Sub uLoadConfig()
Dim ps As CConfSection

    Set mConfig = New CConfFile3
    With mConfig
        .SetFile g_MakePath(g_GetSystemFolderStr(CSIDL_APPDATA)) & "full phat\" & App.Title & "\" & App.Title & ".conf"
        .Load

        If Not .Exists("general") Then
            Set ps = .NewSection("general")
            .Add ps

        Else
            Set ps = .SectionAt(.IndexOf("general"))

        End If

        ' /* [general] */

        With ps
            .AddIfMissing "RefreshInterval", "60"

        End With


        ' /* load up saved sites */

        .Rewind

        Do While .GetNextSection(ps)
            If ps.Name <> "general" Then _
                uAddSite ps.GetValueWithDefault("url", ""), ps.Name

        Loop

        .Save

    End With

End Sub

Public Sub DoPrefs()
Dim pPage As BPrefsPage
Dim pCtl As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "SiteChecker Preferences"
            .SetWidth 500

            Set pPage = new_BPrefsPage("", Nothing, Me)

            With pPage
                .SetMargin 24
                .Add new_BPrefsControl("banner", "", "Defined Sites")

                Set pm = New CTempMsg
                pm.Add "item-height", 38
                pm.Add "checkboxes", 1&
                Set pCtl = new_BPrefsControl("listbox", "site_list", , , "1", pm)
                pCtl.SizeTo 0, 160
                .Add pCtl

                .Add new_BPrefsControl("label", "", "Drag and drop URLs onto the list above to add them.")

'                Set pCtl = new_BPrefsControl("fancyplusminus", "add_remove", "")
'                .Add pCtl

                .Add new_BPrefsControl("fancytoolbar", "toolbar", "Check Now||Remove", , , , False)

'                .Add new_BPrefsControl("banner", "", "Options")
'                .Add new_BPrefsControl("fancytoggle2", "UseDefaultCallback", "Clicking the notification opens the item?", , IIf(gConfig.UseDefaultCallback, "1", "0"))
'                .Add new_BPrefsControl("fancytoggle2", "SuperSensitive", "Notify when either headline or content changes?", , IIf(gConfig.SuperSensitive, "1", "0"))

                .Add new_BPrefsControl("label", "", "SiteChecker " & CStr(App.Major) & "." & CStr(App.Minor) & " (Build " & CStr(App.Revision) & ") " & App.LegalCopyright, , , , False)
                .Add new_BPrefsControl("label", "", "http://www.fullphat.net", , , , False)
                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "quit_app", "Quit SiteChecker")

            End With

            .AddPage pPage
            .Go
            g_SetWindowIconToAppResourceIcon .hwnd

            If .Find("site_list", pCtl) Then _
                RegisterDragDrop pCtl.hwnd, Me

        End With
    End If

    g_ShowWindow mPanel.hwnd, True, True
    SetForegroundWindow mPanel.hwnd

End Sub

Private Sub IDropTarget_DragEnter(ByVal pDataObject As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)
Dim pDrop As CDropContent
Dim sz As String

    Set pDrop = New CDropContent
    If pDrop.SetTo(pDataObject) Then
        mAllowDrop = pDrop.HasFormat("UniformResourceLocator")

        If mDebugMode Then
            With pDrop
                .Rewind
                Me.DebugOutput "--drop content--"
                Do While .GetNextFormat(sz)
                    Me.DebugOutput sz

                Loop
                Me.DebugOutput "--"

            End With
        End If
    End If

    If mAllowDrop Then
        pdwEffect = DROPEFFECT_COPY

    Else
        pdwEffect = DROPEFFECT_NONE

    End If

End Sub

Private Sub IDropTarget_DragLeave()

End Sub

Private Sub IDropTarget_DragOver(ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)

    If mAllowDrop Then
        pdwEffect = DROPEFFECT_COPY

    Else
        pdwEffect = DROPEFFECT_NONE

    End If

End Sub

Private Sub IDropTarget_Drop(ByVal pDataObject As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)
Dim pDrop As CDropContent
Dim pItem As CDropItem

    Set pDrop = New CDropContent
    If pDrop.SetTo(pDataObject) Then
        If pDrop.GetData("UniformResourceLocator", pItem) Then _
            Me.AddNewSite pItem.GetAsString(False)

    End If

    mAllowDrop = False

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pList As BControl
Dim ps As TSite
Dim sz As String
Dim i As Long

    If mPanel.Find("site_list", pList) Then _
        i = Val(pList.GetValue)

    Select Case Control.GetName

    Case "site_list"
        prefskit_SafeEnable mPanel, "toolbar", (Val(Value) <> 0)

    Case "toolbar"

        If i = 0 Then _
            Exit Sub

        Select Case Value

        Case "1"
            Set ps = mSites.TagAt(i)
            ps.Check

        Case "3"
            If i > 0 Then
                Debug.Print mSites.TagAt(i).Name & " / " & mSites.TagAt(i).Value
                mConfig.RemoveSection mConfig.IndexOf(mSites.TagAt(i).Name)
                mSites.Remove i
                uUpdateConfig
                Me.UpdateList

            End If

        End Select


'    Case "add_remove"
'        If Value = "+" Then
''            If g_IsPressed(VK_CONTROL) Then
''                theAddFeedPanel_AddFeed Clipboard.GetText()
''
''            Else
''                Set theAddFeedPanel = New TAddFeedPanel
''                theAddFeedPanel.Go mpanel.hWnd
''
''            End If
'
'        ElseIf (Value = "-") And (i > 0) Then
'
'        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    If Control.GetName = "quit_app" Then _
        Unload Me

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
Dim pSite As TSite
Dim i As Long

    Select Case Control.GetName()

    Case "site_list"
        Select Case Notification

        Case "checked"
            i = Val(Control.GetValue())
            If i = 0 Then _
                Exit Sub

            Set pSite = mSites.TagAt(i)
            i = Val(prefskit_GetItem(Control, "checked", i))
            pSite.SetEnabled (i = 1)

        End Select

    End Select

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

    RevokeDragDrop mPanel.hwnd
    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()

    Me.UpdateList

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub UpdateList()

    If (mPanel Is Nothing) Then _
        Exit Sub

Dim pList As BControl

    If Not mPanel.Find("site_list", pList) Then _
        Exit Sub

Dim szCurrent As String

    szCurrent = pList.GetValue
    If szCurrent = "0" Then _
        szCurrent = "1"

Dim pSite As TSite
Dim sz As String

    With mSites
        .Rewind
        Do While .GetNextTag(pSite) = B_OK
            sz = sz & g_FormattedMidStr(pSite.URL, 52) & "#?0#?" & _
                      IIf(pSite.IsChecking(), "Checking...", _
                          "Last check: " & pSite.LastCheck() & " Last seen: " & pSite.LastSeen()) & "|"

        Loop

    End With

'    Debug.Print "**"
'    Debug.Print sz
'    Debug.Print "**"

    pList.SetText g_SafeLeftStr(sz, Len(sz) - 1)
    pList.SetValue szCurrent

Dim i As Long

    With mSites
        .Rewind
        Do While .GetNextTag(pSite) = B_OK
            i = i + 1
            prefskit_SetItem pList, i, "checked", IIf(pSite.IsEnabled, 1&, 0&)
            prefskit_SetItem pList, i, "image-file", g_MakePath(App.Path) & IIf(pSite.IsChecking, "wait.png", IIf(pSite.LastCheckWasGood, "ok.png", "failed.png"))
'            prefskit_SetItem pList, i, "greyscale", IIf(pSite.IsEnabled, 0&, 1&)

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

Private Sub uUpdateConfig()

    If Not (mConfig Is Nothing) Then _
        mConfig.Save

End Sub

'Private Sub theAddFeedPanel_AddFeed(ByVal URL As String)
'
'    If URL <> "" Then
'        If mFeedRoster.AddFeed(URL, "", g_CreateGUID()) Then
'            uUpdateConfig
'            uUpdateList
'            prefskit_SetValue mpanel, "site_list", prefskit_DoCmd(mpanel, "site_list", B_COUNT_ITEMS)
'
'        End If
'    End If
'
'End Sub

Public Sub AddNewSite(ByVal URL As String)
Dim ps As TSite

    Set ps = New TSite
    mSites.Add ps
    ps.Init URL, ""

Dim pc As CConfSection

    Set pc = New CConfSection
    pc.SetName ps.Guid
    pc.Add "url", ps.URL
    mConfig.Add pc
    mConfig.Save

    Me.DebugOutput "Added '" & ps.URL & "'"
    Me.DebugOutput "guid=" & ps.Guid

    Me.UpdateList

End Sub

Private Sub uAddSite(ByVal URL As String, ByVal Guid As String)

    If (URL = "") Or (Guid = "") Then _
        Exit Sub

    Me.DebugOutput "Got '" & URL & "'"
    Me.DebugOutput "guid=" & Guid

Dim ps As TSite

    Set ps = New TSite
    mSites.Add ps
    ps.Init URL, Guid

End Sub

Private Function uGotMiscResource() As Boolean

    On Error Resume Next

Dim i As Long

    Err.Clear
    i = processor_count()
    uGotMiscResource = (Err.Number = 0)

End Function


