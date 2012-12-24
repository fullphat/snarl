VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FolderSnoop Debug Log"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_ICON = 1

Dim theTrayIcon As BNotifyIcon
Dim mPanel As BPrefsPanel
Dim mFolders As BTagList
Dim mToken As Long

Dim WithEvents theAddWatchPanel As TAddWatchPanel
Attribute theAddWatchPanel.VB_VarHelpID = -1

Dim WithEvents Snarl As Snarl
Attribute Snarl.VB_VarHelpID = -1

Private Const FOLDER_CREATED = "foldercreated"
Private Const FOLDER_RENAMED = "folderrenamed"
Private Const FOLDER_DELETED = "folderdeleted"
Private Const FOLDER_UPDATED = "folderupdated"

Private Const FILE_CREATED = "filecreated"
Private Const FILE_RENAMED = "filerenamed"
Private Const FILE_DELETED = "filedeleted"
Private Const FILE_UPDATED = "fileupdated"

Implements KPrefsPage
Implements KPrefsPanel
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static psns As SHNOTIFYSTRUCT
Dim pWatch As TFolderWatch
Dim pmi As OMMenuItem

    Select Case uMsg

    Case WM_USER
        ' /* tray icon */
        Select Case lParam
        Case WM_LBUTTONDBLCLK
            uDoPrefs

        Case WM_RBUTTONUP
            With New OMMenu
                .AddItem .CreateItem("prefs", "Preferences...")
                .AddSeparator
                .AddItem .CreateItem("quit", "Quit")

                Set pmi = .Track(hWnd)
                If Not (pmi Is Nothing) Then
                    Select Case pmi.Name
                    Case "prefs"
                        uDoPrefs

                    Case "quit"
                        Unload Me

                    End Select

                End If

            End With

        End Select


    Case Is >= (WM_USER + 1)

        ' /* a ShellChangeNotify event */

        If mFolders.Find(CStr(uMsg), pWatch) Then
            Debug.Print "foldersnoop: " & g_HexStr(wParam) & " " & g_HexStr(lParam) & " [" & g_SHNotifyStr(lParam) & "]"
            CopyMemory psns, ByVal wParam, Len(psns)
            uShellChangeEvent lParam, g_GetPathFromPIDL(psns.dwItem1), g_GetPathFromPIDL(psns.dwItem2), pWatch
    
        Else
            Debug.Print "error: watch " & CStr(uMsg) & " not found"
    
        End If


    Case WM_CLOSE
        Me.Add "WM_CLOSE"
        Unload Me


    Case WM_SYSCOMMAND
        If (LoWord(wParam) And &HFFF0&) = SC_MINIMIZE Then
            g_ShowWindow hWnd, False
            ReturnValue = 0
            BWndProcSink_WndProc = True

        End If


    Case snAppMsg

        Select Case wParam
        Case SNARLAPP_DO_PREFS
'            frmMain.Add "_APP_PREFS"
            uDoPrefs

        Case SNARLAPP_DO_ABOUT
            If mToken Then _
                snDoRequest "notify?token=" & CStr(mToken) & _
                            "&title=FolderSnoop " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments & " (Build " & CStr(App.Revision) & ")" & _
                            "&text=" & App.LegalCopyright & vbCrLf & vbCrLf & "Watches OCS" & _
                            "&icon=" & g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png"

        End Select


    End Select

End Function

Private Sub Form_Load()
Dim sz As String

    If App.PrevInstance Then
        ' /* we're already running... */
        sz = Me.Caption
        Me.Caption = ""

        ' /* if -quit specified, tell the other instance to quit */
        If InStr(Command, "-quit") <> 0 Then _
            PostMessage FindWindow("ThunderRT6FormDC", sz), WM_CLOSE, 0, ByVal 0&

        ' /* unload either way */
        Unload Me
        Exit Sub

    End If

    Me.Add "starting..."
    Me.Add "  " & App.Title & " " & App.Major & "." & App.Minor & " Build " & App.Revision

    Set mFolders = new_BTagList()
    Set Snarl = get_snarl()

    window_subclass Me.hWnd, Me

    gNextFreeMsg = WM_USER + 1

    Set theTrayIcon = New BNotifyIcon
    With theTrayIcon
        .SetTo Me.hWnd, WM_USER
        If g_IsIDE() Then
            .Add "icon", Me.Icon.Handle, "FolderSnoop"

        Else
            .Add "icon", LoadImage(App.hInstance, 1&, IMAGE_ICON, 16, 16, 0), "FolderSnoop"

        End If

    End With

    Me.Hide

    ' /* register with Snarl, if it's around */

    If is_snarl_running() Then _
        uRegister


    ' /* load config */

Dim pcf As CConfFile3
Dim pcs As CConfSection
Dim pfw As TFolderWatch

    Set pcf = New CConfFile3
    With pcf
        .SetFile g_MakePath(App.Path) & "foldersnoop.conf"
        .Load
        Do While .GetNextSection(pcs)
            If pcs.Name = "watch" Then
                Set pfw = New TFolderWatch
                If pfw.SetTo(pcs.GetValueWithDefault("path"), Val(pcs.GetValueWithDefault("flags")), pcs.GetValueWithDefault("guid"), pcs.GetValueWithDefault("recurse") = "1") Then _
                    mFolders.Add pfw

            End If

        Loop

    End With


'Dim pFolder As TFolderWatch
'
'    Set pFolder = New TFolderWatch
'    pFolder.SetTo "c:\"
'    mFolders.Add pFolder
'
'    Set pFolder = New TFolderWatch
'    pFolder.SetTo "d:\"
'    mFolders.Add pFolder


End Sub

Private Sub Form_Unload(Cancel As Integer)

    window_subclass Me.hWnd, Nothing

    If Not (mPanel Is Nothing) Then _
        mPanel.Quit

    If Not (theTrayIcon Is Nothing) Then
        theTrayIcon.Remove "icon"
        Set theTrayIcon = Nothing

    End If

    snarl_unregister App.ProductName

End Sub

Private Sub Snarl_SnarlLaunched()

    Me.Add "[snarl launched]"
    uRegister

'Dim pc As BControl
'
'    If Not (mPanel Is Nothing) Then
'        If mPanel.Find("snarl_state", pc) Then _
'            pc.SetText "Snarl is running"
'
'    End If

End Sub

Private Sub Snarl_SnarlQuit()

    Me.Add "[snarl quit]"

'Dim pc As BControl
'
'    If Not (mPanel Is Nothing) Then
'        If mPanel.Find("snarl_state", pc) Then _
'            pc.SetText "Snarl is running"
'
'    End If
'
'    uStartWatching

End Sub

Public Sub Add(ByVal Text As String)

    List1.AddItem Text
    List1.ListIndex = List1.ListCount - 1
    g_Debug Text

End Sub

Private Sub uRegister()
Dim hr As Long

    mToken = 0

    hr = snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", , Me.hWnd, , SNARLAPP_HAS_ABOUT Or SNARLAPP_HAS_PREFS Or SNARLAPP_IS_WINDOWLESS)
    If hr > 0 Then
        Add "snarl token: " & CStr(hr)

'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & FOLDER_CREATED & "&name=Folder created"
'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & FOLDER_RENAMED & "&name=Folder renamed"
'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & FOLDER_DELETED & "&name=Folder deleted"
'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & FOLDER_UPDATED & "&name=Folder updated"
'
'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & FILE_CREATED & "&name=File created"
'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & FILE_RENAMED & "&name=File renamed"
'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & FILE_DELETED & "&name=File deleted"
'        snDoRequest "addclass?token=" & CStr(hr) & "&id=" & FILE_UPDATED & "&name=File updated"

        mToken = hr

    Else
        Add "couldn't register with Snarl (" & CStr(Abs(hr)) & ")"

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
            .SetTitle "FolderSnoop Preferences"
            .SetWidth 400

            Set pPage = new_BPrefsPage("", Nothing, Me)

            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("banner", "", "Folder Watch List")

                Set pm = New CTempMsg
                pm.Add "item-height", 38
'                pm.Add "checkboxes", 1&
                Set pCtl = new_BPrefsControl("listbox", "watch_list", , , "1", pm)
                pCtl.SizeTo 0, 160
                .Add pCtl

                Set pCtl = new_BPrefsControl("fancyplusminus", "add_remove", "")
                .Add pCtl

'                .Add new_BPrefsSeparator


'                .Add new_BPrefsControl("fancytoolbar", "feed_toolbar", "Show Headline|Show Summary|Refresh|Feed Information", , , , False)
'
'                .Add new_BPrefsControl("fancytoggle2", "UseDefaultCallback", "Clicking the notification opens the item?", , IIf(gConfig.UseDefaultCallback, "1", "0"))

'                .Add new_BPrefsControl("banner", "", "Status Changes")
'                .Add new_BPrefsControl("label", "", "Include changes from the following groups:")
''                .Add new_BPrefsControl("fancytoggle2", "UseDefaultCallback", "Clicking the notification opens the item?", , IIf(gConfig.UseDefaultCallback, "1", "0"))

'                .Add new_BPrefsControl("label", "", "FolderSnoop will alert you to incoming IM conversations and phone calls, as well as contact status changes.  Due to limitations of the Communicator API, only certain information is available.")
'                .Add new_BPrefsControl("label", "snarl_state", IIf(mHasSnarl, "Snarl is running", "Snarl is not running"))
'
'                .Add new_BPrefsControl("banner", "", "Options")

                .Add new_BPrefsControl("banner", "", "Debug")
                .Add new_BPrefsControl("fancybutton2", "ShowHideDebug", "Show/Hide Debug Window")

                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "quit_app", "Quit FolderSnoop")
                .Add new_BPrefsControl("label", "", "FolderSnoop " & CStr(App.Major) & "." & CStr(App.Minor) & " (Build " & CStr(App.Revision) & ") " & App.LegalCopyright, , , , False)
'                .Add new_BPrefsControl("label", "", "http://www.fullphat.net", , , , False)

            End With

            .AddPage pPage
            .Go
            g_SetWindowIconToAppResourceIcon .hWnd
            SetForegroundWindow .hWnd

'            uUpdateFeedList

        End With
    End If

    g_ShowWindow mPanel.hWnd, True, True

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pfw As TFolderWatch
Dim i As Long

    Select Case Control.GetName

    Case "add_remove"
        If Value = "+" Then
'            If g_IsPressed(VK_CONTROL) Then
'                theAddFeedPanel_AddFeed Clipboard.GetText()
'
'            Else
                Set theAddWatchPanel = New TAddWatchPanel
                theAddWatchPanel.Go mPanel.hWnd

'            End If

        ElseIf (Value = "-") Then
            i = Val(prefskit_GetValue(mPanel, "watch_list"))
            Set pfw = mFolders.TagAt(i)
            If (pfw Is Nothing) Then _
                Exit Sub

            mFolders.Remove i
            uWriteConfig
            uUpdateList

            snDoRequest "remclass?app-sig=" & App.ProductName & "&id=" & pfw.Guid

            prefskit_SetValue mPanel, "watch_list", CStr(i)

        End If

'    Case "UseDefaultCallback"
'        gConfig.UseDefaultCallback = (Value = "1")
'        uUpdateConfig

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    If Control.GetName = "quit_app" Then
        Unload Me

    ElseIf Control.GetName = "ShowHideDebug" Then
        Me.Visible = Not Me.Visible

    End If

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

    uUpdateList

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub uShellChangeEvent(ByVal EventId As Long, ByVal Path1 As String, ByVal Path2 As String, ByRef Watch As TFolderWatch)
Dim szClass As String
Dim szTitle As String
Dim szText As String
Dim szIcon As String
Dim szUid As String
Dim szReplace As String

    Debug.Print " 1> " & Path1
    Debug.Print " 2> " & Path2

    Select Case EventId

'    Case SHCNE_MEDIAINSERTED, SHCNE_MEDIAREMOVED, SHCNE_DRIVEREMOVED, SHCNE_DRIVEADD, SHCNE_NETSHARE, SHCNE_NETUNSHARE
        ' /* not interested in these */

'Public Const SHCNE_UPDATEIMAGE = &H8000&     '(G) An image in the system image list has changed.
'Public Const SHCNE_DRIVEADDGUI = &H10000     '(G) A drive has been added and the shell should create a new window for the drive.
'Public Const SHCNE_FREESPACE = &H40000       '(G) The amount of free space on a drive has changed.
'Public Const SHCNE_EXTENDED_EVENT = &H4000000 '(G) Not currently used.
'Public Const SHCNE_ASSOCCHANGED = &H8000000   '(G) A file type association has changed.
        
'        Exit Sub

    Case SHCNE_RENAMEITEM
        If (Watch.Flags And FWF_FILE_RENAME) = 0 Then _
            Exit Sub

        szIcon = "file"
        szUid = Path2
        szReplace = Path1
        szTitle = "File renamed"
        szText = Path1 & " renamed to " & Path2
        szClass = FILE_RENAMED

    Case SHCNE_CREATE
        If (Watch.Flags And FWF_FILE_CREATE) = 0 Then _
            Exit Sub

        szIcon = "file"
        szUid = Path1
        szTitle = "File created"
        szText = Path1 & " was created"
        szClass = FILE_CREATED

    Case SHCNE_DELETE
        If (Watch.Flags And FWF_FILE_DELETE) = 0 Then _
            Exit Sub

        szIcon = "file"
        szUid = Path1
        szTitle = "File deleted"
        szText = Path1 & " was deleted"
        szClass = FILE_DELETED


    Case SHCNE_RENAMEFOLDER
        If (Watch.Flags And FWF_FOLDER_RENAME) = 0 Then _
            Exit Sub

        szIcon = "folder"
        szReplace = Path1
        szUid = Path2
        szTitle = "Folder renamed"
        szText = Path1 & " renamed to " & Path2
        szClass = FOLDER_RENAMED

    Case SHCNE_MKDIR
        If (Watch.Flags And FWF_FOLDER_CREATE) = 0 Then _
            Exit Sub

        szIcon = "folder"
        szUid = Path1
        szTitle = "Folder created"
        szText = Path1 & " was created"
        szClass = FOLDER_CREATED

    Case SHCNE_RMDIR
        If (Watch.Flags And FWF_FOLDER_DELETE) = 0 Then _
            Exit Sub

        szIcon = "folder"
        szUid = Path1
        szTitle = "Folder deleted"
        szText = Path1 & " was deleted"
        szClass = FOLDER_DELETED


'Public Const SHCNE_ATTRIBUTES = &H800        '(D) The attributes of an item or folder have changed.
'Public Const SHCNE_UPDATEDIR = &H1000        '(D) The contents of an existing folder have changed,
'                                '    but the folder still exists and has not been renamed.


    Case SHCNE_UPDATEITEM
        szTitle = "Item changed"
        szText = "Attributes for " & Path1 & " were changed"
        szUid = Path1

        If g_IsFolder(Path1) Then
            If (Watch.Flags And FWF_FOLDER_CHANGE) = 0 Then _
                Exit Sub

            szIcon = "folder"
            szClass = FOLDER_UPDATED

        Else
            If (Watch.Flags And FWF_FILE_CHANGE) = 0 Then _
                Exit Sub

            szIcon = "file"
            szClass = FILE_UPDATED

        End If

    Case Else
        Debug.Print "not implemented"
        Exit Sub

    End Select



    If szIcon <> "" Then _
        szIcon = g_MakePath(App.Path) & "icons\" & szIcon & ".png"

    Debug.Print "UID=" & szUid

    If mToken Then _
        snDoRequest "notify?app-sig=" & App.ProductName & _
                    "&id=" & szClass & _
                    "&uid=" & szUid & _
                    IIf(szReplace <> "", "&update-uid=" & szReplace, "") & _
                    "&title=" & szTitle & _
                    "&text=" & szText & _
                    "&icon=" & szIcon


End Sub

Private Sub theAddWatchPanel_Done(NewWatch As TFolderWatch)

    mFolders.Add NewWatch
    uUpdateList

    uWriteConfig

End Sub

Private Sub uUpdateList()

    If (mPanel Is Nothing) Then _
        Exit Sub

Dim pfw As TFolderWatch
Dim sz As String

    With mFolders
        .Rewind
        Do While .GetNextTag(pfw) = B_OK
            sz = sz & g_FormattedMidStr(pfw.Path, 55) & "#?0#?" & pfw.FlagsAsString() & "|"

        Loop

    End With

    sz = g_SafeLeftStr(sz, Len(sz) - 1)
    prefskit_SafeSetText mPanel, "watch_list", sz

End Sub

Private Sub uWriteConfig()
Dim pfw As TFolderWatch
Dim pcf As CConfFile3
Dim pcs As CConfSection

    Set pcf = New CConfFile3
    pcf.SetFile g_MakePath(App.Path) & "foldersnoop.conf"

    With mFolders
        .Rewind
        Do While .GetNextTag(pfw) = B_OK
            Set pcs = New CConfSection
            With pcs
                .SetName "watch"
                .Add "guid", pfw.Guid
                .Add "path", pfw.Path
                .Add "flags", pfw.Flags
                .Add "recurse", pfw.RecurseAsString

            End With

            pcf.Add pcs

        Loop

    End With

    pcf.Save

End Sub



