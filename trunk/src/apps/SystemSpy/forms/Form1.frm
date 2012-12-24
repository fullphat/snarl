VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#"
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

Public Enum SS_PROCESS_MODES
    SS_PROCESS_INCLUSIVE = 0
    SS_PROCESS_EXCLUSIVE = 1

End Enum

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_ICON = 1

Dim mPassword As String
Dim mNextMsg As Long

Dim theTrayIcon As BNotifyIcon
Dim WithEvents theAddPanel As TAddPanel
Attribute theAddPanel.VB_VarHelpID = -1

Public WithEvents theWindowSpy As TWindowSpy
Attribute theWindowSpy.VB_VarHelpID = -1
Public WithEvents theFolderSpy As TFolderSpy
Attribute theFolderSpy.VB_VarHelpID = -1
Public WithEvents theProcessSpy As TProcessSpy
Attribute theProcessSpy.VB_VarHelpID = -1
Public WithEvents theAppSpy As TAppSpy
Attribute theAppSpy.VB_VarHelpID = -1

Private Const WINDOW_APPEARED = "windowappeared"
Private Const WINDOW_DISAPPEARED = "windowdisappeared"

Private Const PROCESS_STARTED = "processstarted"
Private Const PROCESS_STOPPED = "processstopped"

Private Const APP_LAUNCHED = "applaunched"
Private Const APP_QUIT = "appquit"

Private Const FOLDER_CREATED = "foldercreated"
Private Const FOLDER_RENAMED = "folderrenamed"
Private Const FOLDER_DELETED = "folderdeleted"
Private Const FOLDER_UPDATED = "folderupdated"

Private Const FILE_CREATED = "filecreated"
Private Const FILE_RENAMED = "filerenamed"
Private Const FILE_DELETED = "filedeleted"
Private Const FILE_UPDATED = "fileupdated"

Private Const WM_FOLDER_SPY_START = WM_USER + 10
Private Const WM_FOLDER_SPY_END = WM_FOLDER_SPY_START + 64

Implements BWndProcSink

Private Sub Form_Initialize()

    g_InitComCtl

End Sub

Private Sub Form_Load()
Dim bQuit As Boolean
Dim hWndPrev As Long

    hWndPrev = FindWindow("ThunderRT6FormDC", "SystemSpy")
    bQuit = (InStr(Command, "-quit") <> 0)

    If (hWndPrev <> 0) Or (bQuit) Then
        ' /* if -quit specified, tell the other instance to quit */
        If bQuit Then _
            PostMessage hWndPrev, WM_CLOSE, 0, ByVal 0&

        ' /* unload either way */
        Unload Me
        Exit Sub

    End If

    Me.Add "starting..."
    Me.Add "  " & App.Title & " " & App.Major & "." & App.Minor & " Build " & App.Revision
    Me.Caption = "SystemSpy"
    mNextMsg = WM_USER + 10

    Me.Add "checking paths..."

Dim sz As String

    sz = Me.GetConfigPath(False)
    If Not g_IsFolder(sz) Then
        Me.Add "panic: config path error"
        MsgBox "Configuration path missing.  Cannot continue, sorry.", vbCritical Or vbOKOnly, App.Title
        Unload Me
        Exit Sub

    End If

    Me.Add "creating spies..."
    Set theWindowSpy = New TWindowSpy
    Set theFolderSpy = New TFolderSpy
    Set theProcessSpy = New TProcessSpy
    Set theAppSpy = New TAppSpy

    Me.Add "subclassing..."
    window_subclass Me.hWnd, Me

    Me.Add "creating password..."
    mPassword = create_password()

    Me.Add "creating tray icon..."
    Set theTrayIcon = New BNotifyIcon
    With theTrayIcon
        .SetTo Me.hWnd, WM_USER
        If g_IsIDE() Then
            .Add "icon", Me.Icon.Handle, Me.Caption

        Else
            .Add "icon", LoadImage(App.hInstance, "#101", IMAGE_ICON, 16, 16, 0), Me.Caption

        End If

    End With

    Me.Hide

    ' /* register with Snarl, if it's around */

    Me.Add "registering with Snarl..."
    uRegister

    Me.Add "starting spies..."
    theWindowSpy.Go
    theFolderSpy.Go
    theProcessSpy.Go
    theAppSpy.Go

    Me.Add "ready"

    If g_IsIDE() Then _
        uDoPrefs

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Me.Add "stopping..."

    Me.Add "ending spies..."
    Set theFolderSpy = Nothing
    Set theWindowSpy = Nothing
    Set theProcessSpy = Nothing
    Set theAppSpy = Nothing

    Me.Add "unsubclassing..."
    window_subclass Me.hWnd, Nothing

    Me.Add "closing panel..."
    Unload frmSettings

    Me.Add "removing tray icon..."
    If Not (theTrayIcon Is Nothing) Then
        theTrayIcon.Remove "icon"
        Set theTrayIcon = Nothing

    End If

    Me.Add "unregistering..."
    snarl_unregister App.ProductName, mPassword

    Me.Add "done"

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim pmi As OMMenuItem

    Select Case uMsg

    Case WM_USER
        ' /* tray icon */
        Select Case lParam
        Case WM_LBUTTONDBLCLK
            uDoPrefs

        Case WM_RBUTTONUP
            With New OMMenu
                .AddItem .CreateItem("about", "About...")
                .AddItem .CreateItem("prefs", "Preferences...")
                .AddSeparator
                .AddItem .CreateItem("quit", "Quit")

                SetForegroundWindow hWnd
                Set pmi = .Track(hWnd)
                PostMessage hWnd, WM_NULL, 0, ByVal 0&

                If Not (pmi Is Nothing) Then
                    Select Case pmi.Name
                    Case "about"
                        PostMessage hWnd, snAppMsg(), SNARLAPP_DO_ABOUT, ByVal 0&

                    Case "prefs"
                        uDoPrefs

                    Case "quit"
                        Unload Me

                    End Select

                End If

            End With

        End Select


    Case WM_CLOSE
        Me.Add "WM_CLOSE"
        Unload Me


    Case WM_SYSCOMMAND
        If (LoWord(wParam) And &HFFF0&) = SC_MINIMIZE Then
            g_ShowWindow hWnd, False
            ReturnValue = 0
            BWndProcSink_WndProc = True

        End If


    Case snAppMsg()
        Select Case wParam
        Case SNARLAPP_DO_PREFS, SNARLAPP_ACTIVATED
            uDoPrefs

        Case SNARLAPP_DO_ABOUT
            snarl_notify App.ProductName, "", "", mPassword, _
                         "SystemSpy " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments & " (Build " & CStr(App.Revision) & ")", _
                         "Notifies when folder contents change, processes and applications start and stop, and when windows are created and destroyed.", _
                         g_MakePath(App.Path) & "icon.png"
'App.LegalCopyright & vbCrLf &

        Case SNARLAPP_QUIT_REQUESTED
            Unload Me

        End Select


    Case WM_FOLDER_SPY_START To WM_FOLDER_SPY_END
        ' /* a ShellChangeNotify event */
        theFolderSpy.ShellChangeNotify uMsg, wParam, lParam

    End Select

End Function

Private Sub theAddPanel_Done(NewItem As TRule)

'    Debug.Print "add: " & NewItem.Guid & " " & NewItem.Class & " " & NewItem.Title
'
'    ' /* add it to our list */
'
'    mRules.Add NewItem
'
'    ' /* add it to Snarl */
'
'    If mToken <> 0 Then _
'        snDoRequest "addclass?token=" & CStr(mToken) & "&id=" & NewItem.Guid & "&name=Title: " & NewItem.Title & " Class: " & NewItem.Class
'
'    ' /* refresh config window */
'
'    uUpdateList
'
'    ' /* write out the updated config */
'
'    uWriteConfig

End Sub

Private Sub uNotifyWindowEvent(ByRef Rule As TRule, ByVal Title As String, ByVal Class As String, ByVal hWnd As Long)
Dim lIcon As Long
Dim sz As String

    If Title = "" Then _
        Title = "<null>"

    Title = Title & "\n" & Class
    lIcon = g_WindowIcon(hWnd, False, False)

    snDoRequest "notify?app-sig=" & App.ProductName & _
                "&password=" & mPassword & _
                "&class=" & Rule.Guid & _
                "&uid=" & CStr(hWnd) & _
                "&replace-uid=" & CStr(hWnd) & _
                "&title=Window appeared" & _
                "&text=" & Title & _
                "&icon=" & IIf(lIcon = 0, g_MakePath(App.Path) & "new.png", "%" & CStr(lIcon))

End Sub

Public Sub Add(ByVal Text As String)

    List1.AddItem Text
    List1.ListIndex = List1.ListCount - 1
    g_Debug Text

End Sub

Private Function uRegister() As Boolean

    If snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", mPassword, Me.hWnd, , True) < 0 Then _
        Exit Function

    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & WINDOW_APPEARED & "&name=Window appeared"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & WINDOW_DISAPPEARED & "&name=Window disappeared"

    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & FOLDER_CREATED & "&name=Folder created"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & FOLDER_RENAMED & "&name=Folder renamed"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & FOLDER_DELETED & "&name=Folder deleted"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & FOLDER_UPDATED & "&name=Folder updated"

    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & FILE_CREATED & "&name=File created"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & FILE_RENAMED & "&name=File renamed"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & FILE_DELETED & "&name=File deleted"
    snDoRequest "addclass?app-sig=" & App.ProductName & "&password=" & mPassword & "&id=" & FILE_UPDATED & "&name=File updated"

    uRegister = True

'Dim pr As TRule
'
'    With mRules
'        .Rewind
'        Do While .GetNextTag(pr) = B_OK
'            snDoRequest "addclass?token=" & CStr(hr) & "&id=" & pr.Guid & "&name=Title: " & pr.Title & " Class: " & pr.Class
'
'        Loop
'
'    End With

End Function

Private Sub uDoPrefs()

    frmSettings.Go

End Sub

'    Select Case Control.GetName
'
'    Case "add_remove"
'        If Value = "+" Then
'            Set theAddPanel = New TAddPanel
'            theAddPanel.Go mPanel.hWnd
'
'        ElseIf (Value = "-") Then
'            i = Val(prefskit_GetValue(mPanel, "rules"))
'            Set pr = mRules.TagAt(i)
'            If (pr Is Nothing) Then _
'                Exit Sub
'
'            mRules.Remove i
'            uWriteConfig
'            uUpdateList
'
'            snDoRequest "remclass?app-sig=" & App.ProductName & "&id=" & pr.Guid
'
'            prefskit_SetValue mPanel, "rules", CStr(i)
'
'        End If
'
''    Case "UseDefaultCallback"
''        gConfig.UseDefaultCallback = (Value = "1")
''        uUpdateConfig
'
'    End Select

Public Function GetConfigPath(ByVal MakePath As Boolean) As String

    GetConfigPath = g_MakePath(g_GetSystemFolderStr(CSIDL_APPDATA)) & "k23 productions\SystemSpy"
    If MakePath Then _
        GetConfigPath = g_MakePath(GetConfigPath)

End Function

Public Function NextFreeMsg() As Long

    If mNextMsg > WM_FOLDER_SPY_END Then
        Me.Add "out of folder spy slots!"
        Exit Function

    End If

    NextFreeMsg = mNextMsg
    mNextMsg = mNextMsg + 1

End Function

Private Sub uNotifyFolderSpyEvent(ByVal Class As String, ByVal Title As String, ByVal Text As String, ByVal Icon As String, ByVal UID As String, Optional ByVal ReplaceUID As String)

    If Icon <> "" Then _
        Icon = g_MakePath(App.Path) & "icons\" & Icon & ".png"

    If uRegister() Then _
        snDoRequest "notify?app-sig=" & App.ProductName & _
                    "&password=" & mPassword & _
                    "&id=" & Class & _
                    "&uid=" & UID & _
                    IIf(ReplaceUID <> "", "&update-uid=" & ReplaceUID, "") & _
                    "&title=" & Title & _
                    "&text=" & Text & _
                    "&icon=" & Icon

End Sub

Private Sub theAppSpy_AppLaunched(Process As TProcess)
Dim sz As String

    sz = Process.Description
    If sz = "" Then _
        sz = g_RemoveExtension(Process.Name)

    uNotify APP_LAUNCHED, "app-" & CStr(Process.Pid), mPassword, _
            "Application launched", sz, g_MakePath(App.Path) & "icons\app-launched.png"

End Sub

Private Sub theAppSpy_AppQuit(Process As TProcess)
Dim sz As String

    sz = Process.Description
    If sz = "" Then _
        sz = g_RemoveExtension(Process.Name)

    uNotify APP_QUIT, "app-" & CStr(Process.Pid), mPassword, _
            "Application quit", sz, g_MakePath(App.Path) & "icons\app-quit.png"

End Sub

Private Sub theFolderSpy_FileCreated(ByVal Path As String)

    uNotifyFolderSpyEvent FILE_CREATED, "File created", Path & " was created", "file-created", Path

End Sub

Private Sub theFolderSpy_FileDeleted(ByVal Path As String)

    uNotifyFolderSpyEvent FILE_DELETED, "File deleted", Path & " was deleted", "file-deleted", Path

End Sub

Private Sub theFolderSpy_FileRenamed(ByVal Was As String, ByVal Now As String)

    uNotifyFolderSpyEvent FILE_RENAMED, "File renamed", Was & " renamed to " & Now, "file-renamed", Now, Was

End Sub

Private Sub theFolderSpy_FolderCreated(ByVal Path As String)

    uNotifyFolderSpyEvent FOLDER_CREATED, "Folder created", Path & " was created", "folder-created", Path

End Sub

Private Sub theFolderSpy_FolderDeleted(ByVal Path As String)

    uNotifyFolderSpyEvent FOLDER_DELETED, "Folder deleted", Path & " was deleted", "folder-deleted", Path

End Sub

Private Sub theFolderSpy_FolderRenamed(ByVal Was As String, ByVal Now As String)

    uNotifyFolderSpyEvent FOLDER_RENAMED, "Folder renamed", Was & " renamed to " & Now, "folder-renamed", Now, Was

End Sub

Private Sub theProcessSpy_ProcessStarted(Process As TProcess)
Dim sz As String

    sz = Process.Description
    If sz <> "" Then _
        sz = sz & vbCrLf

    sz = sz & Process.Name & " (" & CStr(Process.Pid) & ")"

    uNotify PROCESS_STARTED, "process-" & CStr(Process.Pid), mPassword, _
            "Process started", sz, g_MakePath(App.Path) & "icons\process-started.png"

End Sub

Private Sub theProcessSpy_ProcessStopped(Process As TProcess)
Dim sz As String

    sz = Process.Description
    If sz <> "" Then _
        sz = sz & vbCrLf

    sz = sz & Process.Name & " (" & CStr(Process.Pid) & ")"

    uNotify PROCESS_STOPPED, "process-" & CStr(Process.Pid), mPassword, _
            "Process stopped", sz, g_MakePath(App.Path) & "icons\process-stopped.png"

End Sub

'Public Sub RemoveFolderSpy(ByVal Index As Long)
'
'    theFolderSpy.Remove Index
'
'End Sub

Private Sub theWindowSpy_WindowAppeared(MatchingRule As TRule, ByVal Title As String, ByVal Class As String, ByVal hWnd As Long)

    Debug.Print "window appeared (rule=" & MatchingRule.Title & ") title=" & Title & " class=" & Class

End Sub

Private Sub theWindowSpy_WindowDisappeared(MatchingRule As TRule, ByVal Title As String, ByVal Class As String, ByVal hWnd As Long)

    Debug.Print "window disappeared (rule=" & MatchingRule.Title & ") title=" & Title & " class=" & Class

End Sub

Private Function uNotify(ByVal Class As String, ByVal UID As String, Optional ByVal Password As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Icon As String, Optional ByVal Priority As Long, Optional ByVal Duration As Long = -1, Optional ByVal Callback As String, Optional ByVal PercentValue As Long = -1, Optional ByVal CustomData As String) As Long

    If uRegister() Then _
        uNotify = (snarl_notify(App.ProductName, Class, UID, Password, Title, Text, Icon, Priority, Duration, Callback, PercentValue, CustomData) >= 0)

End Function

