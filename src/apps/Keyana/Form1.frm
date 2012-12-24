VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THOR"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtconsole 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Cousine"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2115
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function GenerateConsoleCtrlEvent Lib "kernel32" (ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long

' These API declarations are for pipe communication stuff.
Private Declare Function CreatePipe Lib "kernel32" (pmReadPipe As Long, pmWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_CLOSE = &H10
Private Const WM_KEYDOWN = &H100
Private Const VK_CONTROL = &H11

' // Types related to pipe communication
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long

End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long

End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long

End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0

Dim mWritePipe As Long
Dim mReadPipe As Long
Dim mSendFlag As Boolean
Dim mLastPos As Integer

Dim mhWndConsole As Long
Dim mpiConsole As PROCESS_INFORMATION

Dim WithEvents thePanel As TAddPanel
Attribute thePanel.VB_VarHelpID = -1
Dim mPanel As BPrefsPanel
Dim mKeywords As BTagList

Implements BWndProcSink
Implements KPrefsPanel
Implements KPrefsPage

Private Function uOpenConsole(ByRef hWndConsole As Long) As Boolean
Dim sa As SECURITY_ATTRIBUTES

    ' /* create pipes */

    With sa
        .nLength = Len(sa)
        .bInheritHandle = True

    End With

Dim hRead As Long

    If CreatePipe(hRead, mWritePipe, sa, 1024) = 0 Then _
        Exit Function

Dim hWrite As Long

    If CreatePipe(mReadPipe, hWrite, sa, 1024) = 0 Then
        CloseHandle mWritePipe
        Exit Function

    End If

Dim si As STARTUPINFO

    ' /* spawn the app, re-directing stdin and stdout to the console's ends of the pipes */

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        .hStdOutput = hWrite
        .hStdError = hWrite
        .hStdInput = hRead
        .wShowWindow = SW_HIDE
        .lpTitle = "w>console" & g_HexStr(Me.hWnd)

    End With

Dim hr As Long

    hr = CreateProcessA(0&, "cmd.exe", _
                            sa, sa, -1, NORMAL_PRIORITY_CLASS, _
                            0&, 0&, si, mpiConsole)

    If hr Then

        Debug.Print "console->hProcess: " & g_HexStr(mpiConsole.hProcess) & " (" & CStr(mpiConsole.dwProcessId) & ")"
        Debug.Print "console->hThread: " & g_HexStr(mpiConsole.hThread) & " (" & CStr(mpiConsole.dwThreadId) & ")"

        ' /* wait for cmd.exe to start up (this seems to fail though) */

        hr = WaitForInputIdle(mpiConsole.hProcess, 1000)
        Debug.Print "wait: " & hr
        If hr <> 0 Then
            Debug.Print "wait failed: " & g_ApiError()
            Sleep 250

        End If

        ' /* find the console process' window */

        hWndConsole = FindWindow("ConsoleWindowClass", "w>console" & g_HexStr(Me.hWnd))
        uOpenConsole = (hWndConsole <> 0)

    Else
        Debug.Print "CreateProcess() failed: " & g_ApiError()

    End If

    ' /* don't need these now */

    CloseHandle hWrite
    CloseHandle hRead

End Function

Private Function uReadConsole(ByRef bShouldQuit As Boolean) As String
Dim cbAvail As Long

    If PeekNamedPipe(mReadPipe, ByVal 0&, 0, 0, cbAvail, 0) = 0 Then
        Debug.Print "uReadConsole: aborted: pipe was closed"
        bShouldQuit = True
        Exit Function

    End If

    ' /* if there is nothing to read, give an empty string */

    If cbAvail = 0 Then _
        Exit Function

    Debug.Print "uReadConsole: PeekNamedPipe->" & CStr(cbAvail)

Dim strBuffer As String
Dim cbRead As Long
Dim sz As String
Dim hr As Long

    ' /* as we know there's something there, read it all in and return it */

    strBuffer = String$(cbAvail, 0)

    hr = ReadFile(mReadPipe, strBuffer, cbAvail, cbRead, 0&)
    If cbRead > 0 Then _
        uReadConsole = Left$(strBuffer, cbRead)

End Function

Private Sub uWriteConsole(ByVal Text As String)
Dim b() As Byte
Dim i As Long

    b = StrConv(Text, vbFromUnicode)
    If WriteFile(mWritePipe, b(0), Len(Text), i, CLng(0)) = 0 Then _
        Debug.Print "uWriteConsole(): failed"

End Sub

Private Sub uCloseConsole()

    ' /* kill the console window */

    If mhWndConsole Then _
        PostMessage mhWndConsole, WM_CLOSE, 0, 0

    ' /* close handles to pipes */

    CloseHandle mWritePipe
    CloseHandle mReadPipe

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim pmi As OMMenuItem

    Select Case uMsg
    Case WM_RBUTTONUP
        With New OMMenu
            .AddItem .CreateItem("cut", "Cut", , False)
            .AddItem .CreateItem("copy", "Copy", , (txtconsole.SelLength > 0))
            .AddItem .CreateItem("paste", "Paste", , (Clipboard.GetText(vbCFText) <> ""))
            .AddSeparator
            .AddItem .CreateItem("sel", "Select All")
            .AddSeparator
            .AddItem .CreateItem("options", "Options...")

            Set pmi = .Track(Me.hWnd)
            If Not (pmi Is Nothing) Then
                Select Case pmi.Name
                Case "cut"
                    SendMessage txtconsole.hWnd, WM_CUT, 0, ByVal 0&
                
                Case "copy"
                    SendMessage txtconsole.hWnd, WM_COPY, 0, ByVal 0&
                
                Case "paste"
                    SendMessage txtconsole.hWnd, WM_PASTE, 0, ByVal 0&

                Case "sel"
                    With txtconsole
                        .SelStart = 1
                        .SelLength = Len(.Text)

                    End With

                Case "options"
                    uDoPrefs

                End Select

            End If

        End With

        ReturnValue = 0
        BWndProcSink_WndProc = True

    End Select

End Function

Private Sub Form_Load()

    If Not uOpenConsole(mhWndConsole) Then
        MsgBox "failed!"
        Unload Me
        Exit Sub

    End If

    Me.Caption = App.Title
    g_SetWindowIconToAppResourceIcon2 Me.hWnd

    Debug.Print "created console: " & g_HexStr(mhWndConsole)
    uWrite vbCrLf & App.Title & " 1.0/" & App.Revision & " (c) full phat products" & vbCrLf

    With txtconsole
        .BackColor = RGB(0, 0, 36)
        .ForeColor = RGB(240, 240, 240)

    End With

    ' /* register with Snarl */

    snarl_register App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png"

    ' /* load config */

    Set mKeywords = new_BTagList()
    uGetKeywords

    uWrite vbCrLf & "right-click this window for options"
    uWrite vbCrLf & "launching command shell..."
    uWrite vbCrLf & vbCrLf

    window_subclass txtconsole.hWnd, Me

    Me.Show

    ' /* enter the read loop, we quit when this quits */

    uReadLoop

    Unload Me

End Sub

Private Sub Form_Resize()

    txtconsole.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not (mPanel Is Nothing) Then
        mPanel.Quit
        Set mPanel = Nothing

    End If

    window_subclass txtconsole.hWnd, Nothing
    uCloseConsole

    snarl_unregister App.ProductName

End Sub

Private Sub thePanel_Done(Item As TKeyword)

    mKeywords.Add Item
    uUpdateList
    uWriteKeywords

    Set thePanel = Nothing

End Sub

Private Sub txtconsole_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
'    Case 13
'        KeyCode = 0

    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown
        KeyCode = 0

    Case Else
        With txtconsole
            .SelLength = 0
            .SelStart = Len(.Text)

        End With

    End Select

End Sub

Private Sub txtconsole_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case 3
        ' /* CTRL+C */
        Debug.Print "sending CTRL+C sequence to console " & g_HexStr(mhWndConsole)
        PostMessage mhWndConsole, WM_KEYDOWN, VK_CONTROL, 0
        PostMessage mhWndConsole, WM_KEYDOWN, vbKeyC, 0

'    Case Else
'        Debug.Print KeyAscii
    
    End Select

End Sub

Private Sub txtConsole_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sz As String

    On Error Resume Next

    If KeyCode = 13 Then

        If mLastPos Then
            sz = Right$(txtconsole.Text, (txtconsole.SelStart - mLastPos))
'            Debug.Print sz & " " & Len(sz)
            uWriteConsole sz

        End If

        mLastPos = txtconsole.SelStart
        mSendFlag = True
        KeyCode = 0

    End If

End Sub

Private Sub uReadLoop()
Dim bQuit As Boolean
Dim sz As String

    Do

        Sleep 10
        DoEvents

        sz = uReadConsole(bQuit)
        If bQuit Then
            Exit Sub

        ElseIf sz <> "" Then

            If mSendFlag Then
                ' /* remove first line, which would be what was sent */
                sz = Right$(sz, Len(sz) - (InStr(1, sz, vbCrLf) + 1))
                mSendFlag = False

            End If

            uWrite sz

            ' /* search for keyword triggers */

            uScanKeywords sz

        End If

    Loop

End Sub

Private Sub uScanKeywords(ByVal Text As String)
Dim pk As TKeyword
Dim i As Long

    i = InStrRev(Text, vbCrLf)
    If i Then _
        Text = g_SafeLeftStr(Text, i - 1)

    With mKeywords
        .Rewind

        Do While .GetNextTag(pk) = B_OK
            pk.Scan Text

        Loop

    End With

End Sub

Private Sub uWrite(ByVal Text As String)

    With txtconsole
        .Text = .Text & Text
        .SelStart = Len(.Text)
        .SelLength = 0
        mLastPos = .SelStart

    End With

End Sub

Private Sub uGetKeywords()
Dim ps As CConfSection
Dim pk As TKeyword

    With New CConfFile3
        .SetFile g_MakePath(App.Path) & "keywords.txt"
        If .Load Then
            .Rewind
            Do While .GetNextSection(ps)
                Set pk = New TKeyword
                If pk.SetFromExisting(ps) Then _
                    mKeywords.Add pk

            Loop

            uWrite vbCrLf & "loaded keyword list (" & g_MakePath(App.Path) & "keywords.txt)"

        Else
            uWrite vbCrLf & "getkeywords: failed to load keyword list"

        End If

    End With

End Sub

Private Sub uWriteKeywords()
Dim pk As TKeyword

    With New CConfFile3
        .SetFile g_MakePath(App.Path) & "keywords.txt"

        mKeywords.Rewind
        Do While mKeywords.GetNextTag(pk) = B_OK
            .Add pk.CreateSection()

        Loop

        .Save

    End With

End Sub

'Public Function KeywordList() As BTagList
'
'    Set KeywordList = mKeywords
'
'End Function

Private Sub uDoPrefs()
Dim pPage As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "Keyana Preferences"
            .SetWidth 400

            Set pPage = new_BPrefsPage("", , Me)

            With pPage
                .SetMargin 0

                .Add new_BPrefsControl("label", "", "Keywords")

                ' /* list */

                Set pm = New CTempMsg
                pm.Add "item-height", 36&
                Set pc = new_BPrefsControl("listbox", "item_list", , , , pm)
                pc.SizeTo 0, 220
                .Add pc

                .Add new_BPrefsControl("fancyplusminus", "add_remove", "")

                ' /* toolbar */

'                .Add new_BPrefsControl("fancytoolbar", "toolbar_options", "Invoke|Actions|Copy to Clipboard|Display", , , , False)

                ' /* icon and content */

'                Set pm = New CTempMsg
'                pm.Add "scale_to_fit", 1&
'                Set pc = new_BPrefsControl("image", "the_icon", "", , , pm)
'                pc.SizeTo 48, 48
'                .Add pc
'
'                .Add new_BPrefsControl("label", "the_detail", Space$(256))
'
''                .Add new_BPrefsControl("fancybutton2", "fb>ack", "Invoke Callback", , , , False)
'
''                .Add new_BPrefsControl("seperator", "")
'                .Add new_BPrefsControl("fancybutton2", "fb>clear", "Clear List")
        
            End With

            .AddPage pPage

            .Go

            uUpdateList

            g_SetWindowIconToAppResourceIcon .hWnd

        End With

    Else
        g_ShowWindow mPanel.hWnd, True, True
        SetForegroundWindow mPanel.hWnd

    End If

End Sub

Private Sub uUpdateList()

    If (mPanel Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    If Not mPanel.Find("item_list", pc) Then
        g_Debug "TConfigPanel.uUpdateList(): can't find item_list control", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim sz As String
Dim pk As TKeyword

    With mKeywords
        If .CountItems Then
            .Rewind
            Do While .GetNextTag(pk) = B_OK
                sz = sz & pk.Description & "#?0#?" & pk.Keyword & "|"

            Loop

            sz = g_SafeLeftStr(sz, Len(sz) - 1)

        End If

        pc.SetText sz

    End With

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Static iCurrent As Long

    Select Case Control.GetName()
    Case "item_list"
        iCurrent = Val(Value)

    Case "add_remove"
        If Value = "+" Then
            Set thePanel = New TAddPanel
            thePanel.Go mPanel.hWnd

        ElseIf Value = "-" Then

            ' /* remove the class from Snarl */
            snDoRequest "remclass?app-sig=" & App.ProductName & _
                        "&id=" & mKeywords.TagAt(iCurrent).Name

            ' /* remove from the list and update */
            mKeywords.Remove iCurrent
            uUpdateList
            uWriteKeywords

        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
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


