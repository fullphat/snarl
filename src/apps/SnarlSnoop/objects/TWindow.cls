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

Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2

Private Const TIMER_PULSE = 1

Private Const WM_TIMER = &H113
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim mhWnd As Long
Dim mToken As Long
Dim mPanel As BPrefsPanel

Dim WithEvents theAddEditPanel As TAddEditPanel
Attribute theAddEditPanel.VB_VarHelpID = -1
Dim mSelected As Long

Dim WithEvents theRoster As TProcessRoster
Attribute theRoster.VB_VarHelpID = -1

Dim mConfig As CConfFile3

'Dim mItem() As CConfSection
'Dim mItems As Long

Implements KPrefsPage
Implements KPrefsPanel
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim pObj As Object
Dim n As Long

    On Error Resume Next

    Select Case uMsg
    Case WM_CREATE
        Form1.Add "WM_CREATE"
        mhWnd = hwnd
        uStart


    Case WM_DESTROY
        Form1.Add "WM_DESTROY"
        If mToken Then _
            sn41UnregisterApp mToken

        If Not (mPanel Is Nothing) Then _
            mPanel.Quit


    Case WM_CLOSE
        PostQuitMessage 0


    Case WM_TIMER
        Select Case wParam
        Case TIMER_PULSE
            theRoster.Refresh

        End Select


    Case sn41Broadcast()
        Select Case wParam
        Case SNARL_LAUNCHED
            Form1.Add "Snarl launched"
            uRegisterWithSnarl

        Case SNARL_QUIT
            Form1.Add "Snarl quit"
            KillTimer hwnd, TIMER_PULSE
            mToken = 0      ' // no point in unregistering...

        End Select


    Case sn41AppMsg()
        Select Case wParam
        Case SNARL41_APP_PREFS
            Form1.Add "_APP_PREFS"
            uDoPrefs

        Case SNARL41_APP_ABOUT
            sn41EZNotify mToken, "", _
                         "SnarlTasks " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments, _
                         App.LegalCopyright & vbCrLf & "Build " & CStr(App.Revision), , _
                         g_MakePath(App.Path) & "icon.png"

        End Select


    Case WM_RELOAD
        Form1.Add "RELOAD SIGNAL"


    Case WM_NOTIFICATION
        Form1.Add "NOTIFICATION: " & g_HexStr(wParam) & " " & CStr(lParam)
        n = LoWord(wParam)

'        Select Case n
'        Case SNARL_NOTIFICATION_ACK, SNARL_NOTIFICATION_MENU
'            If Not uFind(lParam, pObj) Then _
'                Exit Function
'
'            Form1.Add "item found: " & pObj.Class
'
'            If n = SNARL_NOTIFICATION_ACK Then
'                pObj.Display
'
'            ElseIf n = SNARL_NOTIFICATION_MENU Then
'
'                sn41Hide lParam
'                pObj.UnRead = False
'
'                Select Case HiWord(wParam)
'                Case 1
'                    pObj.Display
'
'                Case 2
'                    pObj.Reply.Display
'
'                Case 3
'                    pObj.ReplyAll.Display
'
'                Case 4
'                    pObj.Forward.Display
'
'                End Select
'
'            End If
'
'
''        snChangeAttribute hr, SNARL_ATTRIBUTE_MENU, "Open#?1||Reply#?2|Reply All#?3|Forward#?4||Delete#?5"
'
'        End Select


    Case WM_TEST
        Form1.Add "TEST: " & wParam

'        If wParam = 0 Then
'            uNotifyMail Nothing
'
'        ElseIf wParam = 1 Then
'            uNotifyMail myItems.Item(1)
'
'        ElseIf wParam = 2 Then
'            uFindFirst olMeetingRequest, pObj
'            Debug.Print VarType(pObj)
'            Err.Clear
'            uNotifyAppointment pObj
'            Debug.Print Err.Description
'
'        End If

'    Case Else
'        form1.add "WM_" & g_HexStr(uMsg)

    End Select

End Function

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim i As Long

    Select Case Control.GetName()

    Case "task_list"
        mSelected = Val(Value)

    Case "toolbar"
        Select Case Val(Value)
        Case 1
            ' /* add new */
'            Set theAddEditPanel = New TAddEditPanel
'            theAddEditPanel.Go mPanel.hWnd, uCreateNewTask(), True


        Case 2
            ' /* edit selected */
            If mSelected = 0 Then _
                Exit Sub

'            Set theAddEditPanel = New TAddEditPanel
'            theAddEditPanel.Go mPanel.hWnd, mConfig.SectionAt(mSelected), False


        Case 4
            ' /* delete selected */
            If mSelected = 0 Then _
                Exit Sub

'            i = mConfig.IndexOf(mConfig.SectionAt(mSelected).Name)
'            If i Then
'                If mConfig.RemoveSection(i, True) Then
'                    uSetTaskList i
'                    uAddTasksToSnarl
'
'                End If
'            End If

        End Select

    Case Else

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName()
    Case "task_list"
        If mSelected = 0 Then _
            Exit Sub

        Set theAddEditPanel = New TAddEditPanel
        theAddEditPanel.Go mPanel.hwnd, mConfig.SectionAt(mSelected), False

    End Select

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

Private Sub uRegisterWithSnarl()

    mToken = sn41RegisterApp(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", mhWnd, WM_NOTIFICATION, SNARL41_APP_IS_WINDOWLESS Or SNARL41_APP_HAS_PREFS Or SNARL41_APP_HAS_ABOUT)
    If mToken Then
        Form1.List1.Tag = CStr(mToken)
        Form1.Add "Registered with Snarl (token=" & CStr(mToken) & ")"
        Set theRoster = New TProcessRoster
        With theRoster
            .LockRoster
            .Refresh
            .UnlockRoster

        End With

        sn41AddClass mToken, "padd", "Process Appeared"
        sn41AddClass mToken, "prem", "Process Disappeared"

        SetTimer mhWnd, TIMER_PULSE, 250, 0

    Else
        Form1.Add "Error registering with Snarl"

    End If

End Sub

Private Sub uDoPrefs()

    If Not (mPanel Is Nothing) Then _
        Exit Sub

Dim aPage As BPrefsPage
Dim aMsg As CTempMsg
Dim pc As BControl

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle "SnarlSnooper"
        .SetWidth 380

        Set aPage = new_BPrefsPage("", Nothing, Me)
        With aPage
            .SetMargin 90

            .Add new_BPrefsControl("fancycycle", "", "All processes|Only processes in the Include list|All processes except those in the Exclude list", "Scope:")


            .Add new_BPrefsControl("banner", "", "Include List")

            Set aMsg = New CTempMsg
            aMsg.Add "item-height", 20&
            Set pc = new_BPrefsControl("listbox", "", "", "", "0", aMsg)
'            pc.SizeTo 0, 300
            .Add pc

            .Add new_BPrefsControl("banner", "", "Exclude List")

            Set aMsg = New CTempMsg
            aMsg.Add "item-height", 20&
            Set pc = new_BPrefsControl("listbox", "", "", "", "0", aMsg)
'            pc.SizeTo 0, 300
            .Add pc

'        Set pm = New CTempMsg
'        pm.Add "text", "Use the same settings for all notification classes?"
'        pm.Add "align", 1
'        pm.Add "inset_by", 0
''    mPage.Add new_BPrefsControl("label", "", "If enabled, Snarl will apply the settings defined for the <Other> or <All> class to all notifications generated by the application.")
'
'        .Add new_BPrefsControl("banner", "", "Notification classes")
'
'        Set pm = New CTempMsg
'        pm.Add "checkboxes", 1
'        .Add new_BPrefsControl("listbox", "lb>classes", "", "", 1, pm)
'        .Add new_BPrefsControl("fancybutton2", "fb>cfg_class", "Configure...")




        End With
        .AddPage aPage


'        uSetTaskList
'        If mConfig.CountSections > 0 Then _
'            mSelected = 1

        .Go

'        g_SetWindowIconToAppResourceIcon .hWnd

    End With

End Sub

Private Sub uStart()

    If sn41IsSnarlRunning() Then
        uRegisterWithSnarl

    Else
        Form1.Add "Snarl not running, waiting..."

    End If

End Sub

'Private Sub uLoadTasks()
''Dim pTask As CConfSection
''
''    Set mConfig = New CConfFile3
''    With mConfig
''        .SetFile g_MakePath(App.Path) & ".crond"
''        .Load
''
''    End With
'
'
'
'End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

'Private Sub uSetTaskList(Optional ByVal SelectedItem As Long = 1)
'
'    If (mPanel Is Nothing) Then _
'        Exit Sub
'
'Dim pc As BControl
'
'    If Not mPanel.Find("task_list", pc) Then _
'        Exit Sub
'
'    pc.DoExCmd B_RESET_CONTENT                  ' // reset the list
'
'Dim listContent As String
'Dim pTask As CConfSection
'
'    With mConfig
'        .Rewind
'        Do While .GetNextSection(pTask)
'            listContent = listContent & uItemName(pTask) & "#?" & CStr(0) & "#?" & uGetRecurrenceString(pTask) & "|"
'
'        Loop
'
'    End With
'
'    ' /* set the content */
'
'    pc.SetText g_SafeLeftStr(listContent, Len(listContent) - 1)
'
'    ' /* set the icons */
'
'Dim pm As CTempMsg
'Dim i As Long
'
'    Set pm = New CTempMsg
'
'    With mConfig
'        .Rewind
'        i = 1
'        Do While .GetNextSection(pTask)
'            pm.Replace "index", i
'            pm.Replace "image-file", pTask.GetValueWithDefault("icon")
'            pc.DoExCmd B_SET_ITEM, pm
'            i = i + 1
'
'        Loop
'    End With
'
'    pc.SetValue CStr(SelectedItem)
'
'
'End Sub


'Private Function uGetNewUId() As String
'
'    If (mConfig Is Nothing) Then _
'        Exit Function
'
'Dim i As Integer
'
'    Do While mConfig.Exists("task:" & Format$(i, "0000"))
'        i = i + 1
'
'    Loop
'
'    uGetNewUId = "task:" & Format$(i, "0000")
'
'End Function

Private Sub theAddEditPanel_Update(Item As CConfSection)
Dim pce As CConfEntry
Dim i As Long

    i = mConfig.IndexOf(Item.Name)
    If i = 0 Then
        ' /* add new */
        mConfig.Add Item

    Else
        ' /* update existing */
        With Item
            .Rewind
            Do While .NextEntry(pce)
                mConfig.SectionAt(i).Update pce.Name, pce.Value

            Loop
        End With

    End If

'    mConfig.Save
'    uSetTaskList
'    uAddTasksToSnarl

End Sub

'Private Function uItemName(ByRef Item As CConfSection) As String
'Dim sz As String
'
'    sz = Item.GetValueWithDefault("title")
'    If sz = "" Then _
'        sz = "{" & Item.Name & "}"
'
'    uItemName = sz
'
'End Function

Private Sub theRoster_ProcessAppeared(Details As misc_resource.BPackedData)
Dim szDesc As String
Dim szText As String

    szDesc = Details.ValueOf("description")
    If szDesc <> "" Then
        szDesc = szDesc & " (" & Details.ValueOf("name") & ")"

    Else
        szDesc = Details.ValueOf("name")

    End If

    szText = "pid " & Details.ValueOf("pid") & vbCrLf & Details.ValueOf("path")

    sn41EZNotify mToken, "padd", "Process Appeared", _
                 szDesc & vbCrLf & vbCrLf & szText, , Details.ValueOf("path") & ",-1"

End Sub

Private Sub theRoster_ProcessDisappeared(Details As misc_resource.BPackedData)
Dim szDesc As String
Dim szText As String

    szDesc = Details.ValueOf("description")
    If szDesc <> "" Then
        szDesc = szDesc & " (" & Details.ValueOf("name") & ")"

    Else
        szDesc = Details.ValueOf("name")

    End If

    szText = "pid " & Details.ValueOf("pid") & vbCrLf & Details.ValueOf("path")

    sn41EZNotify mToken, "prem", "Process Disappeared", _
                 szDesc & vbCrLf & vbCrLf & szText, , Details.ValueOf("path") & ",-1"

End Sub
