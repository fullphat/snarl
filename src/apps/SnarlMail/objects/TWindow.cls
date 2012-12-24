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

    '/*********************************************************************************************
    '/
    '/  File:           TWindow.cls
    '/
    '/  Description:    Implements Snarl Mail's message handling window
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Const TIMER_WAIT_FOR_OUTLOOK = 1
Private Const TIMER_WAIT_FOR_QUIT = 2

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Const WM_NOTIFICATION = &H402
Private Const WM_TRAY = &H403

Dim WithEvents myOutlook As Outlook.Application
Attribute myOutlook.VB_VarHelpID = -1
Dim mContacts As Outlook.Items
Dim mhWnd As Long

Dim mToken As Long
Dim thePanel As BPrefsPanel
Attribute thePanel.VB_VarHelpID = -1
Dim mConfig As CConfFile2

Private Const CLASS_NORMAL = "norm"
Private Const CLASS_PERSONAL = "pers"
Private Const CLASS_PRIVATE = "priv"
Private Const CLASS_CONFIDENTIAL = "conf"

Dim mMail As BTagList
Dim mTrayIcon As BNotifyIcon
Dim mPassword As String

Implements BWndProcSink
Implements KPrefsPanel
Implements KPrefsPage

Private Function BWndProcSink_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim pOutlook As Outlook.Application
Dim pItem As TItem
Dim sz As String
Dim n As Long

    On Error Resume Next

    Select Case uMsg
    Case WM_CREATE
        Form1.Add "WM_CREATE"

        ' /* basic initialisation */

        mhWnd = hwnd
        Set mMail = new_BTagList()
        mPassword = snCreatePassword()

        Set mTrayIcon = New BNotifyIcon
        mTrayIcon.SetTo mhWnd, WM_TRAY

        Set mConfig = New CConfFile2
        If g_GetSystemFolder(CSIDL_APPDATA, sz) Then
            sz = g_MakePath(sz) & "full phat\SnarlMail\"
            mConfig.SetTo sz & "SnarlMail.conf"

        Else
            g_Debug "TWindow.WM_CREATE(): %appdata% path error", LEMON_LEVEL_CRITICAL

        End If

        ' /* build config/defaults */

        With mConfig
            .AddIfMissing "show_body", "1"
            .AddIfMissing "use_contact_icon", "1"
            .AddIfMissing "open_on_click", "1"
            .AddIfMissing "auto_mark_as_read", "1"
            .AddIfMissing "show_tray_icon", "1"
'            .AddIfMissing "use_oo", "0"

            If Not .Save() Then _
                g_Debug "TWindow.WM_CREATE(): error writing config", LEMON_LEVEL_WARNING

        End With

        If mConfig.ValueOf("use_oo") = "1" Then _
            uEnableOO True

        uRegister           ' // might as well try now...

        uShowTrayIcon (mConfig.ValueOf("show_tray_icon") = "1")

        If Not uGetOutlook(pOutlook) Then
            SetTimer hwnd, TIMER_WAIT_FOR_OUTLOOK, 5000, 0

        Else
            Form1.Add "got Outlook..."
            Set myOutlook = pOutlook
            Set mContacts = myOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Items

        End If


    Case WM_DESTROY
        ' /* we're disappearing... */
        Form1.Add "WM_DESTROY"
        snarl_unregister App.ProductName, mPassword
        
        If mConfig.ValueOf("use_oo") = "1" Then _
            uEnableOO False

        Set myOutlook = Nothing
        Set mContacts = Nothing
        Set mTrayIcon = Nothing

        If Not (thePanel Is Nothing) Then _
            thePanel.Quit


    Case WM_CLOSE
        PostQuitMessage 0


    Case WM_TIMER
        Select Case wParam
        Case TIMER_WAIT_FOR_OUTLOOK
            If uGetOutlook(pOutlook) Then
                Set myOutlook = pOutlook
                Set mContacts = myOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Items
                Form1.Add "got Outlook"
                KillTimer hwnd, wParam

            End If

        Case TIMER_WAIT_FOR_QUIT
            Form1.Add "settle timeout completed"
            KillTimer hwnd, wParam
            SetTimer hwnd, TIMER_WAIT_FOR_OUTLOOK, 5000, 0

        End Select


    Case snBroadcastMsg()
        Select Case wParam
        Case SNARL_BROADCAST_LAUNCHED
            Form1.Add "[SNARL_BROADCAST_LAUNCHED]"
            uRegister

'        Case SNARL_BROADCAST_QUIT
'            g_Debug "[SNARL_BROADCAST_QUIT]"
'            Form1.Add "Snarl quit"
'            uTidyUp

        End Select


    Case snAppMsg()
        Select Case wParam
        Case SNARLAPP_DO_PREFS
            Form1.Add "[SNARLAPP_DO_PREFS]"
            uDoPrefs

        Case SNARLAPP_DO_ABOUT
            Form1.Add "[SNARLAPP_DO_ABOUT]"
            If uRegister() Then _
                Debug.Print snDoRequest("notify?app-sig=" & App.ProductName & _
                            "&uid=__about" & _
                            "&password=" & mPassword & _
                            "&title=SnarlMail " & CStr(App.Major) & "." & CStr(App.Minor) & IIf(App.Comments <> "", " " & App.Comments, "") & " (Build " & CStr(App.Revision) & ")" & _
                            "&text=" & App.LegalCopyright & vbCrLf & "Visit http://www.fullphat.net for support and updates" & uOutlookVer() & _
                            "&icon=" & g_MakePath(App.Path) & "icon.png" & _
                            "&action=Visit Website,http://www.fullphat.net")

        Case SNARLAPP_ACTIVATED
            ' /* V43 */
            Form1.Add "[SNARLAPP_ACTIVATED]"
            uDoPrefs

        Case SNARLAPP_QUIT_REQUESTED
            ' /* V43 */
            Form1.Add "[SNARLAPP_QUIT_REQUESTED]"
            PostQuitMessage 0


        End Select


    Case WM_TRAY
        If lParam = WM_LBUTTONDBLCLK Then _
            uDoPrefs


    Case WM_NOTIFICATION
        n = LoWord(wParam)
        Select Case n
        Case SNARL_NOTIFY_ACTION
            Form1.Add "** actioned " & CStr(lParam) & " **"
            If mMail.Find(CStr(lParam), pItem) Then _
                pItem.DoAction HiWord(wParam), (mConfig.ValueOf("auto_mark_as_read") = "1")

        Case SNARL_NOTIFY_INVOKED, SNARL_CALLBACK_INVOKED
            Form1.Add "** clicked " & CStr(lParam) & " **"
            If mMail.Find(CStr(lParam), pItem) Then
                If mConfig.ValueOf("open_on_click") = "1" Then _
                    pItem.DoClicked

                If mConfig.ValueOf("auto_mark_as_read") = "1" Then _
                    pItem.MarkAsRead

            End If

        End Select

        mMail.Remove mMail.IndexOf(CStr(lParam))


    Case WM_TEST
        If Not (myOutlook Is Nothing) Then
            Form1.Add "TEST: " & wParam
            Screen.MousePointer = vbArrowHourglass

            Select Case wParam
            Case 0
                uNotify myOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items.GetFirst()

            Case 1
                uNotify uFirstMeetingReq()

            Case 2
                uListInbox

            End Select
    
            Screen.MousePointer = vbArrow
    
        Else
            Form1.Add "not got Outlook"

        End If

    End Select

End Function

Private Sub Class_Initialize()

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    If Control.GetName = "toolbar" Then
        Select Case Val(Value)
        Case 1
            SendMessage mhWnd, WM_TEST, 0, ByVal 0&

        Case 2
            PostMessage mhWnd, snAppMsg(), SNARLAPP_DO_ABOUT, ByVal 0&

        Case 4
            PostQuitMessage 0

        End Select

    Else
        mConfig.Update Control.GetName, Value
        mConfig.Save
        
        If Control.GetName = "show_tray_icon" Then
            uShowTrayIcon (Value = "1")

        End If

    End If

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName
    Case "go_support"
        ShellExecute thePanel.hwnd, "open", "http://groups.google.com/group/snarl-discuss/?pli=1", vbNullString, vbNullString, 5

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

Private Function KPrefsPage_hwnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set thePanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub myOutlook_NewMailEx(ByVal EntryIDCollection As String)
Dim itemID() As String
Dim oItem As Object
Dim i As Long

    On Error Resume Next

    g_Debug "[myOutlook::NewMailEx]"

    With myOutlook.Session
        Form1.Add "** Mail received **"
        itemID = Split(EntryIDCollection, ",")
        For i = 0 To UBound(itemID)
            err.Clear
            Set oItem = .GetItemFromID(itemID(i))
'            Form1.Add "Item #" & CStr(i) & " entry id=" & ItemId(i) & " " & err.Description
            uNotify oItem

        Next i

    End With

End Sub

Private Sub myOutlook_Quit()

    Form1.Add "[myOutlook::Quit] fired"
    Set myOutlook = Nothing
    Set mContacts = Nothing

    Form1.Add "waiting for Outlook to finish closing..."
    SetTimer mhWnd, TIMER_WAIT_FOR_QUIT, 5000, 0

End Sub

Private Function uGetContactIcon(ByVal EmailAddress As String, ByRef Path As String, ByRef FriendlyName As String) As Boolean
Dim szPath As String

    g_Debug "uGetContactIcon(): " & EmailAddress

    szPath = g_GetTempPath(True) & g_MakeFilename2(EmailAddress) & ".png"
    DeleteFile szPath

Dim pOutlook As Outlook.Application

    If Not uGetOutlook(pOutlook) Then _
        Exit Function

    g_Debug "uGetContactIcon(): accquired Outlook..."

Dim pContact As ContactItem

    If Not uFindContactByEmail(EmailAddress, pContact) Then
        g_Debug "uGetContactIcon(): '" & EmailAddress & "' not in contacts"
        Exit Function

    End If

    FriendlyName = pContact.FullName

    If Not pContact.HasPicture Then
        g_Debug "uGetContactIcon(): '" & EmailAddress & "' has no picture associated"
        Exit Function

    End If

    If pContact.Attachments.Count = 0 Then _
        Exit Function

Dim i As Long

    With pContact.Attachments
        For i = 1 To .Count
            If LCase$(.Item(i).DisplayName) = "contactpicture.jpg" Then
                .Item(i).SaveAsFile szPath
                Path = szPath
                uGetContactIcon = True
                Exit Function

            End If
        Next i

    End With

End Function

Private Function uFindContactByEmail(ByVal Email As String, ByRef Contact As ContactItem) As Boolean

    If (mContacts Is Nothing) Then _
        Exit Function

    If mContacts.Count = 0 Then _
        Exit Function

    Email = LCase$(Email)
    g_Debug "uFindContactByEmail(): searching contacts..."

Dim pContact As ContactItem
Dim pObj As Object
Dim i As Long

    With mContacts
        Set pObj = .GetFirst()
        Do While Not (pObj Is Nothing)
            If TypeOf pObj Is ContactItem Then
                Set pContact = pObj
                If (LCase$(pContact.Email1Address) = Email) Or (LCase$(pContact.Email2Address) = Email) Or (LCase$(pContact.Email3Address) = Email) Then
                    g_Debug "uFindContactByEmail(): found contact '" & pContact.FullName & "'"
                    Set Contact = pContact
                    uFindContactByEmail = True
                    Exit Function
    
                End If
    
            Else
                g_Debug "uFindContactByEmail(): " & TypeName(pObj)

            End If

            Set pObj = .GetNext()

        Loop

    End With

    g_Debug "uFindContactByEmail(): not found"

End Function

Private Function uGetSMTPAddress(ByRef Mail As Object) As String
'Const PR_SMTP_ADDRESS = 972947486
'Dim pMsg As MailItem
'Dim pAdr As Recipient

    On Error Resume Next

'    Debug.Print "uGetSMTPAddress(): " & Mail.SenderEmailType & " " & Mail.SenderEmailAddress

    If Mail.SenderEmailType = "EX" Then
        ' /* X400 - code from here: http://anoriginalidea.wordpress.com/2008/01/11/getting-the-smtp-email-address-of-an-exchange-sender-of-a-mailitem-from-outlook-in-vbnet-vsto/ */

'        Set pMsg = Mail.Application.CreateItem(olMailItem)
'        Set pAdr = pMsg.Recipients.Add(Mail.Name)
'        pAdr.Resolve
'        uGetSMTPAddress = getmapiproperty(pAdr.AddressEntry.MAPIOBJECT, PR_SMTP_ADDRESS)
        uGetSMTPAddress = Mail.SenderEmailAddress

    Else
        ' /* assume already SMTP */
        uGetSMTPAddress = Mail.SenderEmailAddress

    End If


'NameSpace.GetRecipientFromID

'Dim oSession As MAPI.Session
'
'    Err.Clear
'    Set oSession = CreateObject("MAPI.Session")
'    If (Err.Number <> 0) Or (oSession Is Nothing) Then
'        g_Debug "uGetSMTPAddress(): couldn't create MAPI.Session", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End If
'
''strEMailAddress = objAddressEntry.Address
''
''' Check if it is an Exchange object
''If Left(strEMailAddress, 3) = "/o=" Then
''
''  ' Get the SMTP address
''  strAddressEntryID = objAddressEntry.Id
''  strEMailAddress =_
''    objSession.GetAddressEntry(strAddressEntryID).Fields(CdoPR_EMAIL).Value
''End If
''
''' Display the SMTP address of current user
''MsgBox "SMTP address of current user: " & strEMailAddress

End Function

Private Sub uDoPrefs()
Dim pPage As BPrefsPage

    If (thePanel Is Nothing) Then
        Set thePanel = New BPrefsPanel
        With thePanel
            .SetHandler Me
            .SetTitle "SnarlMail Preferences"
            .SetWidth 420

            Set pPage = new_BPrefsPage("", , Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("banner", "", "Appearance")
                .Add new_BPrefsControl("fancytoggle2", "show_body", "Include message summary in notification?", , mConfig.ValueOf("show_body"))
                .Add new_BPrefsControl("label", "", "The summary will only be included for non-sensitive emails.")
'                .Add new_BPrefsControl("fancytoggle2", "obscure_sensitive", "Obscure body content if item is marked as sensitive?", , mConfig.GetValue("obscure_sensitive"))
'                .Add new_BPrefsControl("fancytoggle2", "extra_icons", "Use forward/reply icons?", , mConfig.GetValue("extra_icons"))
                .Add new_BPrefsControl("fancytoggle2", "use_contact_icon", "Use contact picture (if available)?", , mConfig.ValueOf("use_contact_icon"))

                .Add new_BPrefsControl("fancytoggle2", "show_tray_icon", "Show icon in System Tray?", , mConfig.ValueOf("show_tray_icon"))

                .Add new_BPrefsControl("banner", "", "Behaviour")
                .Add new_BPrefsControl("label", "", "When the notification is clicked:")
                .Add new_BPrefsControl("fancytoggle2", "open_on_click", "Open item?", , mConfig.ValueOf("open_on_click"))
                .Add new_BPrefsControl("fancytoggle2", "auto_mark_as_read", "Mark item as read?", , mConfig.ValueOf("auto_mark_as_read"))


                .Add new_BPrefsControl("banner", "", "Support")
                .Add new_BPrefsControl("label", "", "SnarlMail currently only supports Microsoft® Outlook®.  It has been developed around Outlook® 2007 but should work correctly with Outlook® 2003 and other newer versions.  If you experience problems with SnarlMail, or want to suggest a new feature, please raise a ticket on the Snarl support group.")
                .Add new_BPrefsControl("fancybutton2", "go_support", "Support Group")

'                .Add new_BPrefsControl("label", "", "something interesting here:")
'
'                Set pC = new_BPrefsControl("fancyedit", "", "2048", "Limit:")
'                pC.Indent 4
'                pC.SizeTo 48, -1
'                .Add pC
'
'                Set pC = new_BPrefsControl("fancyedit", "", "2048", "Limit:")
'                .Add pC

'                .Add new_BPrefsControl("fancybutton2", "about", "About SnarlMail")
'                .Add new_BPrefsSeparator()

                .Add new_BPrefsSeparator()
                .Add new_BPrefsControl("fancytoolbar", "toolbar", "Test Settings|About SnarlMail||Quit SnarlMail")

'                .Add new_BPrefsSeparator()
'                .Add new_BPrefsControl("fancybutton2", "quit", "Quit SnarlMail")

            End With

            .AddPage pPage

            .Go

            g_SetWindowIconToAppResourceIcon .hwnd

        End With
    End If

    g_WindowToFront thePanel.hwnd, True


End Sub

Private Function uGetOutlook(ByRef pOutlook As Outlook.Application) As Boolean

    On Error Resume Next

    err.Clear
    Set pOutlook = GetObject(, "outlook.application")
    If err.Number <> 0 Then
        Debug.Print "TWindow.uGetOutlook(): " & err.Description, LEMON_LEVEL_CRITICAL

    Else
        uGetOutlook = True

    End If

End Function

Private Function uOutlookVer() As String

    If Not (myOutlook Is Nothing) Then _
        uOutlookVer = vbCrLf & vbCrLf & "Outlook " & myOutlook.Version & " detected"

End Function

Private Function uNotify(ByRef Item As Object) As Boolean

    On Error Resume Next

    If (Item Is Nothing) Then _
        Exit Function

    g_Debug "uNotify(): TypeOf()=" & TypeName(Item)

    ' /* register */

    If Not uRegister() Then _
        Exit Function

    g_Debug "uNotify(): was registered ok"

    ' /* determine item type */

    If TypeOf Item Is MailItem Then
        Form1.Add "uNotify(): Is MailItem"
        uNotifyMail Item

    ElseIf TypeOf Item Is MeetingItem Then
        Form1.Add "uNotify(): Is MeetingItem"
        uNotifyMeeting Item

    Else
        Form1.Add "uNotify(): not supported"
        g_Debug "uNotify(): item type '" & TypeName(Item) & "' not supported", LEMON_LEVEL_CRITICAL

    End If

End Function

Private Function uNotifyMail(ByRef Mail As MailItem) As Boolean
Dim bShowBody As Boolean
Dim szClass As String

    Form1.Add "uNotifyMail(): subject='" & Mail.Subject & "' sen=" & CStr(Mail.Sensitivity) & " pri=" & CStr(Mail.Importance)

    Select Case Mail.Sensitivity
    Case olPersonal
        szClass = CLASS_PERSONAL

    Case olPrivate
        szClass = CLASS_PRIVATE

    Case olConfidential
        szClass = CLASS_CONFIDENTIAL

    Case Else
        szClass = CLASS_NORMAL
        bShowBody = True

    End Select

Dim szBody As String

    ' /* get body summary */

    Debug.Print "getting mail body..."

    If (bShowBody) And (mConfig.ValueOf("show_body") = "1") Then
        szBody = g_SafeLeftStr(Mail.Body, 128, True)
        ' /* do some tidying up */
        szBody = Replace$(szBody, vbCrLf & vbCrLf & Chr$(&HA0) & vbCrLf & vbCrLf, " ")
        szBody = Replace$(szBody, vbCrLf & vbCrLf, " ")

    End If

Dim szIcon As String
Dim szFrom As String

    ' /* sender */

    Debug.Print "getting sender name..."
    szFrom = Mail.SenderName
    Debug.Print "sender name: " & szFrom

    ' /* get the sender's contact picture */

    If mConfig.ValueOf("use_contact_icon") = "1" Then _
        uGetContactIcon uGetSMTPAddress(Mail), szIcon, szFrom

    ' /* set content based on message class */

    Select Case Mail.MessageClass
    Case "IPM.Note"
        If szIcon = "" Then _
            szIcon = "!message-new_mail"

    Case "IPM.Note.Rules.OofTemplate.Microsoft"
        If szIcon = "" Then _
            szIcon = "!message-reply-ooo"

    Case "IPM.Outlook.Recall"
        If szIcon = "" Then _
            szIcon = "!message-recall"

    Case Else
        Form1.Add "uNotifyMail(): spurious message class '" & Mail.MessageClass & "'"
        If szIcon = "" Then _
            szIcon = "!message-new_mail"

    End Select

    Form1.Add "uNotifyMail(): uuid=" & Mail.EntryID

Dim hr As Long

    hr = snDoRequest("notify?app-sig=" & App.ProductName & _
                    "&id=" & szClass & _
                    "&uid=" & Mail.EntryID & _
                    "&title=" & szFrom & ": " & Mail.Subject & _
                    "&text=" & szBody & _
                    "&icon=" & szIcon & _
                    "&action=Open,@9&action=Reply,@2&action=Reply All,@3&action=Forward,@4" & _
                    "&priority=" & CStr(Mail.Importance - 1) & _
                    "&password=" & mPassword)

Dim pm As TItem

    Set pm = New TItem
    pm.SetTo Mail, hr
    mMail.Add pm

End Function

Private Function uNotifyMeeting(ByRef Meeting As MeetingItem) As Boolean
Dim bShowBody As Boolean
Dim szClass As String

    Select Case Meeting.Sensitivity
    Case olPersonal
        szClass = CLASS_PERSONAL

    Case olPrivate
        szClass = CLASS_PRIVATE

    Case olConfidential
        szClass = CLASS_CONFIDENTIAL

    Case Else
        szClass = CLASS_NORMAL
        bShowBody = True

    End Select

Dim pAppt As AppointmentItem
Dim szTitle As String
Dim szText As String
Dim szActions As String

    Set pAppt = Meeting.GetAssociatedAppointment(True)

    If (pAppt Is Nothing) Then
        ' /* no associated appointment so probably an accept/decline/tentative */
        Form1.Add "uNotifyMeeting(): no associated appointment"
        g_Debug "uNotifyMeeting(): class is '" & Meeting.MessageClass & "'"

        szTitle = Meeting.Subject
        szText = uTidyup(Meeting.Body)
        szActions = "Open,@9&action=Reply,@2&action=Reply All,@3&action=Forward,@4"

'        szTitle = szFrom & ": " & Meeting.Subject

    Else
        ' /* get body summary */

        Form1.Add "uNotifyMeeting(): has associated appointment"

        If mConfig.ValueOf("show_body") <> "1" Then _
            bShowBody = False

        uGetAppointmentInfo Meeting, pAppt, bShowBody, szTitle, szText ', szActions

    End If

    ' /* sender and icon */

Dim szIcon As String
Dim szFrom As String

    szFrom = Meeting.SenderName

    If mConfig.ValueOf("use_contact_icon") = "1" Then _
        uGetContactIcon uGetSMTPAddress(Meeting), szIcon, szFrom        ' // get contact's picture

    szTitle = szFrom

    szActions = "Open,@9&action=Reply,@2&action=Reply All,@3&action=Forward,@4"

    ' /* set based on message class */

    Select Case Meeting.MessageClass
    Case "IPM.Schedule.Meeting.Resp.Neg"
        szTitle = szTitle & " declined"
        If szIcon = "" Then _
            szIcon = "!message-appt-decline"

    Case "IPM.Schedule.Meeting.Resp.Pos"
        szTitle = szTitle & " accepted"
        If szIcon = "" Then _
            szIcon = "!message-appt-accept"

    Case "IPM.Schedule.Meeting.Resp.Tent"
        szTitle = szTitle & " is undecided"
        If szIcon = "" Then _
            szIcon = "!message-appt-tentative"

    Case "IPM.Schedule.Meeting.Canceled"
        If szIcon = "" Then _
            szIcon = "!message-appt-cancelled"

    Case "IPM.Schedule.Meeting.Request"
        szActions = "Open,@9&action=Accept,@12&action=Tentative,@13&action=Decline,@14"
        If szIcon = "" Then _
            szIcon = "!message-appt-new"

    Case Else
        g_Debug "uNotifyMeeting(): class is '" & Meeting.MessageClass & "'"
        szActions = "Open,@9"
        If szIcon = "" Then _
            szIcon = "!message-appt-new"

    End Select

Dim hr As Long

    hr = snDoRequest("notify?app-sig=" & App.ProductName & _
                     "&id=" & szClass & _
                     "&uid=" & Meeting.EntryID & _
                     "&title=" & szTitle & _
                     "&text=" & szText & _
                     "&icon=" & szIcon & _
                     "&action=" & szActions & _
                     "&priority=" & CStr(Meeting.Importance - 1) & _
                     "&password=" & mPassword)

Dim pm As TItem

    Set pm = New TItem
    pm.SetTo Meeting, hr
    mMail.Add pm

End Function

Private Function uFirstMeetingReq() As MeetingItem
Dim pItem As Object

    With myOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items
        Set pItem = .GetFirst()
        Do While Not (pItem Is Nothing)
            If TypeOf pItem Is MeetingItem Then
                Set uFirstMeetingReq = pItem
                g_Debug "uFirstMeetingReq(): found '" & uFirstMeetingReq.Subject & "'"
                Exit Function

            End If

            Set pItem = .GetNext()

        Loop

    End With

    Form1.Add "uFirstMeetingReq(): no suitable item found"

End Function

Private Function uDate(ByVal d As Date, Optional ByVal DayOnly As Boolean) As String

    If DateDiff("d", Now, d) = 0 Then
        If DayOnly Then
            uDate = "Today"

        Else
            uDate = "Today " & Format$(d, " hh:mm")

        End If

    ElseIf DateDiff("d", Now, d) = 1 Then
        If DayOnly Then
            uDate = "Tomorrow"

        Else
            uDate = "Tomorrow " & Format$(d, " hh:mm")

        End If

    Else
        If DayOnly Then
            uDate = Format$(d, "ddd d mmm")

        Else
            uDate = Format$(d, "ddd d mmm hh:mm")

        End If

    End If

End Function

Private Function uRegister() As Boolean
Dim hr As Long

    hr = snDoRequest("register?app-sig=" & App.ProductName & _
                     "&app-title=" & App.Title & _
                     "&icon=" & g_MakePath(App.Path) & "icon.png" & _
                     "&reply-to=" & CStr(mhWnd) & _
                     "&reply-with=" & CStr(WM_NOTIFICATION) & _
                     "&app-daemon=1" & _
                     "&password=" & mPassword)

    If hr > 0 Then

'    If snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", , _
                      mhWnd, WM_NOTIFICATION, _
                      SNARLAPP_IS_WINDOWLESS Or SNARLAPP_HAS_ABOUT Or SNARLAPP_HAS_PREFS) > 0 Then

        ' /* add classes */

        snDoRequest "addclass?app-sig=" & App.ProductName & _
                    "&id=" & CLASS_NORMAL & _
                    "&name=Normal messages" & _
                    "&password=" & mPassword

        snDoRequest "addclass?app-sig=" & App.ProductName & _
                    "&id=" & CLASS_PERSONAL & _
                    "&name=Personal messages" & _
                    "&password=" & mPassword

        snDoRequest "addclass?app-sig=" & App.ProductName & _
                    "&id=" & CLASS_PRIVATE & _
                    "&name=Private messages" & _
                    "&password=" & mPassword

        snDoRequest "addclass?app-sig=" & App.ProductName & _
                    "&id=" & CLASS_CONFIDENTIAL & _
                    "&name=Confidential messages" & _
                    "&password=" & mPassword

        uRegister = True

    Else
        g_Debug "TWindow.uRegister(): failed (" & CStr(hr) & ")"

    End If

End Function

Private Sub uShowTrayIcon(ByVal Show As Boolean)
Dim h As Long

    If Show Then
        If Not g_IsIDE Then _
            h = g_GetResourceIcon(1)

        mTrayIcon.Add "app", IIf(h, h, Form1.Icon.Handle), "SnarlMail"

    Else
        mTrayIcon.Remove "app"

    End If

End Sub


















''Dim i As Integer
''
''        For i = 1 To Len(szBody)
''            Debug.Print Mid$(szBody, i, 1) & "-" & Asc(Mid$(szBody, i, 1)) & " ";
''
''        Next i


'Private Function uCreateEmblem(ByVal Text As String) As mfxBitmap
'
'    If Len(Text) > 1 Then _
'        Text = "+"
'
'    With New mfxView
'        .SizeTo 64, 64
'        .EnableSmoothing True
'        .SetHighColour rgba(255, 0, 0)
'        .FillEllipse .Bounds
'        .SetHighColour rgba(0, 0, 0)
'        .StrokeEllipse .Bounds
'        .SetHighColour rgba(255, 255, 255)
'        .SetFont "Arial", 36, True
'        .DrawString Text, .Bounds, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER
'        Set uCreateEmblem = .ConvertToBitmap()
'
'    End With
'
'End Function

'Private Function uFindFirst(ByVal ItemClass As OlObjectClass, ByRef Item As Object, Optional ByVal MaxSearchCount As Long = 32) As Boolean
'
'    If (myItems Is Nothing) Then _
'        Exit Function
'
'Dim i As Long
'
'    With myItems
'        For i = 1 To MIN(MaxSearchCount, .Count)
'            Debug.Print .Item(i).Subject & " (" & .Item(i).Class & ")"
'            If .Item(i).Class = ItemClass Then
'                Set Item = .Item(i)
'                uFindFirst = True
'                Exit Function
'
'            End If
'        Next i
'    End With
'
'End Function

'Private Sub uBuildIcon(ByVal Filename As String, ByVal Priority As OlImportance, ByVal Sensitivity As OlSensitivity, Optional ByVal Badge As String, Optional ByVal AttachmentCount As Long)
'Dim pbm As mfxBitmap
'Dim cxIcon As Long
'Dim sz As String
'Dim pt As BPoint
'
'    Set pbm = load_image_obj(Filename)
'    If Not is_valid_image(pbm) Then _
'        Set pbm = uCreateFakeImage()
'
'    Set pbm = uMakeSquareImage(pbm)
'
'    cxIcon = Fix(pbm.Width / 3)
'
'    With New mfxView
'        .SizeTo pbm.Width, pbm.Height
'        .DrawScaledImage pbm
'
'        Set pt = new_BPoint(0, .Height - cxIcon - 1)
'
'        ' /* priority emblem */
'        If Priority = olImportanceHigh Then
'            .DrawScaledImage uCreateEmblem("!"), pt, new_BPoint(cxIcon, cxIcon)
'            pt.OffsetBy cxIcon, 0
'
'        End If
'
'        If AttachmentCount > 0 Then
'            .DrawScaledImage uCreateEmblem(CStr(AttachmentCount)), pt, new_BPoint(cxIcon, cxIcon)
'            pt.OffsetBy cxIcon, 0
'
'        End If
'
'        ' /* badge (only one) top-right */
'
''Dim szBadge As String
''
''        Select Case Acceptance
''        Case olMeetingAccepted
''            szBadge = "emblem-accept"
''
''        Case olMeetingDeclined
''            szBadge = "emblem-decline"
''
''        Case olMeetingTentative
''            szBadge = "emblem-tentative"
''
'''        Case Else
'''            If IsMeeting Then _
'''                szBadge = "emblem-meeting"
''
''        End Select
'
'        If Badge <> "" Then
'            cxIcon = cxIcon * 1.5
'            .DrawScaledImage load_image_obj(g_MakePath(App.Path) & Badge & ".png"), _
'                             new_BPoint(.Width - cxIcon, .Height - cxIcon), new_BPoint(cxIcon, cxIcon)
'
'        End If
'
'        .WriteToFile g_MakePath(App.Path) & ".temp_icon.png", "image/png"
'
'    End With
'
'End Sub

Private Sub uListInbox()

    g_Debug "uListInbox()", LEMON_LEVEL_PROC_ENTER
    
    If (myOutlook Is Nothing) Then _
        Exit Sub

Dim pe As MeetingItem
Dim pm As MailItem
Dim po As Object

    With myOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items
        Set po = .GetFirst()
        Do While Not (po Is Nothing)
            If TypeOf po Is MailItem Then
                Set pm = po
                If pm.MessageClass <> "IPM.Note" Then _
                    Debug.Print "Mailitem: " & pm.MessageClass

                ' /* message classes */

                '// "IPM.Note"
                '// "IPM.Note.Rules.OofTemplate.Microsoft"
                '// "IPM.Note.Microsoft.Missed.Voice"

            ElseIf TypeOf po Is MeetingItem Then
                Set pe = po
                Debug.Print "MeetingItem: " & pe.MessageClass

                ' /* message classes */

                '// "IPM.Schedule.Meeting.Request"
                '// "IPM.Schedule.Meeting.Resp.Pos"

            Else
                g_Debug TypeName(po)

            End If

            Set po = .GetNext()

        Loop

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub uGetAppointmentInfo(ByRef pMeet As MeetingItem, ByRef pAppt As AppointmentItem, ByVal ShowBody As Boolean, ByRef Title As String, ByRef Text As String) ', ByRef Actions As String)
Dim sz As String

    sz = pAppt.Subject
    If sz = "" Then _
        sz = pMeet.Subject

    Form1.Add "uGetAppointmenuInfo(): subject=" & sz & " start=" & pAppt.Start & " end=" & pAppt.End

    ' /* subject */

    If sz <> "" Then _
        Text = Text & sz & vbCrLf

    ' /* time */

Dim dStart As Date
Dim dEnd As Date

    dStart = pAppt.Start
    dEnd = pAppt.End

    If pAppt.AllDayEvent Then
        ' /* all day */
        If DateDiff("h", dStart, dEnd) > 24 Then
            sz = uDate(dStart, True) & "-" & uDate(dEnd, True) & " (all day)"

        Else
            ' /* same day */
            sz = uDate(dStart, True) & " (all day)"

        End If

    Else
        ' /* not all day */
        If DateDiff("d", dStart, dEnd) = 0 Then
            ' /* same day */
            sz = uDate(dStart) & "-" & Format$(dEnd, "hh:mm")

        Else
            sz = uDate(dStart) & "-" & Format$(dEnd, "ddd d mmm hh:mm")

        End If

    End If

    Text = Text & sz & vbCrLf

    ' /* location */

    If pAppt.Location <> "" Then _
        Text = Text & pAppt.Location & vbCrLf

    ' /* body */

    If ShowBody Then
        sz = uTidyup(pAppt.Body)
        If sz <> "" Then _
            Text = Text & IIf(Text <> "", "--" & vbCrLf, "") & sz & vbCrLf

    End If

    Text = g_SafeLeftStr(Text, Len(Text) - 2)

'    Actions = "Open,@9&action=Accept,@12&action=Tentative,@13&action=Decline,@14"

End Sub

Private Function uTidyup(ByVal BodyText As String) As String

    uTidyup = g_SafeLeftStr(BodyText, 128, True)
    uTidyup = Replace$(uTidyup, vbCrLf & vbCrLf & Chr$(&HA0) & vbCrLf & vbCrLf, " ")
    uTidyup = Replace$(uTidyup, vbCrLf, " ")
    uTidyup = trim(uTidyup)
'    uTidyup = Replace$(uTidyup, vbCrLf & vbCrLf, " ")
'    uTidyup = vbCrLf & vbCrLf & uTidyup

End Function

Private Sub uEnableOO(ByVal Enabled As Boolean)

    If Enabled Then
        g_Debug "TWindow.uEnableOO(): launching oo..."
        ShellExecute mhWnd, "open", g_MakePath(App.Path) & "outlookoutsmarter.exe", vbNullString, vbNullString, SW_HIDE

    Else
        g_Debug "TWindow.uEnableOO(): quitting oo..."
        ShellExecute mhWnd, "open", g_MakePath(App.Path) & "outlookoutsmarter.exe", "-quit", vbNullString, SW_HIDE
    
    End If

End Sub
