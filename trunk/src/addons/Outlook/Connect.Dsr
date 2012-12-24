VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7005
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   15270
   _ExtentX        =   26935
   _ExtentY        =   12356
   _Version        =   393216
   Description     =   "Displays incoming emails as a Snarl notification."
   DisplayName     =   "Snarl Notifier 2.0 for Microsoft Outlook"
   AppName         =   "Microsoft Outlook"
   AppVer          =   "Microsoft Outlook 12.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME = "w>snOutlook"
Private Const WM_NOTIFICATION = &H402

Private Const CLASS_NORMAL = "norm"
Private Const CLASS_PERSONAL = "pers"
Private Const CLASS_PRIVATE = "priv"
Private Const CLASS_CONFIDENTIAL = "conf"

Dim mConfig As CConfFile2
Dim mPassword As String
Dim mMail As BTagList
Dim mhWnd As Long

Private WithEvents pOLApp As Outlook.Application
Attribute pOLApp.VB_VarHelpID = -1

Implements BWndProcSink

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    On Error Resume Next

Dim sz As String

    Set pOLApp = Application
    mPassword = create_password()
    Set mMail = new_BTagList()
'    Set mContacts = pOLApp.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Items

    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me, CLASS_NAME)

    Set mConfig = New CConfFile2
    If g_GetSystemFolder(CSIDL_APPDATA, sz) Then
'        sz = g_MakePath(sz) & "full phat\SnarlMail\"
'        mConfig.SetTo sz & "SnarlMail.conf"

    Else
        g_Debug "TWindow.WM_CREATE(): %appdata% path error", LEMON_LEVEL_CRITICAL

    End If

    ' /* 1.01: try to register now */
    
    uRegister

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    Set pOLApp = Nothing
    snarl_unregister App.ProductName, mPassword

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim pItem As TItem
Dim n As Long

    Select Case uMsg

    Case WM_NOTIFICATION
        n = LoWord(wParam)
        Select Case n
        Case SNARL_NOTIFY_ACTION
'            Form1.Add "** actioned " & CStr(lParam) & " **"
            If mMail.Find(CStr(lParam), pItem) Then _
                pItem.DoAction HiWord(wParam), True         '// (mConfig.ValueOf("auto_mark_as_read") = "1")

        Case SNARL_NOTIFY_INVOKED, SNARL_CALLBACK_INVOKED
            If mMail.Find(CStr(lParam), pItem) Then _
                pItem.DoClicked

'        Case SNARL_NOTIFY_INVOKED, SNARL_CALLBACK_INVOKED
''            Form1.Add "** clicked " & CStr(lParam) & " **"
'            If mMail.Find(CStr(lParam), pItem) Then
'                If mConfig.ValueOf("open_on_click") = "1" Then _
'                    pItem.DoClicked
'
'                If mConfig.ValueOf("auto_mark_as_read") = "1" Then _
'                    pItem.MarkAsRead
'
'            End If

        End Select

        mMail.Remove mMail.IndexOf(CStr(lParam))


'    Case WM_TEST
'        If Not (myOutlook Is Nothing) Then
'            Form1.Add "TEST: " & wParam
'            Screen.MousePointer = vbArrowHourglass
'
'            Select Case wParam
'            Case 0
'                uNotify myOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items.GetFirst()
'
'            Case 1
'                uNotify uFirstMeetingReq()
'
'            Case 2
'                uListInbox
'
'            End Select
'
'            Screen.MousePointer = vbArrow
'
'        End If

    End Select

End Function

Private Sub pOLApp_NewMailEx(ByVal EntryIDCollection As String)
Dim itemID() As String
Dim oItem As Object
Dim i As Long

    On Error Resume Next

    g_Debug "[pOLApp::NewMailEx]"

    With pOLApp.Session
'        Form1.Add "** Mail received **"
        itemID = Split(EntryIDCollection, ",")
        For i = 0 To UBound(itemID)
            Err.Clear
            Set oItem = .GetItemFromID(itemID(i))
'            Form1.Add "Item #" & CStr(i) & " entry id=" & ItemId(i) & " " & err.Description
            uNotify oItem

        Next i

    End With

End Sub

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
'        Form1.Add "uNotify(): Is MailItem"
        uNotifyMail Item

    ElseIf TypeOf Item Is MeetingItem Then
'        Form1.Add "uNotify(): Is MeetingItem"
        uNotifyMeeting Item

    Else
'        Form1.Add "uNotify(): not supported"
        g_Debug "uNotify(): item type '" & TypeName(Item) & "' not supported", LEMON_LEVEL_CRITICAL

    End If

End Function

Private Function uNotifyMail(ByRef Mail As MailItem) As Boolean
Dim bShowBody As Boolean
Dim szClass As String

'    Form1.Add "uNotifyMail(): subject='" & Mail.Subject & "' sen=" & CStr(Mail.Sensitivity) & " pri=" & CStr(Mail.Importance)

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

    If (bShowBody) Then ' And (mConfig.ValueOf("show_body") = "1") Then
        
        If (snGetSystemFlags() And SNARL_SF_USER_AWAY) Then
            ' /* V43: if the user is away, don't send the body content */
            szBody = "..."

        Else
            szBody = uGetBodyText(Mail)

        End If
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
'        Form1.Add "uNotifyMail(): spurious message class '" & Mail.MessageClass & "'"
        If szIcon = "" Then _
            szIcon = "!message-new_mail"

    End Select

'    Form1.Add "uNotifyMail(): uuid=" & Mail.EntryID

Dim hr As Long

    hr = snDoRequest("notify?app-sig=" & App.ProductName & _
                    "&id=" & szClass & _
                    "&uid=" & Mail.EntryID & _
                    "&title=" & szFrom & ": " & Mail.Subject & _
                    "&text=" & szBody & _
                    "&icon=" & szIcon & _
                    "&action=Mark as Read,@1&action=Reply,@2&action=Reply All,@3&action=Forward,@4" & _
                    "&priority=" & CStr(Mail.Importance - 1) & _
                    "&sensitivity=" & CStr(Mail.Sensitivity * 16) & _
                    "&password=" & mPassword & _
                    "&callback=@9&callback-label=Read")

Dim pm As TItem

    Set pm = New TItem
    pm.SetTo Mail, hr, mPassword
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
'        Form1.Add "uNotifyMeeting(): no associated appointment"
        g_Debug "uNotifyMeeting(): class is '" & Meeting.MessageClass & "'"

'        szTitle = Meeting.Subject
        szText = uTidyup(Meeting.Body)
'        szActions = "Reply,@2&action=Reply All,@3&action=Forward,@4&action=Mark as Read,@1"
'        szTitle = szFrom & ": " & Meeting.Subject

    Else
        ' /* get body summary */

'        Form1.Add "uNotifyMeeting(): has associated appointment"

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

    ' /* default actions */
    szActions = "action=Mark as Read,@1&action=Reply,@2&action=Reply All,@3&action=Forward,@4"
'    szActions = "Open,@9&action=Reply,@2&action=Reply All,@3&action=Forward,@4"

    ' /* set based on message class */
    Select Case Meeting.MessageClass
    Case "IPM.Schedule.Meeting.Resp.Neg"
        ' /* meeting response: declined */
        szTitle = szTitle & " declined"
        If szIcon = "" Then _
            szIcon = "!message-appt-decline"

    Case "IPM.Schedule.Meeting.Resp.Pos"
        ' /* meeting response: accepted */
        szTitle = szTitle & " accepted"
        If szIcon = "" Then _
            szIcon = "!message-appt-accept"

    Case "IPM.Schedule.Meeting.Resp.Tent"
        ' /* meeting response: tentative */
        szTitle = szTitle & " is undecided"
        If szIcon = "" Then _
            szIcon = "!message-appt-tentative"

    Case "IPM.Schedule.Meeting.Canceled"
        ' /* meeting cancellation */
        If szIcon = "" Then _
            szIcon = "!message-appt-cancelled"

    Case "IPM.Schedule.Meeting.Request"
        ' /* meeting request */
        szActions = "action=Accept,@12&action=Tentative,@13&action=Decline,@14"
        If szIcon = "" Then _
            szIcon = "!message-appt-new"

    Case Else
        g_Debug "uNotifyMeeting(): class is '" & Meeting.MessageClass & "'"
        szActions = ""
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
                     IIf(szActions <> "", "&" & szActions, "") & _
                     "&priority=" & CStr(Meeting.Importance - 1) & _
                     "&sensitivity=" & CStr(Meeting.Sensitivity * 16) & _
                     "&password=" & mPassword & _
                     "&callback=@9&callback-label=View")

Dim pm As TItem

    Set pm = New TItem
    pm.SetTo Meeting, hr, mPassword
    mMail.Add pm

End Function

Private Function uRegister() As Boolean
Dim hr As Long

    hr = snarl_register(App.ProductName, App.Title, "", mPassword, mhWnd, WM_NOTIFICATION)

'    hr = snDoRequest("register?app-sig=" & App.ProductName & _
                     "&app-title=" & App.Title & _
                     "&reply-to=" & CStr(0) & _
                     "&reply-with=" & CStr(0) & _
                     "&app-daemon=1" & _
                     "&password=" & mPassword)

'                     "&icon=" & g_MakePath(App.Path) & "icon.png" & _

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

Private Function uGetContactIcon(ByVal EmailAddress As String, ByRef Path As String, ByRef FriendlyName As String) As Boolean
Dim szPath As String

    g_Debug "uGetContactIcon(): " & EmailAddress

    szPath = g_GetTempPath(True) & g_MakeFilename2(EmailAddress) & ".png"
    DeleteFile szPath

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

'    If (mContacts Is Nothing) Then _
'        Exit Function
'
'    If mContacts.Count = 0 Then _
'        Exit Function
'
'    Email = LCase$(Email)
'    g_Debug "uFindContactByEmail(): searching contacts..."
'
'Dim pContact As ContactItem
'Dim pObj As Object
'Dim i As Long
'
'    With mContacts
'        Set pObj = .GetFirst()
'        Do While Not (pObj Is Nothing)
'            If TypeOf pObj Is ContactItem Then
'                Set pContact = pObj
'                If (LCase$(pContact.Email1Address) = Email) Or (LCase$(pContact.Email2Address) = Email) Or (LCase$(pContact.Email3Address) = Email) Then
'                    g_Debug "uFindContactByEmail(): found contact '" & pContact.FullName & "'"
'                    Set Contact = pContact
'                    uFindContactByEmail = True
'                    Exit Function
'
'                End If
'
'            Else
'                g_Debug "uFindContactByEmail(): " & TypeName(pObj)
'
'            End If
'
'            Set pObj = .GetNext()
'
'        Loop
'
'    End With
'
'    g_Debug "uFindContactByEmail(): not found"

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

Private Function uTidyup(ByVal BodyText As String) As String

    uTidyup = g_SafeLeftStr(BodyText, 128, True)
    uTidyup = Replace$(uTidyup, vbCrLf & vbCrLf & Chr$(&HA0) & vbCrLf & vbCrLf, " ")
    uTidyup = Replace$(uTidyup, vbCrLf, " ")
    uTidyup = trim(uTidyup)
'    uTidyup = Replace$(uTidyup, vbCrLf & vbCrLf, " ")
'    uTidyup = vbCrLf & vbCrLf & uTidyup

End Function

Private Sub uGetAppointmentInfo(ByRef pMeet As MeetingItem, ByRef pAppt As AppointmentItem, ByVal ShowBody As Boolean, ByRef Title As String, ByRef Text As String) ', ByRef Actions As String)
Dim sz As String

    sz = pAppt.Subject
    If sz = "" Then _
        sz = pMeet.Subject

'    Form1.Add "uGetAppointmenuInfo(): subject=" & sz & " start=" & pAppt.Start & " end=" & pAppt.End

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

'Private Function uGetOutlook(ByRef pOutlook As Outlook.Application) As Boolean
'
'    On Error Resume Next
'
'End Function

Private Function uGetBodyText(ByRef Mail As MailItem) As String
Dim pDoc As HTMLDocument
Dim sz As String

    On Error Resume Next

    If Mail.BodyFormat <> olFormatHTML Then
        sz = Mail.Body

    Else
        Err.Clear
        Set pDoc = New HTMLDocument
        If (Err.Number <> 0) Or (ISNULL(pDoc)) Then
            sz = Mail.Body

        Else
            pDoc.Body.innerHTML = Mail.Body
            sz = pDoc.Body.innerText

        End If
    End If

    sz = g_SafeLeftStr(sz, 128, True)
    sz = Replace$(sz, vbCrLf & vbCrLf & Chr$(&HA0) & vbCrLf & vbCrLf, " ")
    sz = Replace$(sz, vbCrLf & vbCrLf, " ")

    uGetBodyText = sz

End Function

Private Sub pOLApp_Reminder(ByVal Item As Object)
Dim pa As AppointmentItem

    If TypeOf Item Is AppointmentItem Then
        Set pa = Item

    Else
'        List1.AddItem TypeName(Item)
        Exit Sub

    End If

Dim szSubject As String
Dim szClass As String

    szSubject = "<REDACTED>"

    Select Case pa.Sensitivity
    Case olPersonal
        szClass = CLASS_PERSONAL

    Case olPrivate
        szClass = CLASS_PRIVATE

    Case olConfidential
        szClass = CLASS_CONFIDENTIAL

    Case Else
        szClass = CLASS_NORMAL
        szSubject = pa.Subject
        If szSubject = "" Then _
            szSubject = "<subject>"

        If pa.Location <> "" Then _
            szSubject = szSubject & vbCrLf & pa.Location

    End Select

Dim hr As Long

    hr = snDoRequest("notify?app-sig=" & App.ProductName & _
                     "&id=" & szClass & _
                     "&uid=" & pa.EntryID & _
                     "&title=Reminder" & _
                     "&text=" & szSubject & vbCrLf & uTime(pa.Start) & " (" & uGetMinutes(pa.Start) & ")" & _
                     "&action=Open,@99" & _
                     "&priority=" & CStr(pa.Importance - 1) & _
                     "&password=" & mPassword)

Dim pm As TItem

    Set pm = New TItem
    pm.SetTo pa, hr, mPassword
    mMail.Add pm

End Sub

Private Function uTime(ByVal dTime As Date) As String

    If (Day(dTime) = Day(Now)) And (Month(dTime) = Month(Now)) And (Year(dTime) = Year(Now)) Then
        uTime = Format$(dTime, "Short Time")

    Else
        uTime = CStr(dTime)
        
    End If

End Function

Private Function uGetMinutes(ByVal dTime As Date) As String
Dim i As Long

    i = DateDiff("n", Now, dTime)
    If i = 0 Then
        uGetMinutes = "now"

    ElseIf i < 0 Then
        uGetMinutes = "overdue by " & CStr(Abs(i)) & " min" & IIf(i = -1, "", "s")

    Else
        uGetMinutes = "in " & CStr(i) & " min" & IIf(i = 1, "", "s")

    End If

End Function
