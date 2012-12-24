VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Index As Long                ' // set by AlertAt() or FindAlert()

    ' /* Class defaults - set by snSetAlertClassDefault()  (V38) */
Public DefaultIcon As String

Public Token As Long                '// V41
Dim mLastToken As Long              '// V41: returned in certain cases (e.g. during duplicate suppression)

    ' /* V39 - used to suppress duplicates */
Dim mPrevTitle As String
Dim mPrevText As String
Dim mPrevTick As Long

Public Enum SN_CLASS_DUPLICATES
    SN_CD_ALLOW
    SN_CD_BLOCK
    SN_CD_DELAY

End Enum

Public Enum SN_CLASS_FLAGS
    SN_CF_DISABLED = 1

End Enum

Private Type T_ALERT
    Class As String                         '// unique identifier - once assigned, can't be changed
    Desc As String                          '// displayed name - added in V38, doc'd in V39
    App As TApp                             '// app who this class belongs to
    Config As ConfigSection                 '// V41: persistent settings
    Defaults As BPackedData
    Redirectors As BTagList                 '// each style/scheme to use
    NetForwards As BTagList                 '// guids

End Type

Dim mContent As T_ALERT

    ' /* R2.31 */
Public LastNotificationGeneratedTick As Long

    ' /* R2.6 */
Dim mData As BPackedData

Implements mObject

Private Sub Class_Initialize()

    With mContent
        .Class = SNARL_CLASS_GENERAL
        .Desc = ""
        Set .App = New TApp
'        .App.SetAsSnarl

    End With

    ' /* 44.51: create a blank pack */
    Set mData = New BPackedData
    mData.Add "duration", "-1"

End Sub

Private Property Get MObject_Type() As String
End Property

Friend Sub bInit(ByVal Id As String, ByVal Name As String, ByRef Owner As TApp, ByRef Settings As ConfigSection, ByVal Flags As SN_CLASS_FLAGS)
Dim bEnabled As Boolean

    ' /* initialize */

    g_Debug "TAlert.bInit()", LEMON_LEVEL_PROC_ENTER

    With mContent
        .Class = Id
        .Desc = Name
        Set .App = Owner
        Set .Config = Settings

        ' /* defaults */
        ' /* added in R2.1 (moved here in V41): if this global setting is enabled, new classes are auto-disabled when created */

        g_Debug "setting defaults..."
        bEnabled = True

        If (g_ConfigGet("ignore_new_classes") = "1") Or ((Flags And SN_CF_DISABLED)) Then _
            bEnabled = False

'                        "use_custom_icon::0#?custom_icon::#?" & _
'                        "use_custom_sound::0#?custom_sound::#?" & _
'                        "use_custom_ations::0#?custom_actions::#?" & _

        ' /* set defaults - if a default value is blank, we don't need to include it here */

        Set .Defaults = New BPackedData
        With .Defaults
            .Add "enabled", IIf(bEnabled, "1", "0")
            .Add "show_on_screen", "1"
            .Add "position", CStr(SN_SP_DEFAULT_POS)
            .Add "duration", CStr(SN_ND_APP_DECIDES)
            .Add "custom_timeout", "10"
            .Add "custom_priority", "0"
            .Add "priority", "0"
            .Add "use_style", "0"
            .Add "duplicates", CStr(SN_CD_ALLOW)
            .Add "log_during_dnd", "1"
            .Add "block_spam", "0"
            .Add "normal_mode", "1"
            .Add "away_mode", "1"
            .Add "busy_mode", "1"
            .Add "override-normal", "1"
            .Add "override-away", "1"
            .Add "override-busy", "1"
            .Add "block_if_foreground", "0"
            .Add "log_in_history", "3"
            .Add "colour-tint", CStr(rgba(0, 0, 0))
            .Add "send_to_subscribers", "0"
            .Add "redact-when", "0"
            .Add "send_to_redirects", "1"
            .Add "redact-always", "0"

        End With

Dim sn As String
Dim sv As String

        ' /* read redirection info */
        Set .Redirectors = new_BTagList()
        If Not (Settings Is Nothing) Then
            With New BPackedData
                .SetTo Settings.GetValueWithDefault("forward", "")

                .Rewind
                Do While .GetNextItem(sn, sv)
                    mContent.Redirectors.Add new_BTagItem(sn, sv)

                Loop
            End With
        End If

        ' /* read forwarding info */
        Set .NetForwards = new_BTagList()
        If Not (Settings Is Nothing) Then
            With New BPackedData
                .SetTo Settings.GetValueWithDefault("netforward", "")
                .Rewind
                Do While .GetNextItem(sn, sv)
                    mContent.NetForwards.Add new_BTagItem(sn, sv)

                Loop
            End With
        End If
    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function IsEnabled() As Boolean

    If Not (mContent.Config Is Nothing) Then _
        IsEnabled = (Me.GetValue("enabled") = "1")

End Function

Public Sub SetEnabled(ByVal Enabled As Boolean)

    If Not (mContent.Config Is Nothing) Then _
        Me.SetValue "enabled", IIf(Enabled, "1", "0")

End Sub

Public Function Name() As String

    Name = mContent.Class

End Function

Public Function CustomStyle() As String

    If Not (mContent.Config Is Nothing) Then _
        CustomStyle = Me.GetValue("style")

End Function

Public Function StyleName() As String

    StyleName = style_GetStyleName(CustomStyle)

End Function

Public Function SchemeName() As String

    SchemeName = style_GetSchemeName(CustomStyle)

End Function

Public Sub SetStyleAndScheme(ByVal Style As String, ByVal Scheme As String)

    If (mContent.Config Is Nothing) Then _
        Exit Sub

Dim sz As String

    sz = LCase$(Style & "/" & Scheme)               ' // convert to style/scheme pair

    If sz = LCase$(Me.CustomStyle) Then _
        Exit Sub                                    ' // ignore if no change

'    If ((Me.GetValue("style") = "") And (sz = g_ConfigGet("default_style"))) Or (sz = LCase$(Me.GetValue("style"))) Then _
        Exit Sub

    Me.SetValue "style", sz
    Debug.Print "*** scheme changed to '" & sz & "' for class " & mContent.Class & " ***"

End Sub

'Public Function Sound() As E_CLASS_SOUND
'
'    Sound = E_CLASS_SOUND_APP_DECIDES
'
'    If (mContent.Config Is Nothing) Then _
'        Exit Function
'
'Dim sz As String
'
'    If mContent.Config.Find("sound", sz) Then _
'        Sound = Val(sz)
'
'End Function
'
'Public Function CustomSound() As String
'
'    If (mContent.Config Is Nothing) Then _
'        Exit Function
'
'    CustomSound = mContent.Config.GetValueWithDefault("custom_sound", "")
'
'End Function

Public Sub Reset()

    If (mContent.Config Is Nothing) Or (mContent.Defaults Is Nothing) Then _
        Exit Sub

Dim s1 As String
Dim s2 As String

    With mContent
        .Defaults.Rewind
        Do While .Defaults.GetNextItem(s1, s2)
            .Config.Update s1, s2

        Loop

        Set .Redirectors = new_BTagList
        Set .NetForwards = new_BTagList

    End With

End Sub

Public Function Description() As String

    If mContent.Class = "_all" Then
        Description = IIf(mContent.App.CountAlerts > 1, "Other Notifications", "All Notifications")

    Else
        Description = IIf(mContent.Desc = "", mContent.Class, mContent.Desc)

    End If

End Function

Public Function IsClickThru() As Boolean

    If g_ConfigGet("ignore_input") = "1" Then
        IsClickThru = True

    ElseIf Not (mContent.Config Is Nothing) Then
        IsClickThru = (Me.GetValue("ignore_input") = "1")

    End If

End Function

Public Function GetValue(ByVal Name As String) As String

    ' /* pre-set to default */
    If Not (mContent.Defaults Is Nothing) Then _
        GetValue = mContent.Defaults.ValueOf(Name)

Dim sz As String

    If NOTNULL(mContent.Config) Then
        If mContent.Config.Find(Name, sz) Then _
            GetValue = sz

    End If

End Function

Public Sub SetValue(ByVal Name As String, ByVal Value As String)

    If Not (mContent.Config Is Nothing) Then
        mContent.Config.Update Name, Value
        mContent.App.WriteConfig

    End If

End Sub

Public Function App() As TApp

    If (mContent.App Is Nothing) Then _
        MsgBox "ALERT!"

    Set App = mContent.App

End Function

'Public Sub TestNotification()
'Dim pInfo As T_NOTIFICATION_INFO
'
'    With pInfo
'        .Title = Me.Description
'        .Text = "Test Notification"
'        .Timeout = -1
'        .APIVersion = VB.App.Major
'        .IntFlags = VB.App.Major
'        Set .ClassObj = Me
'.ClassObj.App.Show41
'
'    End With
'
'    Me.ShowNotification pInfo, Nothing
'
'End Sub

Friend Function ShowNotification(ByRef Info As T_NOTIFICATION_INFO, ByRef Args As BPackedData) As Long

    ' /* IMPORTANT! This should *only* be called by TApp.Show(), TApp.Show41() and TAlert.TestNotification() */

    g_SetLastError SNARL_ERROR_SYSTEM
    If (g_NotificationRoster Is Nothing) Then _
        Exit Function

    ' /* do nothing if class is disabled */

    If Me.GetValue("enabled") = "0" Then
        g_Debug "TAlert.ShowNotification(): class '" & mContent.Class & "' has been disabled by the user"
        g_SetLastError SNARL_ERROR_CLASS_BLOCKED
        Exit Function

    End If

    ' /* check for duplicate content */

Dim lDup As SN_CLASS_DUPLICATES

    If (Info.Title = mPrevTitle) And (Info.Text = mPrevText) Then
        lDup = Me.GetValue("duplicates")
        If lDup = SN_CD_BLOCK Then
            g_Debug "TAlert.ShowNotification(): blocked duplicate", LEMON_LEVEL_INFO
            ShowNotification = mLastToken
            Exit Function

        ElseIf (lDup = SN_CD_DELAY) And ((GetTickCount() - mPrevTick) <= Val(g_ConfigGet("suppress_delay"))) Then
            g_Debug "TAlert.ShowNotification(): blocked duplicate within delay timeout", LEMON_LEVEL_INFO
            ShowNotification = mLastToken
            Exit Function

        End If

    Else
        mPrevTitle = Info.Title
        mPrevText = Info.Text

    End If

    mPrevTick = GetTickCount()

Dim lDur As SN_NOTIFICATION_DURATION

    ' /* figure out timeout */
    lDur = Val(Me.GetValue("duration"))
    Select Case lDur
    Case SN_ND_APP_DECIDES
        ' /* V39: app can modify duration via class default */
'        Debug.Print "TIMEOUT: APP_DECIDES: " & Info.Timeout & " -> " & Me.DefaultTimeout
        If Info.Timeout = -1 Then _
            Info.Timeout = g_SafeLong(mData.ValueOf("duration"))

'        If Me.DefaultTimeout > 0 Then _
            Info.Timeout = Me.DefaultTimeout

    Case SN_ND_CUSTOM
'        Debug.Print "TIMEOUT: CUSTOM"
        Info.Timeout = Val(Me.GetValue("custom_timeout"))

'    Case SN_ND_DEFAULT
'        Info.Timeout = Val(Me.GetValue("default_duration"))

    Case Else
        g_Debug "TAlert.ShowNotification(): invalid duration setting '" & CStr(lDur) & "'", LEMON_LEVEL_WARNING
        Info.Timeout = Val(Me.GetValue("default_duration"))

    End Select

    ' /* V38 - if any item isn't provided we use the class default */

    If Info.Title = "" Then _
        Info.Title = mData.ValueOf("title")

    If Info.Text = "" Then _
        Info.Text = mData.ValueOf("text")

    ' /* V40.25: custom icon */

    If Me.GetValue("custom_icon") <> "" Then            '// Me.GetValue("use_custom_icon") = "1" Then
        Info.IconPath = Me.GetValue("custom_icon")

    ElseIf Info.IconPath = "" Then
        Info.IconPath = Me.DefaultIcon

    End If

'    If Not (Args Is Nothing) Then
'        If (Info.IconPath = "") And (Args.Exists("icon-base64")) Then
'            ' /* R2.4 Beta 4: use Base64 encoded icon data, if there is any */
'            Info.IconPath = g_GetBase64Icon(Args.ValueOf("icon-base64"))
'
'        ElseIf (Info.IconPath = "") And (Args.Exists("icon-phat64")) Then
'            ' /* R2.4.2 DR3: phat64 encoding? */
'            Info.IconPath = g_GetPhat64Icon(Args.ValueOf("icon-phat64"))
'
'        End If
'
'    End If

    ' /* R2.4 DR7: if still no icon, use app's icon */
    If Info.IconPath = "" Then _
        Info.IconPath = Me.App.Icon

    ' /* R2.31d3: custom ack */
    If Me.GetValue("custom_ack") <> "" Then
        Info.DefaultAck = Me.GetValue("custom_ack")

    ElseIf Info.DefaultAck = "" Then
        ' /* R2.4 DR7: name change */
        If mData.Exists("callback") Then
            Info.DefaultAck = mData.ValueOf("callback")

        ElseIf mData.Exists("ack") Then
            Info.DefaultAck = mData.ValueOf("ack")

        End If

    End If

    ' /* V40.25: custom priority */
    If Me.GetValue("custom_priority") = "1" Then _
        Info.Priority = Val(Me.GetValue("priority"))

    ' /* sound */

    ' /* the notification roster takes care of the global sound settings; we
    '    just need to provide our preferred sound which, as of R2.4, is as follows:
    '       1. user-defined sound
    '       2. notification-defined sound
    '       3. class-defined sound
    '
    ' */

    If Me.GetValue("custom_sound") <> "" Then
'    If Me.GetValue("use_custom_sound") = "1" Then
        Info.SndFile = Me.GetValue("custom_sound")

    ElseIf Info.SndFile = "" Then
        ' /* R2.5.1: fixed logic */
        Info.SndFile = mData.ValueOf("sound")

    End If


    Info.Position = Val(Me.GetValue("position"))

'    MsgBox Me.GetValue("use_style"), , Me.Name & " > " & Me.Description

    If Me.GetValue("use_style") = "1" Then
        Info.StyleName = style_GetStyleName(Me.GetValue("style"))
        Info.SchemeName = style_GetSchemeName(Me.GetValue("style"))

    End If

'        MsgBox Info.StyleName & " + " & Info.SchemeName

    ' /* R2.4 DR7: now store originating class object */
    Set Info.ClassObj = Me

    ' /* R2.4.2: if text begins with http(s):// and no default callback specified
    '    and auto-detect URLs is enabled, set the callback now */

    If (Info.DefaultAck = "") And (g_ConfigGet("auto_detect_url") = "1") Then
        If (g_SafeLeftStr(LCase$(Info.Text), 7) = "http://") Or (g_SafeLeftStr(LCase$(Info.Text), 8) = "https://") Then
            g_Debug "TAlert.ShowNotification(): auto-detected URL '" & Info.Text & "'"
            Info.DefaultAck = Info.Text

        End If

    End If

Dim sn As String
Dim sv As String

    If (Info.Actions Is Nothing) Then _
        Set Info.Actions = new_BTagList()                   ' // create list even if there are no actions defined

    ' /* R2.4.2 DR3: add any custom actions */

    If Me.GetValue("custom_actions") <> "" Then
'    If Me.GetValue("use_custom_actions") = "1" Then
        With New TPackedData
            If .SetTo(Me.GetValue("custom_actions"), ";", "=") Then
                .Rewind
                Do While .GetNextItem(sn, sv)
                    Info.Actions.Add new_BTagItem(sn, sv)

                Loop
            End If
        End With
    End If

    ' /* R2.4 DR7 - append any existing actions */

    If NOTNULL(Args) Then
        With Args
            .Rewind
            Do While .GetNextItem(sn, sv)
                If sn = "action" Then _
                    uGetAction Info.Actions, sv
    
            Loop
        End With
    End If

    ' /* R2.5 Beta 2: check "log" value */

Dim bLog As Boolean

    bLog = True
    
'    If NOTNULL(Args) Then
'        If Args.ValueOf("log") = "0" Then _
'            bLog = False
'
'    End If

    ' /* R2.5.1: if, after processing, the title, text and icon are all (still) NULL, fail */

    If (Info.Title = "") And (Info.Text = "") And (Info.IconPath = "") Then
        g_Debug "TAlert.ShowNotification(): must supply at least a title, some text or an icon", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    ' /* background tint */

Dim dw As Long

    dw = g_SafeLong(Me.GetValue("colour-tint"))
    If dw <> rgba(0, 0, 0) Then
        dw = rgba(get_red(dw), get_green(dw), get_blue(dw))
        If Info.OriginalContent <> "" Then _
            Info.OriginalContent = Info.OriginalContent & "#?"

        Info.OriginalContent = Info.OriginalContent & "colour-tint::" & CStr(dw)

    End If

    ' /* V44.56 - last check: must have one of title, text or icon */

    If (Info.Title = "") And (Info.Text = "") And (Info.IconPath = "") Then
        g_Debug "TAlert.ShowNotification(): not displaying: must have at least one of title, text or icon", LEMON_LEVEL_CRITICAL
        g_SetLastError SNARL_ERROR_ARG_MISSING
        ShowNotification = 0

    Else
        ' /* send it to the notification roster */
        ShowNotification = g_NotificationRoster.Add(Info, Args, bLog)
        mLastToken = Info.Token
'    Debug.Print "TAlert.ShowNotification(): token=" & CStr(mLastToken) & " sender=" & CStr(mContent.App.Token) & " (" & mContent.App.Signature & ") title='" & Info.Title & "' text='" & Info.Text & "'"

    End If

End Function

Public Sub AddRedirector(ByVal StyleAndScheme As String, Optional ByVal Flags As SN_REDIRECTION_FLAGS = SN_RF_ALWAYS)

    If (mContent.Redirectors Is Nothing) Then _
        Exit Sub

    With mContent
        If .Redirectors.IndexOf(StyleAndScheme, False) <> 0 Then
            .Redirectors.Update StyleAndScheme, CStr(Flags)

        Else
            .Redirectors.Add new_BTagItem(StyleAndScheme, CStr(Flags))

        End If
        .Config.Update "forward", taglist_as_string(.Redirectors)
        .App.WriteConfig                                                ' // write the config

    End With

End Sub

Public Sub RemRedirector(ByVal StyleAndScheme As String)

    If (mContent.Redirectors Is Nothing) Then _
        Exit Sub

    With mContent
        .Redirectors.Remove .Redirectors.IndexOf(StyleAndScheme)
        .Config.Update "forward", taglist_as_string(.Redirectors)
        .App.WriteConfig

    End With

End Sub

Public Function RedirectList() As BTagList

    Set RedirectList = mContent.Redirectors

End Function

Public Sub AddNetForward(ByVal Guid As String)

    If (mContent.NetForwards Is Nothing) Then _
        Exit Sub

    With mContent
        .NetForwards.Add new_BTagItem(Guid, "")
        .Config.Update "netforward", taglist_as_string(.NetForwards)
        .App.WriteConfig                                                ' // write the config

    End With

End Sub

Public Sub RemNetForward(ByVal Guid As String)

    If (mContent.NetForwards Is Nothing) Then _
        Exit Sub

    With mContent
        .NetForwards.Remove .NetForwards.IndexOf(Guid)
        .Config.Update "netforward", taglist_as_string(.NetForwards)
        .App.WriteConfig

    End With

End Sub

'Public Function HasNetForward(ByVal Guid As String) As Boolean
'
'    If Not (mContent.Redirectors Is Nothing) Then _
'        HasRedirector = (mContent.Redirectors.IndexOf(StyleAndScheme) <> 0)
'
'End Function

Public Function NetForwardList() As BTagList

    Set NetForwardList = mContent.NetForwards

End Function







Private Sub uGetAction(ByRef ListToAddTo As BTagList, ByVal PackedAction As String)
Dim i As Long

    ' /* only interested in _first_ comma */

    i = InStr(PackedAction, ",")
    If i = 0 Then _
        Exit Sub

Dim szLbl As String
Dim szCmd As String

    szLbl = g_SafeLeftStr(PackedAction, i - 1)
    szCmd = g_SafeRightStr(PackedAction, Len(PackedAction) - i)

    ' /* must have a label and command */

    If (szLbl = "") Or (szCmd = "") Then _
        Exit Sub

    ListToAddTo.Add new_BTagItem(szLbl, szCmd)

End Sub

Friend Sub DoRedirection(ByRef Info As T_NOTIFICATION_INFO)

    If ISNULL(mContent.Redirectors) Then _
        Exit Sub

'Dim pCopyInfo As T_NOTIFICATION_INFO
Dim fConditions As SN_REDIRECTION_FLAGS
Dim pt As BTagItem

    With mContent.Redirectors
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            fConditions = Val(pt.Value)

            Debug.Print "TAlert.DoRedirection(): busy:" & g_IsDND() & " away:" & g_IsAway() & " cond:" & g_HexStr(fConditions)

            If (g_IsDND()) And ((fConditions And SN_RF_WHEN_BUSY) <> 0) Then
                ' /* redirect when busy */
                fConditions = -1

            ElseIf (g_IsAway()) And ((fConditions And SN_RF_WHEN_AWAY) <> 0) Then
                ' /* redirect when away */
                fConditions = -1

            ElseIf (fConditions And SN_RF_WHEN_ACTIVE) <> 0 Then
                ' /* redirect when active */
                fConditions = -1

            Else
                ' /* no shadowing */
                fConditions = 0

            End If

            If fConditions <> 0 Then
                g_Debug "TAlert.DoRedirection(): redirecting to '" & pt.Name & "'..."
                g_StyleRoster.RedirectTo pt.Name, Info
'                LSet pCopyInfo = Info
'                pCopyInfo.StyleName = LCase$(style_GetStyleName(pt.Name))
'                pCopyInfo.SchemeName = LCase$(style_GetSchemeName(pt.Name))
'                g_NotificationRoster.Redirect pCopyInfo

            End If
        Loop

    End With

End Sub

Friend Sub DoForwarding(ByRef Info As T_NOTIFICATION_INFO)

    If ISNULL(mContent.NetForwards) Then _
        Exit Sub

Dim psForwarder As ConfigSection
Dim pt As BTagItem
Dim i As Long

    With mContent.NetForwards
        .Rewind
        Do While .GetNextTag(pt) = B_OK

            i = g_SubsRoster.Config.FindSection(pt.Name)
            If i Then
                g_SubsRoster.ForwardNotification g_SubsRoster.Config.SectionAt(i), Info

            Else
                g_Debug "TAlert.DoForwarding(): can't find detail for forward '" & pt.Name & "/" & pt.Value & "'", LEMON_LEVEL_CRITICAL

            End If

        Loop

    End With

End Sub

Public Function RealIconPath() As String

    If Me.DefaultIcon <> "" Then
        RealIconPath = Me.DefaultIcon

    Else
        RealIconPath = mContent.App.RealIconPath()

    End If

End Function

Public Sub Augment(ByRef Data As BPackedData)

    If NOTNULL(Data) Then
        Set mData = Data
        If Not mData.Exists("duration") Then _
            mData.Add "duration", "-1"

    End If

''        Debug.Print .Description & " >> " & Data.Exists("icon")
'
'        .DefaultIcon = Data.ValueOf("icon")
'
'        If (.DefaultIcon = "") And (Data.Exists("icon-base64")) Then
'            ' /* R2.4 Beta 4: check for MIME encoded icon */
'            .DefaultIcon = g_GetBase64Icon(Data.ValueOf("icon-base64"))
'
'        ElseIf (.DefaultIcon = "") And (Data.Exists("icon-phat64")) Then
'            ' /* R2.4.2 DR3: phat64 encoding? */
'            .DefaultIcon = g_GetPhat64Icon(Data.ValueOf("icon-phat64"))
'
'        End If

End Sub

Public Function AppProvidedSettings() As BPackedData

    Set AppProvidedSettings = mData

End Function
