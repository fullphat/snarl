VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMissedNotificationsPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TMissedNotificationsPanel.cls
    '/
    '/  Description:    Provides the missed notifications panel
    '/
    '/  Created:        9-Sep-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mSelected As TNotification
Dim mPanel As BPrefsPanel

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
'Dim pmi As OMMenuItem
'Dim pc As BControl
'Dim sz As String
'Dim i As Long
'
'    Select Case Control.GetName()
'    Case "item_list"
'        Set mSelected = g_NotificationRoster.MissedList.TagAt(Val(Value))
'
'        ' /* toolbar */
'        prefskit_SafeEnable mPanel, "toolbar_options", (Value <> "0")
'
'        If mPanel.Find("toolbar_options", pc) Then
'            prefskit_SetItem pc, 1, "enabled", IIf(mSelected.Info.DefaultAck <> "", 1&, 0&)
'            prefskit_SetItem pc, 2, "enabled", IIf(mSelected.Info.Actions.CountItems, 1&, 0&)
'
'        End If
'
''Dim sz As String
'
'        sz = Replace$(mSelected.Info.Title, vbCrLf, "¶ ")
'        If sz <> "" Then _
'            sz = sz & vbCrLf
'
'        sz = sz & Replace$(mSelected.Info.Text, vbCrLf, "¶ ")
'
'        prefskit_SafeSetText mPanel, "the_detail", sz
'
'Dim pm As CTempMsg
'
'        Set pm = New CTempMsg
'        pm.Add "image-file", g_TranslateIconPath(mSelected.Info.IconPath, "")
'        If mPanel.Find("the_icon", pc) Then _
'            pc.DoExCmd B_SET_ITEM, pm
'
'
''        prefskit_SafeSetText mPanel, "the_title", Replace$(mSelected.Title, vbCrLf, "¶ ")
''        prefskit_SafeSetText mPanel, "the_text", Replace$(mSelected.Text, vbCrLf, "¶ ")
''        prefskit_SafeEnable mPanel, "fb>ack", (mSelected.DefaultAck <> "")
''
''Debug.Print "********************************* " & mSelected.Actions.CountItems
''
''        sz = Format$(mSelected.DateStamp, "hh:mm:ss") & _
''             IIf(DateDiff("d", Now, mSelected.DateStamp) = "0", " today", " on " & Format$(mSelected.DateStamp, "short date")) & vbCrLf & _
''             Replace$(mSelected.ClassObj.App.Name, vbCrLf, "¶ ") & " (" & mSelected.ClassObj.Description & ")"
''
''        If mSelected.RemoteHost <> "" Then _
''        If Not (mSelected.Socket Is Nothing) Then _
''            sz = sz & " on " & mSelected.RemoteHost
''
'''        If mSelected.DefaultAck <> "" Then _
''            sz = sz & "  Callback: " & mSelected.DefaultAck
''
''        prefskit_SafeSetText mPanel, "the_detail", sz
'''        prefskit_SafeSetText mPanel, "the_callback", Replace$(mSelected.DefaultAck, vbCrLf, "¶ ")
''
'''        prefskit_SafeSetText mPanel, "sending_app", "Sent from " & mSelected.Sender & " using class '" & mSelected.Class & "'"
'''        prefskit_SafeSetText mPanel, "sending_class", mSelected.Class
''
'    Case "toolbar_options"
'
'        Select Case Val(Value)
'        Case 1
'            If mSelected.Info.DefaultAck <> "" Then _
'                g_ProcessAck mSelected.Info.DefaultAck
'
'        Case 2
'            With New OMMenu
'                For i = 1 To mSelected.Info.Actions.CountItems
'                    .AddItem .CreateItem(mSelected.Info.Actions.TagAt(i).Value, mSelected.Info.Actions.TagAt(i).Name)
'                    If i < mSelected.Info.Actions.CountItems Then _
'                        .AddSeparator
'
'                Next i
'
'                Set pmi = .Track(Control.Page.Panel.hWnd)
'                If Not (pmi Is Nothing) Then
'                    If g_SafeLeftStr(pmi.Name, 1) = "@" Then
'                        ' /* TO-DO: dynamic callback */
'
'                    Else
'                        g_ProcessAck pmi.Name
'
'                    End If
'                End If
'
'            End With
'
'        Case 3
'            ' /* copy to clipboard */
'            Clipboard.Clear
'            Clipboard.SetText mSelected.Info.Title & vbCrLf & mSelected.Info.Text & vbCrLf & mSelected.Info.ClassObj.App.Name
'
'        Case 4
'            ' /* display */
'            If Not g_IsDND() Then _
'                g_NotificationRoster.Add mSelected.Info, Nothing
'
'        End Select
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
'Dim pc As BControl
'
'    Select Case Control.GetName()
'    Case "fb>clear"
'        If Not (g_NotificationRoster Is Nothing) Then
'            g_NotificationRoster.ClearMissedList
'            UpdateList
'
'            prefskit_SafeSetText mPanel, "the_title", ""
'            prefskit_SafeSetText mPanel, "the_text", ""
'            prefskit_SafeSetText mPanel, "sending_app", ""
'
'        End If
'
'
'    Case "fb>ack"
'        If mSelected.Info.DefaultAck <> "" Then
'            g_ProcessAck mSelected.Info.DefaultAck
'
'        ElseIf (IsWindow(mSelected.Info.hWndReply) <> 0) And (mSelected.Info.uReplyMsg <> 0) Then
'            PostMessage mSelected.Info.hWndReply, mSelected.Info.uReplyMsg, SNARL_CALLBACK_INVOKED, ByVal mSelected.Info.Token
'
'        Else
'            Debug.Print "TMissedNotificationsPanel.fb>ack(): no ack or reply window/message"
'
'        End If
'
'    End Select

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

    g_NotificationRoster.ResetMissedList
    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go()
Dim pPage As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "Missed Notifications"
            .SetWidth 500

            Set pPage = new_BPrefsPage("", , Me, 500)

            With pPage
                .SetMargin 0

                .Add new_BPrefsControl("label", "", "While you were away...")

                ' /* list */

                Set pm = New CTempMsg
                pm.Add "item-height", 36&
                Set pc = new_BPrefsControl("listbox", "item_list", , , , pm)
                pc.SizeTo 0, 220
                .Add pc

                ' /* toolbar */

                .Add new_BPrefsControl("fancytoolbar", "toolbar_options", "Invoke|Actions|Copy to Clipboard|Display", , , , False)

                ' /* icon and content */

                Set pm = New CTempMsg
                pm.Add "scale_to_fit", 1&
                Set pc = new_BPrefsControl("image", "the_icon", "", , , pm)
                pc.SizeTo 48, 48
                .Add pc

                .Add new_BPrefsControl("label", "the_detail", Space$(512))

'                .Add new_BPrefsControl("fancybutton2", "fb>ack", "Invoke Callback", , , , False)
        
'                .Add new_BPrefsControl("seperator", "")
                .Add new_BPrefsControl("fancybutton2", "fb>clear", "Clear List")
        
            End With

            .AddPage pPage


            .Go

            UpdateList

            g_SetWindowIconToAppResourceIcon .hWnd

        End With

    Else

        g_ShowWindow mPanel.hWnd, True, True

    End If

End Sub

Public Sub UpdateList()

    If (mPanel Is Nothing) Then
        g_Debug "TMissedNotificationsPanel.UpdateList(): panel not available", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    If (g_NotificationRoster Is Nothing) Then
        g_Debug "TMissedNotificationsPanel.UpdateList(): can't find notification roster", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim pc As BControl

    If Not mPanel.Find("item_list", pc) Then
        g_Debug "TMissedNotificationsPanel.UpdateList(): can't find item_list control", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim pMissed As BTagList

    Set pMissed = g_NotificationRoster.MissedList

    ' /* if no missed notifications, just clear the list and disable the buttons */

    If pMissed.CountItems = 0 Then
        prefskit_SafeEnable mPanel, "fb>clear", False
        pc.SetText ""
        Exit Sub

    End If

Dim pItem As TNotification
Dim iCurrent As Long
Dim szt As String
Dim sz As String

    ' /* store the current selected item */

    iCurrent = Val(pc.GetValue())

    With pMissed
        ' /* enable/disable the 'Clear List' button based on number of missed */
        prefskit_SafeEnable mPanel, "fb>clear", True

        ' /* build the content string */

        .Rewind
        Do While .GetNextTag(pItem) = B_OK
            szt = g_FormattedMidStr(Replace$(pItem.Info.Title, "|", ":") & " (" & g_When(pItem.Info.DateStamp) & ")", 70)

            sz = sz & Replace$( _
                                szt & "#?0#?" & _
                                g_FormattedMidStr(Replace$(pItem.Info.Text, "|", ":"), 80) & "|", vbCrLf, "¶")

'                sz = sz & g_FormattedMidStr(Replace$(pInfo.Title, "|", ":"), 54) & "#?0#?" & "From " & pInfo.Sender & " (" & pInfo.Class & ") at " & Format$(pInfo.DateStamp, "hh:mm:ss") & "|"

        Loop

Dim i As Long

        pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

        ' /* set the icons */

        With pMissed
            For i = 1 To .CountItems
                Set pItem = .TagAt(i)
                pc.DoExCmd B_EXTENDED_COMMANDS.B_SET_ITEM, prefskit_CreateImageMessage(i, load_image_obj(g_TranslateIconPath(pItem.Info.IconPath, "")), pItem.WasSeen)

            Next i

        End With

    End With

    pc.SetValue CStr(iCurrent + 1)

End Sub



