VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TExtPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TExtPage.cls
    '/
    '/  Description:    The loaded extensions preferences page
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mPage As BPrefsPage
Dim mCurrent As TExtension

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg
Dim i As Long

    On Error Resume Next

    With mPage
        .Add new_BPrefsControl("label", "", "Extensions are modules loaded by Snarl during startup; " & _
                                            "they can register with Snarl and create notifications in exactly the " & _
                                            "same way applications can.")

        Set pm = New CTempMsg
        pm.Add "item-height", 36&
        Set pc = new_BPrefsControl("listbox", "lb>extensions", , , "1", pm)
        pc.SizeTo 0, 200
        .Add pc

        .Add new_BPrefsControl("fancytoolbar", "extension_toolbar", "Start|Stop|Details|Configure...||Show in Explorer||Uninstall")
        .Add new_BPrefsSeparator
        .Add new_BPrefsControl("fancytoolbar", "toolbar-general", "Reload Extensions|")
'        .Add new_BPrefsControl("fancytoolbar", "toolbar-general", "Reload Extensions|Install Extension...")
'        .Add new_BPrefsSeparator

        .Add new_BPrefsControl("label", "", "To prevent extensions from being loaded, hold down the 'E' key while Snarl loads.")

'        Set pm = New CTempMsg
'        pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\info_emblem.png"
'        Set pc = new_BPrefsControl("labelex", "", "To prevent extensions from being loaded, hold down the 'E' key while Snarl loads.", , , pm)
'        pc.SizeTo 0, 24
'        .Add pc

    End With

    uUpdateExtensions

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pmi As OMMenuItem
Dim pm As MMessage
Dim pc As BControl
Dim b As Boolean

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()
    Case "lb>extensions"
        ' /* extension selected */

        Set mCurrent = Nothing

        If (g_ExtnRoster Is Nothing) Then _
            Exit Sub

        Set mCurrent = melonRosterItemAt(g_ExtnRoster, Val(Value))
        If (mCurrent Is Nothing) Then _
            Exit Sub

'        If mPage.Find("ext_desc", pc) Then _
            pc.SetText px.Description

'        If mPage.Find("ext_info", pc) Then _
            pc.SetText "Version: " & uStr(px.InfoString())

'        If mPage.Find("ext_info", pc) Then _
            pc.SetText px.Copyright

        uUpdateStatus

'        If mPage.Find("fb2>ext.cfg", pc) Then
'            b = False
'            If px.IsEnabled Then _
'                b = (px.Flags And SNARL_EXT_IS_CONFIGURABLE)
'
'            pc.SetEnabled b
'
'        End If

    Case "extension_toolbar"
        '//"Start|Stop|Details|Configure...||Show in Explorer"

        If (mCurrent Is Nothing) Then
            g_Debug "TExtPage.ControlChanged(): extension_toolbar: no selected extension", LEMON_LEVEL_CRITICAL
            Exit Sub

        End If

        Select Case Val(Value)
        Case 1
            mCurrent.SetEnabled True
            uUpdateExtensions
            g_ExtnRoster.WriteExcludeList

        Case 2
            mCurrent.SetEnabled False
            uUpdateExtensions
            g_ExtnRoster.WriteExcludeList

        Case 3
            mCurrent.ShowDetails
            
        Case 4
            mCurrent.DoPrefs mPage.Panel.hWnd

        Case 6
            ShellExecute mPage.Panel.hWnd, "open", g_MakePath(mCurrent.Path), vbNullString, vbNullString, SW_SHOW
            
'            ' /* advanced */
'            With New OMMenu
'                .AddItem .CreateItem("expl", "Show in Explorer...")
'                .AddItem .CreateItem("inst", "Install Extension...")
'
'                Set pmi = .Track(mPage.Panel.hWnd)
'                If Not (pmi Is Nothing) Then
'                    Select Case pmi.Name
'                    Case "expl"
'
'                    Case "inst"
'                        uDoInstallExtension
'
'                    End Select
'                End If
'
'            End With

        Case 8
            ' /* uininstall */
            uUninstall

        End Select


    Case "toolbar-general"

        Select Case Val(Value)
        Case 1
            ' /* reload */
            If Not (g_ExtnRoster Is Nothing) Then
                g_ExtnRoster.Restart
                uUpdateExtensions

            End If

        Case 2
            ' /* install */
'            uDoInstallExtension

        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()
    Case "lb>extensions"
        ' /* extension double-clicked */

        If Not (mCurrent Is Nothing) Then _
            mCurrent.DoPrefs mPage.Panel.hWnd

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
Dim bEnable As Boolean
Dim pc As BControl
Dim pm As CTempMsg

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

    If Control.GetName() = "lb>extensions" Then
        Select Case Notification
        Case "update_list"
            uUpdateExtensions

        Case "checked"
            If (mCurrent Is Nothing) Then _
                Exit Sub

            ' /* construct the extended command message */
            Set pm = New CTempMsg
            pm.Add "what", "checked"
            bEnable = (Control.DoExCmd(B_GET_ITEM, pm) <> "0")  ' // get enable/disable from checkbox setting

            mCurrent.SetEnabled bEnable
            g_ExtnRoster.WriteExcludeList
            uUpdateStatus

            pm.Add "image-object", load_image_obj(IIf(mCurrent.State = SN_ES_FAILED, g_MakePath(App.Path) & "etc\icons\bad.png", mCurrent.Icon))
            pm.Add "subtext", IIf(mCurrent.State = SN_ES_FAILED, "Failed to start: " & mCurrent.LastError, mCurrent.Description)
            pm.Add "greyscale", IIf(mCurrent.State = SN_ES_NOT_LOADED, -1&, 0&)
            Control.DoExCmd B_SET_ITEM, pm

            If Control.Page.Panel.Find("cb>apps", pc) Then _
                pc.Notify "update_list", Nothing

        End Select
    End If

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 0

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()

End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)

End Sub

Private Function KPrefsPage_hWnd() As Long

End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)

End Sub

Private Sub uUpdateExtensions()
Dim pc As BControl

    If (g_ExtnRoster Is Nothing) Or (Not (mPage.Find("lb>extensions", pc))) Then _
        Exit Sub

    pc.DoExCmd "lock"

Dim list_content As String
Dim px As TExtension
Dim pm As CTempMsg
Dim pr As MRoster
Dim z As String
Dim i As Long

    z = pc.GetValue()

    list_content = ""
    Set pr = g_ExtnRoster
    pr.Rewind
    Do While pr.NextItem(px)
        list_content = list_content & px.Name & "#?" & CStr(0) & "#?" & px.Description & "|"

    Loop

    ' /* set the content */

    list_content = g_SafeLeftStr(list_content, Len(list_content) - 1)
    pc.SetText list_content

    prefskit_SafeEnable mPage.Panel, "extension_toolbar", (list_content <> "")

    If list_content <> "" Then

        ' /* set the checkboxes and icons */

        i = 0
        pr.Rewind
        Do While pr.NextItem(px)
            i = i + 1

            Set pm = New CTempMsg
            pm.Add "index", i
    '        pm.Add "checked", IIf(px.State = SN_ES_LOADED, 1&, 0&)
            pm.Add "greyscale", IIf(px.State = SN_ES_NOT_LOADED, 1&, 0&)
    
            If px.State = SN_ES_FAILED Then
                pm.Add "image-object", load_image_obj(g_MakePath(App.Path) & "etc\icons\bad.png")
    
    '        ElseIf InStr(px.Icon, ",") Then
    '            pm.Add "image-icon", px.Icon
    
            Else
                pm.Add "image-object", load_image_obj(g_TranslateIconPath(px.Icon, ""))

            End If
    
            pc.DoExCmd B_SET_ITEM, pm

        Loop

        pc.SetValue z

    End If

    pc.DoExCmd "unlock"

End Sub

Private Sub uUpdateStatus()

    If (mPage Is Nothing) Or (mCurrent Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    ' /* toolbar */

    If mPage.Find("extension_toolbar", pc) Then
        '//"Start|Stop|Details|Configure...||Show in Explorer"
        prefskit_SetItem pc, 1, "enabled", IIf(mCurrent.State = SN_ES_LOADED, 0&, 1&)
        prefskit_SetItem pc, 2, "enabled", IIf(mCurrent.State <> SN_ES_NOT_LOADED, 1&, 0&)
        prefskit_SetItem pc, 3, "enabled", IIf(mCurrent.State <> SN_ES_FAILED, 1&, 0&)
        prefskit_SetItem pc, 4, "enabled", IIf(mCurrent.IsConfigurable, 1&, 0&)
        prefskit_SetItem pc, 6, "enabled", IIf(mCurrent.State <> SN_ES_FAILED, 1&, 0&)

    End If

End Sub

Private Function uStr(ByVal str As String) As String

    If str = "" Then
        uStr = "<n/a>"

    Else
        uStr = str

    End If

End Function

'Private Sub uDoInstallExtension()
'Dim szError As String
'Dim szPath As String
'Dim sz As String
'
'    sz = InputBox("Enter the name of the extension to install", "Install Extension")
'    If sz = "" Then _
'        Exit Sub
'
''    If LCase$(g_GetExtension(sz)) <> "extension" Then _
'        sz = sz & ".extension"
'
''    If Not (g_ExtnRoster Is Nothing) Then
''        If g_ExtnRoster.LoadExtensionByClass(sz, szError) Then
''            If g_GetUserFolderPath(szPath) Then
''                uCreateMarkerFile g_MakePath(szPath) & "extensions\" & sz & ".extension"
''                If Not (g_ExtnRoster Is Nothing) Then _
''                    g_ExtnRoster.WriteExcludeList
''
''            End If
''
''            uUpdateExtensions
''
''        Else
''            MsgBox szError, vbOKOnly Or vbExclamation, "Error Installing Extension"
''
''        End If
''
''    End If
'
'End Sub

Private Sub uCreateMarkerFile(ByVal File As String)
Dim n As Integer

    On Error Resume Next
    
    n = FreeFile()
    Open File For Output As #n
    Close #n

End Sub

Private Sub uUninstall()

    If MsgBox("Are you sure you want to uninstall " & mCurrent.Name & "?", vbYesNo Or vbQuestion, App.Title) = vbNo Then _
        Exit Sub

Dim bFailed As Boolean
Dim sz As String

    sz = mCurrent.Class & ".extension"

    mCurrent.SetEnabled False

    ' /* delete from program files first */
    If g_Exists(g_MakePath(App.Path) & "extensions\" & sz) Then
        If DeleteFile(g_MakePath(App.Path) & "extensions\" & sz) = 0 Then _
            bFailed = True

    End If

    ' /* %COMMONAPPDATA% */
    If g_Exists(g_GetUserFolderPathStr(True) & "extensions\" & sz) Then
        If DeleteFile(g_GetUserFolderPathStr(True) & "extensions\" & sz) = 0 Then _
            bFailed = True

    End If

    ' /* %APPDATA% */
    If g_Exists(g_GetUserFolderPathStr(False) & "extensions\" & sz) Then
        If DeleteFile(g_GetUserFolderPathStr(False) & "extensions\" & sz) = 0 Then _
            bFailed = True

    End If

    If Not bFailed Then
        g_ExtnRoster.Remove mCurrent.Class
        uUpdateExtensions

    Else
        MsgBox "Unable to uninstall " & mCurrent.Name & ".  You might not have sufficient permissions.", vbExclamation Or vbOKOnly, App.Title

    End If

End Sub
