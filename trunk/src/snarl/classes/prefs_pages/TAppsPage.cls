VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAppsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TAppsPage.cls
    '/
    '/  Description:    The registered applications preferences page
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mCurrentApp As TApp
Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

    With mPage
        .SetMargin 0
        .Add new_BPrefsControl("label", "", "The following applications are registered with Snarl:")

        Set pm = New CTempMsg
        pm.Add "item-height", 34&
        Set pc = new_BPrefsControl("listbox", "cb>apps", "", "", "0", pm)
        pc.SizeTo 0, (9 * 34) + 2
        .Add pc

        .Add new_BPrefsControl("fancytoolbar", "ftb>app", "Configure Events...|Details|Unregister|Allow||Quit")
        .Add new_BPrefsControl("fancybutton2", "the_configure_app_button", "Configure...")

'        Set pm = New CTempMsg
'        pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\info_emblem.png"
'        Set pc = new_BPrefsControl("labelex", "app_info_label", "To prevent extensions from being loaded, hold down the 'E' key while Snarl loads.", , , pm)
'        pc.SizeTo 0, 64
'        .Add pc

    End With

    ' /* now, update the app list content */

    uUpdateAppList

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pc As BControl
Dim sz As String
Dim lr As VbMsgBoxResult

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()

    Case "cb>apps"
        Set mCurrentApp = g_AppRoster.AppAt(Val(Value))
        If Not (mCurrentApp Is Nothing) Then

'            ' /* set the "use one class" toggle button, this updates the class list */
'            prefskit_SetValue mPage.Panel, "ft>use_same", IIf(mCurrentApp.UseSameSettingsForAllClasses, "1", "0")
'            prefskit_SetItemEx mPage.Panel, "app_info_label", "image-file", , mCurrentApp.RealIconPath

            If mPage.Find("the_configure_app_button", pc) Then
                pc.SetText mCurrentApp.Name & " Settings..."
                pc.SetEnabled (mCurrentApp.HasConfig)

            End If

'            If mPage.Find("the_app_toolbar", pc) Then
'                prefskit_SetItem pc, 1, "enabled", (mCurrentApp.HasConfig())
'                prefskit_SetItem pc, 1, "text", mCurrentApp.Name & " Settings..."
'                 /* R2.4.2 DR3: "quit" button */
'                prefskit_SetItem pc, 2, "enabled", (mCurrentApp.IsQuittable())
'                prefskit_SetItem pc, 2, "text", "Close " & mCurrentApp.Name
'
'            End If

            If mPage.Find("ftb>app", pc) Then
                prefskit_SetItem pc, 2, "enabled", (mCurrentApp.HasAbout)
'                prefskit_SetItem pc, 3, "enabled", (mCurrentApp.HasConfig)
'                prefskit_SetItem pc, 3, "text", mCurrentApp.Name & " Settings..."
                ' /* V42.21: "Unregister" button */
                prefskit_SetItem pc, 3, "enabled", (mCurrentApp.Token <> gSnarlToken)
                ' /* "revoke ban" */
                prefskit_SetItem pc, 4, "enabled", (mCurrentApp.Token <> gSnarlToken)
                prefskit_SetItem pc, 4, "text", IIf(mCurrentApp.IsBanned(), "Allow", "Ban")
                ' /* R2.4.2 DR3: "quit" button */
                prefskit_SetItem pc, 6, "enabled", (mCurrentApp.IsQuittable())

            Else
                Debug.Print "TAppsPage.ControlChanged(): 'ftb>app' control not found"

            End If

        Else
            Debug.Print "TAppsPage.ControlChanged(): error setting current selected app"

        End If

    Case "ftb>app"
        If Not (mCurrentApp Is Nothing) Then
            Select Case Val(Value)
            Case 1
                frmAbout.ShowClassConfigPanel mPage.Panel.hWnd, mCurrentApp, ""

            Case 2
                ' /* details (about)*/
                mCurrentApp.DoAbout

            Case 3
                ' /* unregister */
                lr = MsgBox("Do you want to delete the settings for " & mCurrentApp.Name & "?  (If you're not sure, select 'No').", _
                            vbQuestion Or vbYesNoCancel, _
                            "Unregister Application")

                If (lr = vbYes) Or (lr = vbNo) Then _
                    g_AppRoster.Unregister mCurrentApp.Token, mCurrentApp.Password, (lr = vbYes)

            Case 4
                ' /* block/unblock */
                mCurrentApp.SetBlocked Not mCurrentApp.IsBanned
                uUpdateAppList

            Case 6
                ' /* quit */
                mCurrentApp.DoQuitPlease

            End Select

        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()

    Case "the_configure_app_button"
        mCurrentApp.DoSettings mPage.Panel.hWnd

    Case "cb>apps"
        frmAbout.ShowClassConfigPanel mPage.Panel.hWnd, mCurrentApp, ""

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

    Select Case Control.GetName()
    Case "cb>apps"
        ' /* applications list box */
        Select Case Notification
        Case "update_list"
            uUpdateAppList False

        End Select

    End Select

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 0
    
End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing
    Set mCurrentApp = Nothing

End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub uUpdateAppList(Optional ByVal AndSelectFirstItem As Boolean)

    If (mPage Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    If Not mPage.Find("cb>apps", pc) Then _
        Exit Sub

Dim sv As String
Dim sz As String
Dim i As Long

    sv = pc.GetValue()

    With g_AppRoster
        If .CountApps() Then
            For i = 1 To .CountApps()
                sz = sz & .AppAt(i).NameEx & IIf(.AppAt(i).IsBanned, " (Banned)", "") & "#?" & CStr(i) & "#?" & "" & IIf(i < .CountApps(), "|", "")
            
            Next i
        End If
    End With

    ' /* set the content */
    pc.SetText sz

    ' /* set the icons */
    With g_AppRoster
        If .CountApps() Then
            For i = 1 To .CountApps()
                If .AppAt(i).IsBanned Then
                    sz = g_MakePath(App.Path) & "etc\icons\ext-error.png"

                Else
                    sz = .AppAt(i).RealIconPath
'                    sz = g_TranslateIconPath(.AppAt(i).Icon, "")
'                    Debug.Print .AppAt(i).Icon & " >> " & sz

                End If
                prefskit_SetItem pc, i, "image-file", sz

            Next i
        End If
    End With

    If AndSelectFirstItem Then
        pc.SetValue "1"

    Else
        If Val(sv) > Val(pc.DoExCmd(B_COUNT_ITEMS)) Then _
            sv = pc.DoExCmd(B_COUNT_ITEMS)
        
        pc.SetValue sv

    End If

End Sub

