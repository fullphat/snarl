VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBetaPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TBetaPanel.cls
'/
'/  Description:    Modal Beta-release/Update Available panel
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Dim mPanel As BPrefsPanel
Dim mPage As BPrefsPage

Dim mNewVersion As String
Dim mUpdatePanel As Boolean
Dim mhWndPanel As Long

Implements KPrefsPanel
Implements KPrefsPage

Private Sub Class_Terminate()

    If Not (mPanel Is Nothing) Then _
        mPanel.Quit

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    If mUpdatePanel Then
        ' /* update available panel */

        mPage.Add new_BPrefsControl("label", "", "A new version of Snarl is available!")
        mPage.Add new_BPrefsControl("label", "", mNewVersion & " is now available (you have " & App.Comments & ").")

        mPage.Add new_BPrefsControl("label", "", "Release Notes:")
        Set pc = new_BPrefsControl("rtf", "")
        Set pm = New CTempMsg
        pm.Add "file", g_MakePath(App.Path) & ".changelog"
        pc.DoExCmd "load", pm
        mPage.Add pc

        mPage.Add new_BPrefsControl("fancybutton2", "fb2>get", "Download Now")

    Else
        ' /* beta notice panel */

        Set pc = new_BPrefsControl("rtf", "fb2>ok", "")
        Set pm = New CTempMsg
        pm.Add "file", g_MakePath(App.Path) & "beta.rtf"
        pc.DoExCmd "load", pm
        mPage.Add pc

        mPage.Add new_BPrefsControl("fancybutton2", "fb2>new", "What's new?")
        mPage.Add new_BPrefsControl("seperator", "")
        mPage.Add new_BPrefsControl("fancybutton2", "fb2>ok", "Don't show this again")

    End If

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pc As BControl

    Select Case Control.GetName()
    Case "fb2>ok"
        g_Prefs.AgreeBetaUsage = True
        g_WriteConfig
        mPanel.Quit

    Case "fb2>new"
        ShellExecute 0, "open", g_MakePath(App.Path) & "whats new.txt", vbNullString, vbNullString, SW_SHOW

'    Case "fb2>get"
'        ' /* download update */
'        frmAbout.bStartUpdateDownload
'        mPanel.Quit

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
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

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

'    If mUpdatePanel Then
'        EnableWindow mhWndPanel, -1
'        ' /* delete the update change log file */
'        DeleteFile g_MakePath(App.Path) & ".changelog"
'
'        frmAbout.bDownloadPanelQuit
'
'    End If

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(Optional ByVal IsUpdatePanel As Boolean = False, Optional ByVal hWndPanel As Long, Optional ByVal NewVersion As String)

    mUpdatePanel = IsUpdatePanel
    mhWndPanel = hWndPanel
    mNewVersion = NewVersion

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetWindow 1
        .SetTitle IIf(IsUpdatePanel, "Snarl Update Available", App.Title & " " & App.Comments & " Public Beta Information")
        .SetWidth 500

        .AddPage new_BPrefsPage("", , Me)

        .Go
        g_SetWindowIconToAppResourceIcon .hWnd

        If IsUpdatePanel Then
            SetWindowLong .hWnd, GWL_HWNDPARENT, hWndPanel
            EnableWindow hWndPanel, 0

        End If

        g_MoveWindow .hWnd, 0, 0, , True
        g_ShowWindow .hWnd, True, True

    End With

End Sub





