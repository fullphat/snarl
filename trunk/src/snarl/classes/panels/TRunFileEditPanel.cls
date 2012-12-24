VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRunFileEditPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRunEditFilePanel.cls
    '/
    '/  Description:    Provides the Run File scheme configuration panel
    '/
    '/  Created:        3-Jun-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mPanel As BPrefsPanel
Dim mScheme As TRunFileScheme
Dim mhWndOwner As Long

Public Event Done()

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Select Case Control.GetName()
    Case "version"
        uSetHint

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim prf As TRunFileScheme

    Select Case Control.GetName
'    Case "go_user_guide"
'        ShellExecute Control.Page.Panel.hWnd, "open", g_MakePath(App.Path) & "etc\RunFileUserGuide\index.html", vbNullString, vbNullString, 5

    Case "save"
        With mScheme
            .SetVersion Val(prefskit_GetValue(mPanel, "version"))
            .SetTemplate prefskit_GetValue(mPanel, "template")
            .SetTarget prefskit_GetValue(mPanel, "target")
            .SetName prefskit_SafeGetText(mPanel, "name")

        End With

        mPanel.Quit

    End Select

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

    RaiseEvent Done
    Set mPanel = Nothing
    EnableWindow mhWndOwner, -1
    g_ShowWindow mhWndOwner, True, True

End Sub

Private Sub KPrefsPanel_Ready()

    uSetHint

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Function Go(ByVal hWndOwner As Long, ByRef Scheme As TRunFileScheme) As Long
Dim pPage As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

    mhWndOwner = hWndOwner
    Set mScheme = Scheme

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle g_Quote(Scheme.Name) & " Settings"
        .SetWidth 500
        .SetWindow 1

        Set pPage = new_BPrefsPage("", , Me)
        With pPage
            .SetMargin 100
            .Add new_BPrefsControl("fancyedit", "name", mScheme.Name, "Name:")
            .Add new_BPrefsControl("fancyfilepicker", "target", mScheme.Target, "Target file:")
            .Add new_BPrefsControl("fancycycle", "version", "Simple|Custom|Unabridged", "Type:", CStr(mScheme.Version))

            Set pc = new_BPrefsControl("label", "hint")
            pc.SizeTo 0, 60
            .Add pc

            .Add new_BPrefsControl("fancyedit", "template", mScheme.Template, "Template:")

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "save", "Save")

        End With

        .AddPage pPage
        .Go

        g_SetWindowIconToAppResourceIcon .hWnd

        uRefreshList

    End With

Dim dw As Long

    dw = GetWindowLong(mPanel.hWnd, GWL_STYLE)
    dw = dw And (Not WS_MINIMIZEBOX)
    SetWindowLong mPanel.hWnd, GWL_STYLE, dw

Dim rcThis As RECT
Dim rcOwner As RECT

    SetWindowLong mPanel.hWnd, GWL_HWNDPARENT, hWndOwner
    EnableWindow hWndOwner, 0

    GetWindowRect hWndOwner, rcOwner
    GetWindowRect mPanel.hWnd, rcThis
    g_RectNormalise rcThis

    g_MoveWindow mPanel.hWnd, _
                 rcOwner.Left + Fix(((rcOwner.Right - rcOwner.Left) - rcThis.Right) / 2), _
                 rcOwner.Top + Fix(((rcOwner.Bottom - rcOwner.Top) - rcThis.Bottom) / 2)

    g_ShowWindow mPanel.hWnd, True, True


End Function

Private Sub uAddNewEntry()
Dim sz As String

    sz = style_GetSnarlStylesPath()
    If sz = "" Then
        MsgBox "Couldn't create new entry: base path error", vbExclamation Or vbOKOnly, App.Title
        Exit Sub

    End If

    sz = g_MakePath(sz) & "runfile"
    If Not g_Exists(sz) Then
        MsgBox "Couldn't create new entry: runfile path error", vbExclamation Or vbOKOnly, App.Title
        Exit Sub

    End If

    With New CConfFile
        .SetTo g_MakePath(sz) & g_CreateGUID(True) & ".runfile"
        .Add "version=1"
        .Add "target=<enter_target_here>"
        .Add "template=<enter_template_here>"
        .Save

        ShellExecute mPanel.hWnd, "open", .Filename, vbNullString, vbNullString, SW_SHOW

    End With

    g_RunFileLoadSchemes
    uRefreshList

End Sub

Private Sub uRemEntry(ByVal Index As Long)
Dim prf As TRunFileScheme

    Set prf = gRunFiles.TagAt(Index)
    DeleteFile prf.File
    gRunFiles.Remove Index

    uRefreshList

End Sub

Private Sub uRefreshList()
Dim pc As BControl

    If Not mPanel.Find("list", pc) Then _
        Exit Sub

Dim sz As String
Dim n As Long

Dim pt As BTagItem

    n = Val(pc.GetValue())

    With gRunFiles
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            sz = sz & pt.Name & "|"

        Loop

    End With

    pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)
    pc.SetValue CStr(MAX(n, 1))

End Sub

Private Sub uSetHint()

    Select Case mScheme.Version
    Case 1
        prefskit_SafeSetText mPanel, "hint", _
                             "Content is sent as " & _
                             """<notification title>"" ""<notification text>"" ""<sender>""" & vbCrLf & _
                             "Any elements prefixed with either ""data-"" or ""value-"" will also be included as additional arguments."

    Case 2
        prefskit_SafeSetText mPanel, "hint", _
                             "The template defines the notification elements and the order in which they are sent, " & _
                             "with entries surrounded by % symbols replaced with the corresponding notification element (e.g. " & _
                             "%title% will be replaced with the notification title)."

    Case 3
        prefskit_SafeSetText mPanel, "hint", _
                             "Notification content is sent unabridged as a single argument in URL form (e.g. ""title=Hello, world!&&text=From Snarl"")."

    End Select

End Sub
