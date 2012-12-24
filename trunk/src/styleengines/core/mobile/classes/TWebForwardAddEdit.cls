VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWebForwardAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TWebForwardAddEdit.cls
    '/
    '/  Description:    Provides the Web Forward Add/Edit panel
    '/
    '/  Created:        5-Jul-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Dim mhOwner As Long
Dim mPanel As BPrefsPanel
Dim mFile As CConfFile2

Public Event Quit()

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Select Case Control.GetName()
    Case "use_post"
        uSetValues

    Case "toolbar"
        Select Case Value
        Case "1"
            uDoPreview
        
        Case "3"
            uSave

'        Case "4"
'            ShellExecute mPanel.hWnd, "open", "https://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=WebForward_SDK", vbNullString, vbNullString, SW_SHOW

        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim szSrc As String
Dim sz As String

    Select Case Control.GetName()
    Case "pick_icon"
        With New CFileDialog
            .Title = "Choose Icon"
            .AddFilter "Pictures", "*.bmp;*.png;*.jpg;*.jpeg;*.gif;*.tga", True, True
            .AddFilter "All files", "*.*", False, False
            .DialogType = E_DIALOG_OPEN
            If .Go(True, E_FILE_DIALOG_CENTRE_SCREEN) Then
                szSrc = .SelectedFile
                sz = g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & g_MakeFilename2(mFile.ValueOf("name")) & ".icon"
                If g_Exists(sz) Then
                    If MsgBox("Replace existing icon?", vbQuestion Or vbYesNo, "WebForward") = vbNo Then _
                        Exit Sub

                End If
                CopyFile szSrc, sz, 0

            End If
        End With

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

    EnableWindow mhOwner, -1
    g_ShowWindow mhOwner, True, True
    RaiseEvent Quit

End Sub

Private Sub KPrefsPanel_Ready()

    uSetValues

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long, ByRef File As CConfFile2)
Dim pp As BPrefsPage

    mhOwner = hWndPanel

    If (File Is Nothing) Then
        Set mFile = New CConfFile2

    Else
        Set mFile = File

    End If

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle "WebForward: " & IIf(mFile.Filename = "", "Add new forward", "Editing " & mFile.ValueOf("name"))
        .SetWidth 500
        .SetWindow 1

        Set pp = new_BPrefsPage("", , Me)

        With pp
            .SetMargin 100

            .Add new_BPrefsControl("fancyedit", "name", mFile.ValueOf("name"), "Name:")
            .Add new_BPrefsControl("fancybutton2", "pick_icon", "Icon...")
            .Add new_BPrefsControl("fancyedit", "url", mFile.ValueOf("url"), "Base URL:")
            .Add new_BPrefsControl("label", "", "The remote path to the server, e.g.: ""http://www.myserver.local/do?""")
            
            .Add new_BPrefsControl("fancyedit", "template", mFile.ValueOf("template"), "Query:")
            .Add new_BPrefsControl("label", "", "This is appended to the base url, e.g. ""event=#title#&&description=#text#""")

            .Add new_BPrefsControl("fancytoggle2", "use_post", "Use POST? (Default is GET)", "Method:", mFile.ValueOf("use_post"))

            .Add new_BPrefsControl("fancyedit", "headers", mFile.ValueOf("headers"), "Headers:")
            .Add new_BPrefsControl("label", "", "Separate individual headers with ""\n"" markers, e.g. ""Accept-Language: en-US\nAccept: text/plain""")
            .Add new_BPrefsControl("fancyedit", "content", mFile.ValueOf("content"), "Content:")
            .Add new_BPrefsControl("fancytoggle2", "suppress_std_headers", "Suppress default headers (""From"" and ""User-Agent"")?", , mFile.ValueOf("suppress_std_headers"))

            .Add new_BPrefsControl("label", "", "Use #app# to retrieve the registered application name and #icon-base64# to insert the icon as Base64 encoded data.", "Note:")

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancytoolbar", "toolbar", "Preview||Save")
            
'            .Add new_BPrefsControl("fancybutton2", "go_user_guide", "User Guide")

'            .Add new_BPrefsControl("fancyfilepicker", "custom_headers", , "Template:")
'            .Add new_BPrefsControl("fancytoggle2", "use_custom_headers", "Use custom headers?")
'            .Add new_BPrefsControl("fancyfilepicker", "custom_headers", , "Template:")

'            .Add new_BPrefsSeparator
'            .Add new_BPrefsControl("fancybutton2", "save", "Save")


'            .Add new_BPrefsControl("fancycycle", "version", "Simple|Complex|Unabridged", "Type:", CStr(mScheme.Version))
            
        End With

        .AddPage pp

        .Go
        g_SetWindowIconToAppResourceIcon .hWnd


Dim dw As Long

        dw = GetWindowLong(.hWnd, GWL_STYLE)
        dw = dw And (Not WS_MINIMIZEBOX)
        SetWindowLong .hWnd, GWL_STYLE, dw

Dim rcThis As RECT
Dim rcOwner As RECT

        SetWindowLong .hWnd, GWL_HWNDPARENT, hWndPanel
        EnableWindow hWndPanel, 0
    
        GetWindowRect hWndPanel, rcOwner
        GetWindowRect .hWnd, rcThis
        g_RectNormalise rcThis
    
        g_MoveWindow .hWnd, _
                     rcOwner.Left + Fix(((rcOwner.Right - rcOwner.Left) - rcThis.Right) / 2), _
                     rcOwner.Top + Fix(((rcOwner.Bottom - rcOwner.Top) - rcThis.Bottom) / 2)
    
        g_ShowWindow .hWnd, True, True

    End With

End Sub

Private Function uCreateNewScheme(ByVal Name As String) As Boolean
Dim sz As String

    sz = g_MakePath(style_GetSnarlStylesPath()) & "webforward"
    If Not g_Exists(sz) Then _
        Exit Function

    sz = sz & "\" & g_MakeFilename(Name) & ".webforward"
    Debug.Print sz

    If g_Exists(sz) Then _
        Exit Function

    mFile.SetTo sz
    uCreateNewScheme = True

End Function

Private Sub uSetValues()
Dim sz As String

    sz = prefskit_GetValue(mPanel, "use_post")
    prefskit_SafeEnable mPanel, "content", (sz = "1")

End Sub

Private Sub uDoPreview()
Dim pf As CConfFile2

    Set pf = New CConfFile2
    With pf
        .Update "name", prefskit_SafeGetText(mPanel, "name")
        .Update "url", prefskit_SafeGetText(mPanel, "url")
        .Update "template", prefskit_SafeGetText(mPanel, "template")
        .Update "use_post", prefskit_GetValue(mPanel, "use_post")
        .Update "headers", prefskit_SafeGetText(mPanel, "headers")
        .Update "content", prefskit_SafeGetText(mPanel, "content")
        .Update "suppress_std_headers", prefskit_GetValue(mPanel, "suppress_std_headers")

    End With

Dim pi As notification_info

    ' /* mimic sample V42 content */

    With pi
        .Title = "WebForward"   ' // app name
        .Text = "title::WebForward#?text::Preview Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor " & _
                "incididunt ut labore et dolore magna aliqua.#?icon::" & g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "webforward.png"

    End With

    With New TWebForwardInstance
        .Go pf, pi

    End With

End Sub

Private Sub uSave()

    If (prefskit_SafeGetText(mPanel, "name") <> "") And (prefskit_SafeGetText(mPanel, "url") <> "") Then

        If mFile.Filename = "" Then
            ' /* new * /
            If Not uCreateNewScheme(prefskit_SafeGetText(mPanel, "name")) Then
                MsgBox "A scheme with that name already exists", vbExclamation Or vbOKOnly, "WebForward"
                Exit Sub
        
            End If

        End If

        With mFile
            .Update "name", prefskit_SafeGetText(mPanel, "name")
            .Update "url", prefskit_SafeGetText(mPanel, "url")
            .Update "template", prefskit_SafeGetText(mPanel, "template")
            .Update "use_post", prefskit_GetValue(mPanel, "use_post")
'                .Update "use_custom_headers", ""
            .Update "headers", prefskit_SafeGetText(mPanel, "headers")
'                .Update "use_custom_content", ""
            .Update "content", prefskit_SafeGetText(mPanel, "content")
            .Update "suppress_std_headers", prefskit_GetValue(mPanel, "suppress_std_headers")
            .Save

        End With

    Else
        MsgBox "Scheme must have at least a name and base URL", vbExclamation Or vbOKOnly, "WebForward"
        Exit Sub

    End If

    mPanel.Quit

End Sub




