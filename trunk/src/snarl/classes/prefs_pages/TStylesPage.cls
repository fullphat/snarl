VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStylesPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPage As BPrefsPage
Dim mStyle As TStyle
Dim mScheme As String

Implements KPrefsPage
Implements MWndProcSink

Private Sub KPrefsPage_AllAttached()

    uUpdateStyleList

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

    mPage.Add new_BPrefsControl("label", "", "The following styles are installed:")

    Set pm = New CTempMsg
    pm.Add "item-height", 36&
    mPage.Add new_BPrefsControl("combo", "installed_styles", , , , pm)

    ' /* style description */
    Set pc = new_BPrefsControl("label", "lbl>desc", " ")
    pc.SizeTo 0, 26
    mPage.Add pc

'    ' /* settings button */
'    mPage.Add new_BPrefsControl("fancybutton2", "fb>cfg", "Settings")

    ' /* 'the <style> has <n> schemes' label */
    mPage.Add new_BPrefsControl("label", "lbl>schemes", "   ")


    Set pm = New CTempMsg
    pm.Add "item-height", 24&
    mPage.Add new_BPrefsControl("listbox", "installed_schemes", , , , pm)



'        Set pm = New CTempMsg
'        pm.Add "image-file", g_MakePath(App.Path) & "icons\snarl.png"
'        pm.Add "image-height", 32
'        pm.Add "valign", "centre"
'        .Add new_BPrefsControl("labelex", "", "Snarl R2.1 (V39.1) © 2005-2008 k23 productions", , , pm)
'
'        With New CFolderContent2
'            If .SetTo(g_MakePath(App.Path) & "icons\pics") Then
'                pm.Replace "image-file", g_MakePath(.Path) & .EntryAt((Rnd * (.CountEntries - 1)) + 1)
'
'            End If
'
'        End With

    pm.Add "scale_to_fit", 1&

    Set pc = new_BPrefsControl("image", "image_preview", "", , , pm)
    pc.SizeTo 0, 100
    mPage.Add pc

    mPage.Add new_BPrefsControl("fancytoolbar", "ftb>style", "Configure...|Preview|Preview Priority|Set as Default||Advanced...")

'    mPage.Add new_BPrefsControl("label", "current_default_style", "Default style is " & g_Capitalize(Replace$(g_ConfigGet("default_style"), "/", " (") & ")"))

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pInfo As T_NOTIFICATION_INFO
Dim pc As BControl
Dim pm As CTempMsg
Dim sz As String
Dim dw As Long
Dim szScheme As String
Dim szText As String

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()

'    Case "installed_styles"
'        Set mStyle = g_StyleRoster.StyleAt(Val(Value))
'        If (mStyle Is Nothing) Then _
'            Exit Sub
'
'        If Control.Page.Find("lbl>schemes", pc) Then _
'            pc.SetText "This style has " & CStr(mStyle.CountSchemes) & " scheme" & IIf(mStyle.CountSchemes = 1, "", "s") & ":"
'
''            pc.SetText "The " & mStyle.Name & " style has " & CStr(mStyle.CountSchemes) & " scheme" & IIf(mStyle.CountSchemes = 1, "", "s") & ":"
'
''        If Control.Page.Find("fb>cfg", pc) Then _
'            pc.SetEnabled (mStyle.Flags And S_STYLE_IS_CONFIGURABLE)
'
'        If Control.Page.Find("ftb>style", pc) Then
'            Set pm = New CTempMsg
'            pm.Add "index", 1&
'            pm.Add "enabled", IIf((mStyle.Flags And S_STYLE_IS_CONFIGURABLE), 1&, 0&)
'            pc.DoExCmd B_SET_ITEM, pm
'
'        End If
'
''            pc.SetEnabled (mStyle.Flags And S_STYLE_IS_CONFIGURABLE)
'
'
'
'        uUpdateSchemeList
'
'        If Control.Page.Find("lbl>desc", pc) Then _
'            pc.SetText mStyle.Description
'
'

'    Case "installed_schemes"
'
''MsgBox "changed!"
'
'        If mStyle.Name = "" Then _
'            Exit Sub
'
'        Set pm = New CTempMsg
'        mScheme = mStyle.SchemeAt(Val(Value))
'        If mScheme = "<Default>" Then _
'            mScheme = ""
'
'        sz = g_MakePath(mStyle.Path) & mScheme & ".preview"
'        If Not g_Exists(sz) Then
'            If uCreatePreviewImage() Then
'                sz = uMakePreviewName()
'
'            Else
'                sz = g_MakePath(mStyle.Path) & "style.preview"
'
'            End If
'
'        End If
'
'        If Not g_Exists(sz) Then _
'            sz = g_MakePath(App.Path) & "etc\icons\blank.png"
'
'        pm.Add "image-file", sz
'        If Control.Page.Find("image_preview", pc) Then _
'            pc.DoExCmd B_SET_ITEM, pm
'
'
'    Case "ftb>style"
'        Select Case Val(Value)
'        Case 1
'            ' /* configure style */
'            uDoConfigureStyleWindow
'
'        Case 2, 3
'            ' /* preview or preview priority */
'            If (mStyle.Name = "") Or (Not mPage.Find("installed_schemes", pc)) Then _
'                Exit Sub
'
'            ' /* must have a selected scheme - this should *always* be the case, but... */
'
'            dw = Val(pc.GetValue())
'            If dw <> 0 Then _
'                g_DoSchemePreview2 mStyle.Name, mStyle.SchemeAt(dw), (Val(Value) = 3), 0
'
'        Case 4
'            ' /* make default */
'            If (mStyle.Name = "") Or (Not mPage.Find("installed_schemes", pc)) Then _
'                Exit Sub
'
'            ' /* must have a selected scheme - this should *always* be the case, but... */
'
'            dw = Val(pc.GetValue())
'            If dw = 0 Then _
'                Exit Sub
'
'            szScheme = LCase$(mStyle.Name)
'
'            ' /* if we have a scheme other than 'default' adjust the name accordingly... */
'
'            If mStyle.SchemeAt(dw) <> "" Then _
'                szScheme = szScheme & "/" & LCase$(mStyle.SchemeAt(dw))
'
'            g_ConfigSet "default_style", szScheme
'            g_Debug "TStylePage.ControlInvoked(): default style is now '" & g_ConfigGet("default_style") & "'"
'
''            If Control.Page.Panel.Find("current_default_style", pc) Then _
'                pc.SetText "Default style is " & g_Capitalize(Replace$(g_ConfigGet("default_style"), "/", " (") & ")")
'
'
'        Case 6
'            ' /* style engine info */
'
'            With New TStyleEnginePanel
'                .Go mPage.Panel.hWnd
'
'            End With
'
''Dim pEngineInfo As T_SNARL_STYLE_ENGINE_INFO
''
''            With g_StyleRoster
''                For dw = 1 To .CountEngines
''                    If .EngineAt(dw, pEngineInfo) Then
''                        With pEngineInfo
''                            sz = sz & IIf(.Name = "", "<unknown>", .Name) & " V" & CStr(.Version) & "." & CStr(.Revision) & " " & .Description & " " & .Path & vbCrLf
''
''                        End With
''                    End If
''                Next dw
''
''            End With
''
''            MsgBox sz, vbOKOnly Or vbInformation, App.Title
'
'        End Select
'
    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)
'Dim pc As BControl
'
'    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"
'
'    Select Case Control.GetName()
'
'    Case "fb>cfg"
'        ' /* configure style */
'        uDoConfigureStyleWindow
'
'
''    Case "preview_style"
''        If (mStyle.Info.Name = "") Or (Not mPage.Find("installed_schemes", pc)) Then _
''            Exit Sub
''
''        ' /* must have a selected scheme - this should *always* be the case, but... */
''
''        dw = Val(pc.GetValue())
''        If dw = 0 Then _
''            Exit Sub
''
''        ' /* if we have a scheme other than 'default' adjust the text accordingly... */
''
''        If mStyle.Scheme(dw) <> "" Then
''            szText = " using the " & mStyle.Scheme(dw) & " scheme"
''            szScheme = "/" & LCase$(mStyle.Scheme(dw))
''
''        End If
''
''        szText = "This is a preview of the " & mStyle.Info.Name & " style" & szText
''        szScheme = LCase$(mStyle.Info.Name) & szScheme
''
'''        MsgBox "*****> " & mStyle.Info.Path
''
''        g_NotificationRoster.Add 0, "", "Scheme Preview", szText, E_DEFAULT_DURATION, _
''                                 g_MakePath(App.Path) & "etc\icons\style_preview.png", 0, 0, "", szScheme, ""
''
''
''    Case "set_default_style"
''        If (mStyle.Info.Name = "") Or (Not mPage.Find("installed_schemes", pc)) Then _
''            Exit Sub
''
''        ' /* must have a selected scheme - this should *always* be the case, but... */
''
''        dw = Val(pc.GetValue())
''        If dw = 0 Then _
''            Exit Sub
''
''        szScheme = LCase$(mStyle.Info.Name)
''
''        ' /* if we have a scheme other than 'default' adjust the name accordingly... */
''
''        If mStyle.Scheme(dw) <> "" Then _
''            szScheme = szScheme & "/" & LCase$(mStyle.Scheme(dw))
''
''
''        gPrefs.default_style = szScheme
''        g_Debug "TStylePage.ControlInvoked(): default style is now '" & gPrefs.default_style & "'"
''        g_WriteConfig
'
''    Case "restart_style_roster"
''        If Not (g_StyleRoster Is Nothing) Then
''            melonLibClose g_StyleRoster
''            Sleep 500
'''            MsgBox "Click OK when you're ready for the Style Roster to start up", vbInformation Or vbOKOnly, App.Title
''            melonLibOpen g_StyleRoster
''            uUpdateStyleList
''
''        End If
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

    If (Control.GetName() = "installed_styles") And (Notification = "update_list") Then _
        uUpdateStyleList

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

Private Sub uUpdateStyleList()
Dim pc As BControl

    If (Not (mPage.Find("installed_styles", pc))) Or (g_StyleRoster Is Nothing) Then _
        Exit Sub

Dim px As TStyle
Dim sz As String

    ' /* build the list content */

    With g_StyleRoster
        .Rewind
        Do While .GetNextStyle(px)
            With px
                sz = sz & .Name & " " & .Release & "#?0#?" & _
                          IIf(.Copyright <> "", .Copyright & " ", "") & _
                          IIf(.URL <> "", "(" & .URL & ")", "") & _
                          "|"

            End With
        Loop

    End With

    ' /* set the combo's content (this also clears the previous content) */

    sz = g_SafeLeftStr(sz, Len(sz) - 1)
    If sz = "" Then
        pc.SetText "<No styles>"
        pc.SetEnabled False

        ' /* also clear the schemes list and the label */

        If mPage.Find("installed_schemes", pc) Then _
            pc.SetText ""

        If mPage.Find("lbl>schemes", pc) Then _
            pc.SetText ""

        Exit Sub

    Else
        pc.SetEnabled True
        pc.SetText sz

    End If

    ' /* set the icons */

Dim pm As CTempMsg
Dim i As Long

    Set pm = New CTempMsg

    With g_StyleRoster
        If .CountStyles() Then
            For i = 1 To .CountStyles
                Set px = .StyleAt(i)
                sz = px.IconPath

                ' /* sort of backwards-compatability/kludge: if a style doesn't appear to have
                '    a specific icon path, create a default one */

'                MsgBox .StyleAt(i).Name & " >> " & sz

                If sz = "" Then _
                    sz = g_MakePath(px.Path) & px.Name & ".png"

                ' /* use default icon if custom one not found */

                If Not g_Exists(sz) Then _
                    sz = g_MakePath(App.Path) & "etc\icons\style.png"

                pm.Replace "index", i
                pm.Replace "image-file", sz
                pc.DoExCmd B_SET_ITEM, pm

            Next i
        End If
    End With

    pc.SetValue "1"

End Sub

Private Sub uUpdateSchemeList()
Dim pc As BControl

    If (Not (mPage.Find("installed_schemes", pc))) Or (mStyle Is Nothing) Then _
        Exit Sub

    ' /* set the listbox content */

    pc.SetText mStyle.GetSchemeList(True)
'    MsgBox "selected item: " & pc.GetValue

    ' /* set the icons */

Dim pm As CTempMsg
Dim i As Long

    Set pm = New CTempMsg
    With mStyle

        prefskit_SafeSetText mPage.Panel, "lbl>schemes", "This style has " & CStr(.CountSchemes) & " scheme" & IIf(.CountSchemes = 1, "", "s") & ":"

        If .CountSchemes Then
            For i = 1 To .CountSchemes
                pm.Replace "index", i
                pm.Replace "image-file", g_MakePath(.Path) & .SchemeAt(i) & ".icon"
                pc.DoExCmd B_SET_ITEM, pm

            Next i
        End If
    End With

    pc.SetValue "1"

End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
'Dim ps As TStyle
'
'    On Error Resume Next
'
'    If uMsg = WM_DESTROY Then
'
'        ' /* config window has gone */
'
'        RemoveSubClass hWnd                     ' // really necessary?
'        EnableWindow mPage.Panel.hWnd, -1
'
'        ' /* it's likely that the schemes will have changes so refresh them now.  This
'        '    is _really_ ugly code and needs to be sorted out once a better style
'        '    type library has been developed */
'
''        If Not (g_StyleRoster Is Nothing) Then
''            If g_StyleRoster.RefreshStyle(mStyle.Info.Name, ps) Then _
''                LSet mStyle = ps
''
''        End If
'
'        uUpdateSchemeList
'        g_ShowWindow mPage.Panel.hWnd, True, True
'
'Dim pc As BControl
'
'        If mPage.Find("installed_schemes", pc) Then _
'            KPrefsPage_ControlChanged pc, pc.GetValue
'
'    End If

End Function

'Private Function uCreatePreviewImage() As Boolean
'
'    If (mStyle Is Nothing) Then _
'        Exit Function
'
'    If (mStyle.Flags And S_STYLE_IS_WINDOWLESS) Then _
'        Exit Function
'
'    DeleteFile uMakePreviewName()
'
'Dim pInfo As T_NOTIFICATION_INFO
'
'    With pInfo
'        .Title = "Scheme Preview"
'        .Text = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do..." ' mScheme
'        .Timeout = 10
'        .IconPath = mStyle.IconPath
'        .StyleName = mStyle.Name
'        .SchemeName = LCase$(mScheme)
'        ' /* R2.4 DR7 */
'        Set .ClassObj = New TAlert
'        ' /* R2.4 DR8 */
'        Set .Actions = new_BTagList()
'
'    End With
'
'    If (mStyle.Flags And S_STYLE_V42_CONTENT) Then _
'        g_KludgeNotificationInfo pInfo
'
'Dim pw As CSnarlWindow
'
'    Set pw = New CSnarlWindow
'    If pw.Create(pInfo, mStyle.CreateInstance(), mStyle.Flags, g_GetStylePath(mStyle.Name)) = M_OK Then
'        pw.CaptureToFile uMakePreviewName()
'        pw.Remove
'        uCreatePreviewImage = True
'
'    Else
'        Debug.Print "create failed"
'
'    End If
'
'End Function
'
'Private Function uMakePreviewName() As String
'
'    If (mStyle Is Nothing) Then _
'        Exit Function
'
'    uMakePreviewName = g_MakeValidPath(g_MakePath(mStyle.Path) & mStyle.Name & "-" & mScheme & ".prvw")
'
'End Function

'Private Function uDoConfigureStyleWindow() As Boolean
'
'    If (mStyle.Engine Is Nothing) Then _
'        Exit Function
'
'Dim hWndConfig As Long
'
'    hWndConfig = mStyle.Engine.GetConfigWindow(mStyle.Name)
'    If IsWindow(hWndConfig) = 0 Then
'        g_Debug "TStylePage.ControlChanged(): ConfigRequested() failed (" & g_HexStr(hWndConfig) & ")"
''        MsgBox "This style cannot be configured.", vbInformation Or vbOKOnly, App.Title
'        Exit Function
'
'    End If
'
'    ' /* valid window */
'
'Dim rcConfig As RECT
'Dim rc As RECT
'
'    AddSubClass hWndConfig, Me
'    SetWindowLong hWndConfig, GWL_HWNDPARENT, mPage.Panel.hWnd
'    EnableWindow mPage.Panel.hWnd, 0
'
'    GetWindowRect mPage.Panel.hWnd, rc
'    GetWindowRect hWndConfig, rcConfig
'    g_RectNormalise rcConfig
'
'    g_MoveWindow hWndConfig, _
'                 rc.Left + Fix(((rc.Right - rc.Left) - rcConfig.Right) / 2), _
'                 rc.Top + Fix(((rc.Bottom - rc.Top) - rcConfig.Bottom) / 2)
'
'    g_ShowWindow hWndConfig, True, True
'    uDoConfigureStyleWindow = True
'
'End Function
