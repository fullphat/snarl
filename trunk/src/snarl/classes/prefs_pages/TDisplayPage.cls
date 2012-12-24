VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDisplayPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

Dim pdsp As TDisplaySubPage

    Set pm = New CTempMsg
    pm.Add "height", 360
    Set pc = new_BPrefsControl("tabstrip", "ts>display", , , , pm)

    Set pdsp = New TDisplaySubPage
    pdsp.Name = "def"
    BTabStrip_AddPage pc, "Appearance", new_BPrefsPage("pg>" & pdsp.Name, , pdsp)

    Set pdsp = New TDisplaySubPage
    pdsp.Name = "lay"
    BTabStrip_AddPage pc, "Layout", new_BPrefsPage("pg>" & pdsp.Name, , pdsp)

    Set pdsp = New TDisplaySubPage
    pdsp.Name = "vis"
    BTabStrip_AddPage pc, "Behaviour", new_BPrefsPage("pg>" & pdsp.Name, , pdsp)

'    Set pdsp = New TDisplaySubPage
'    pdsp.Name = "aud"
'    BTabStrip_AddPage pc, "Sounds", new_BPrefsPage("pg>" & pdsp.Name, , pdsp)

    mPage.Add pc


'    ' /* notification settings */
'
'    mPage.Add new_BPrefsControl("banner", "", "Notification Settings")
'
'    ' /* opacity */
'
'    Set pm = New CTempMsg
'    pm.Add "min", 1&
'    pm.Add "max", 100&
'    pm.Add "freq", 10&
'    pm.Add "suffix", "%"
'    mPage.Add new_BPrefsControl("fancyslider", "global_opacity", "", "Opacity:", CStr(gPrefs.global_opacity), pm)
'
'    ' /* spacing */
'
'    Set pm = New CTempMsg
'    pm.Add "min", -20&
'    pm.Add "max", 20&
'    pm.Add "freq", 5&
'    mPage.Add new_BPrefsControl("fancyslider", "margin_spacing", "", "Spacing:", CStr(gPrefs.margin_spacing), pm)
'
'    ' /* dropshadow */
'
'    mPage.Add new_BPrefsControl("fancytoggle2", "ft>use_dropshadow", "Add dropshadow to notifications?", "", IIf(gPrefs.use_dropshadow, "1", "0"), pm)
'
'    Set pm = New CTempMsg
'    pm.Add "min", 1&
'    pm.Add "max", 8&
'    pm.Add "freq", 1&
'    mPage.Add new_BPrefsControl("fancyslider", "dropshadow_strength", "", "Shadow Strength:", CStr(gPrefs.dropshadow_strength), pm)
'
'    ' /* icon theme */
'
'    mPage.Add new_BPrefsControl("combo", "c>icon_theme", "")
'    uUpdateIconThemes
'
'
'
'    ' /* font smoothing */
'
'    mPage.Add new_BPrefsControl("banner", "", "Font Smoothing")
'    mPage.Add new_BPrefsControl("fancycycle", "font_smoothing", "MelonType#?0|None#?1|Antialiased#?2|ClearType#?3|Windows Default#?4", "Mode:", CStr(gPrefs.font_smoothing + 1))
'
'    Set pm = New CTempMsg
'    pm.Add "min", 1&
'    pm.Add "max", 100&
'    pm.Add "freq", 10&
'    pm.Add "suffix", "%"
'    mPage.Add new_BPrefsControl("fancyslider", "melontype_contrast", "", "Contrast:", CStr(gPrefs.melontype_contrast), pm)
'
'
'    ' /* class defaults */
'
'    mPage.Add new_BPrefsControl("banner", "", "Class Defaults")
'    mPage.Add new_BPrefsControl("label", "", "These settings are used as the default values for any uncustomised classes.")
'    mPage.Add new_BPrefsControl("fancybutton2", "fb2>class_defaults", "Configure")



'    mPage.Add new_BPrefsControl("separator", "")
    mPage.Add new_BPrefsControl("fancybutton2", "test_display_settings", "Test Settings")

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
'Dim fUpdate As Boolean
'Dim pc As BControl
'Dim i As Long
'
'    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"
'
'    Select Case Control.GetName()
'
''    Case "default_position"
''        gPrefs.default_position = Val(Control.GetValue()) - 1
'
'    Case "font_smoothing"
'        gPrefs.font_smoothing = Val(Control.GetValue()) - 1
'        If mPage.Find("melontype_contrast", pc) Then _
'            pc.SetEnabled (Val(Value) = 1)
'
'        fUpdate = True
'
'    Case "melontype_contrast"
'        gPrefs.melontype_contrast = Val(Control.GetValue())
'        fUpdate = True
'
'    Case "global_opacity"
'        gPrefs.global_opacity = Val(Value)
'        fUpdate = True
'
'    Case "margin_spacing"
'        gPrefs.margin_spacing = Val(Value)
'        fUpdate = True
'
'    Case "ft>use_dropshadow"
'        gPrefs.use_dropshadow = (Val(Value) <> 0)
'        If mPage.Find("dropshadow_strength", pc) Then _
'            pc.SetEnabled (Val(Value) <> 0)
'
'        fUpdate = True
'
'    Case "dropshadow_strength"
'        gPrefs.dropshadow_strength = Val(Value)
'        fUpdate = True
'
'    Case "c>icon_theme"
'        i = Val(Value)
'        If i = 1 Then
'            gPrefs.icon_theme = ""
'
'        Else
'            gPrefs.icon_theme = gIconTheme(i - 1).Name
'
'        End If
'        fUpdate = True
'
'    End Select
'
'    If fUpdate Then _
'        g_WriteConfig

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    If Control.GetName() = "test_display_settings" Then _
        g_PrivateNotify "", "Settings Test", "This is a test of the current display settings", 0, , 1, , , NF_REMOTE Or NF_SECURE, True

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

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

'Private Sub uUpdateIconThemes()
'
'    If (mPage Is Nothing) Then _
'        Exit Sub
'
'Dim pc As BControl
'Dim sz As String
'Dim i As Long
'Dim n As Long
'
'    If Not (mPage.Find("c>icon_theme", pc)) Then _
'        Exit Sub
'
'    sz = "(Default Theme)|"
'
'    If gIconThemes Then
'        For i = 1 To gIconThemes
'            If LCase$(gIconTheme(i).Name) = LCase$(gPrefs.icon_theme) Then _
'                n = i
'
'            sz = sz & gIconTheme(i).Name & "|"
'
'        Next i
'
'    End If
'
'    pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)
'    pc.SetValue CStr(n + 1)
'
'End Sub




