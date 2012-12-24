VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TTinyPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements KPrefsPage
Implements IStylePage

Private Sub IStylePage_Init(Page As prefs_kit_d2.BPrefsPage)
Dim sz As String

    With Page

        .Add new_BPrefsControl("banner", "", "Appearance")
        .Add new_BPrefsControl("fancycycle", "tiny.iconsize", "Small|Medium|Large", "Icon size:", gSettings.GetValueWithDefault("tiny.iconsize", "1"))

        sz = uCreatePackedFontStr("Tahoma", 7)
        .Add new_BPrefsControl("fancyfontpicker", "tiny.font", "", "Font:", gSettings.GetValueWithDefault("tiny.font", sz))
        .Add new_BPrefsControl("fancytoggle2", "tiny.cleartype", "Use ClearType™ font aliasing?", "", gSettings.GetValueWithDefault("tiny.cleartype", "0"))

        .Add new_BPrefsControl("fancytoggle2", "tiny.useglass", "Glass effect?", "", gSettings.GetValueWithDefault("tiny.useglass", "1"))
        .Add new_BPrefsControl("fancytoggle2", "tiny.invertshading", "Invert shading?", "", gSettings.GetValueWithDefault("tiny.invertshading", "0"))

        .Add new_BPrefsControl("banner", "", "Colours")
        .Add new_BPrefsControl("fancycolourpicker", "tiny.background", "", "Background:", gSettings.GetValueWithDefault("tiny.background", CStr(rgba(0, 0, 0))))
        .Add new_BPrefsControl("fancycolourpicker", "tiny.foreground", "", "Text:", gSettings.GetValueWithDefault("tiny.foreground", CStr(rgba(255, 255, 255))))
        .Add new_BPrefsControl("fancycolourpicker", "tiny.border", "", "Border:", gSettings.GetValueWithDefault("tiny.border", CStr(rgba(252, 243, 0))))
        .Add new_BPrefsControl("fancycolourpicker", "tiny.shading", "", "Shading:", gSettings.GetValueWithDefault("tiny.shading", CStr(rgba(0, 0, 0))))

    End With

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    gSettings.Update Control.GetName, Value
    gSettings.Save

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pc As BControl
Dim szn As String
Dim szv As String

    Select Case Control.GetName
    Case "reset"

        With gTinyDefaults
            .Rewind
            Do While .GetNextItem(szn, szv)
                If Control.Page.Find(szn, pc) Then _
                    pc.SetValue szv

            Loop

        End With

        With Control.Page

'            If .Find("tiny.iconsize", pc) Then _
'                pc.SetValue "1"
'
'            If .Find("tiny.font", pc) Then _
'                pc.SetValue "name::Tahoma#?size::7#?bold::1"
'
'            If .Find("tiny.cleartype", pc) Then _
'                pc.SetValue "0"
'
'            If .Find("tiny.useglass", pc) Then _
'                pc.SetValue "1"
'
'            If .Find("tiny.invertshading", pc) Then _
'                pc.SetValue "0"
'
'
'            If .Find("tiny.background", pc) Then _
'                pc.SetValue CStr(rgba(0, 0, 0))
'
'            If .Find("tiny.foreground", pc) Then _
'                pc.SetValue CStr(rgba(255, 255, 255))
'
'            If .Find("tiny.border", pc) Then _
'                pc.SetValue CStr(rgba(252, 243, 255))
'
'            If .Find("tiny.shading", pc) Then _
'                pc.SetValue CStr(rgba(0, 0, 0))


        End With

    Case "test"
        snPrivatePreviewScheme "Tiny", "Top"

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Function uCreatePackedFontStr(ByVal Typeface As String, ByVal TypeSize As String) As String
Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.Add "name", Typeface
    ppd.Add "size", TypeSize
    uCreatePackedFontStr = ppd.AsString()

End Function

