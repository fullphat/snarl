VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const LIB_DATE = "2-Mar-2012"

Dim mStyles As TStyleBundle
Dim mLastErr As String

Implements IStyleEngine

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = mStyles.Count()

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance
Dim ps As TStyle

    If Not mStyles.Find(StyleName, ps) Then _
        Exit Function

Dim pMV As MTVInstance
Dim pIB As InfoBarInstance
Dim pt As TinyInstance

    Select Case LCase$(StyleName)
    Case "music video"
        Set pMV = New MTVInstance
        Set pMV.Style = ps
        Set IStyleEngine_CreateInstance = pMV

    Case "info bar"
        Set pIB = New InfoBarInstance
        Set pIB.Style = ps
        Set IStyleEngine_CreateInstance = pIB

    Case "tiny"
        Set pt = New TinyInstance
        Set pt.Style = ps
        Set IStyleEngine_CreateInstance = pt

    End Select

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = "Provides a number of single-instance styles which appear as banners"

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long

    IStyleEngine_GetConfigWindow = mStyles.Configure(StyleName)

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT
Dim ps As TStyle

    Set mStyles = New TStyleBundle
    With mStyles
    
        Set ps = New TStyle
        ps.SetTo "Music Video", "Displays the notification as a banner along the bottom of the screen.", _
                 S_STYLE_NO_DROPSHADOW Or S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE Or S_STYLE_CLICK_THROUGH Or S_STYLE_V42_CONTENT Or S_STYLE_IS_CONFIGURABLE, _
                 SCHEME_1, 1, 303, _
                 LIB_DATE, "info@fullphat.net", "http://getsnarl.info", _
                 mStyles.CreateDefaults("title_font=Arial Rounded MT Bold&title_size=26&title_bold=1&text_font=Arial Rounded MT Bold&text_size=12&col-background=" & CStr(rgba(24, 24, 24))), _
                 E_SF_CAN_CHANGE_BACKGROUND Or E_SF_CAN_CHANGE_TEXT_FONT Or E_SF_CAN_CHANGE_TITLE_FONT

        .Add ps
            
        Set ps = New TStyle
        ps.SetTo "Tiny", "Displays the notfication in a tiny bubble in the top right corner of the screen.", _
                 S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE Or S_STYLE_IS_CONFIGURABLE Or S_STYLE_V42_CONTENT, _
                 "Top|Bottom|Top Left|Top Right|Bottom Left|Bottom Right", 2, 0, _
                 LIB_DATE, "info@fullphat.net", "http://getsnarl.info", _
                 mStyles.CreateDefaults("title_font=Verdana&title_size=7&text_font=Verdana&text_size=7&use_cleartype=1&icon_size=24&use_glass=1&invert_shading=0&col-background=" & CStr(rgba(96, 32, 0)) & "&col-foreground=" & CStr(rgba(255, 255, 255)) & "&col-border=" & CStr("-10040320") & "&col-shading=" & CStr(rgba(0, 0, 0))), _
                 E_SF_CAN_CHANGE_BACKGROUND Or E_SF_CAN_CHANGE_BORDER Or E_SF_CAN_CHANGE_FONT_SMOOTHING Or E_SF_CAN_CHANGE_FOREGROUND Or E_SF_CAN_CHANGE_SHADING Or E_SF_CAN_CHANGE_TEXT_FONT Or E_SF_CAN_CHANGE_TITLE_FONT Or E_SF_CAN_CHANGE_ICON_SIZE Or E_SF_CAN_INVERT_SHADING Or E_SF_CAN_USE_GLASS_EFFECT

        .Add ps
    
        Set ps = New TStyle
        ps.SetTo "Info Bar", "Displays the notfication in a thin bar along the top of the screen.", _
                 S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE Or S_STYLE_V42_CONTENT, _
                 "Amiga|OS X|geOShell|Cloud:9ine|Amiga Guru", 2, 0, _
                 LIB_DATE, "info@fullphat.net", "http://getsnarl.info", Nothing

        .Add ps
    
    End With

    IStyleEngine_Initialize = M_OK

End Function

Private Function IStyleEngine_LastError() As String

    IStyleEngine_LastError = mLastErr

End Function

Private Function IStyleEngine_Name() As String

    IStyleEngine_Name = App.Title

End Function

Private Function IStyleEngine_Path() As String

    IStyleEngine_Path = App.Path

End Function

Private Function IStyleEngine_Revision() As Long

    IStyleEngine_Revision = App.Revision

End Function

Private Sub IStyleEngine_StyleAt(ByVal Index As Long, Style As libSnarlStyles.style_info)

    mStyles.GetInfo Index, Style

End Sub

Private Sub IStyleEngine_TidyUp()
End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function




'    With Style
'        ' /* generic */
'        .Copyright = App.LegalCopyright
'        .Date = LIB_DATE
'        .Path = App.Path
'        .SupportEmail = "support@fullphat.net"
'        .URL = "www.fullphat.net"
'
'        Select Case Index
'        Case 1
'            .Name = "Music Video"
'            .Description = "Displays the notification as a banner along the bottom of the screen."
'            .Major = 1
'            .Minor = 303
'            .Flags = S_STYLE_NO_DROPSHADOW Or S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE Or S_STYLE_CLICK_THROUGH Or S_STYLE_V42_CONTENT Or S_STYLE_IS_CONFIGURABLE
'            .Schemes = SCHEME_1 ' & "|" & SCHEME_2
'
'        Case 2
'            .Name = "Tiny"
'            .Description = "Displays the notfication in a tiny bubble in the top right corner of the screen."
'            .Major = 1
'            .Minor = 0
'            .Flags = S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE Or S_STYLE_IS_CONFIGURABLE Or S_STYLE_V42_CONTENT
'            .Schemes = "Top|Bottom|Top Left|Top Right|Bottom Left|Bottom Right"
'
'        Case 3
'            .Name = "Info Bar"
'            .Description = "Displays the notfication in a thin bar along the top of the screen."
'            .Major = 1
'            .Minor = 0
'            .Flags = S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE
'            .Schemes = "Amiga|OS X|geOShell|Cloud:9ine|Amiga Guru"
'
'        End Select
'
'        .IconPath = g_MakePath(App.Path) & .Name & ".png"
'
'    End With
