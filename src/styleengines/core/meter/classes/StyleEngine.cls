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

    ' /*
    '
    '   StyleEngine.cls --
    '
    '   © 2010 full phat products.  Some Rights Reserved.
    '   Licenced as per the Simplified BSD (http://www.opensource.org/licenses/bsd-license.php)
    '
    '   Notes
    '   -----
    '
    '   Standardised V42 entries accepted:
    '
    '   Generic groups
    '   --------------
    '   value-:         must be a numeric value, may be a decimal or integer
    '   label-:         may be alphanumeric
    '   colour-:        may use HTML notation (i.e. #RRGGBB) or rgb notation (i.e. r,g,b)
    '   data-:          free for custom use
    '
    '   Defined Entries
    '   ---------------
    '   value-percent:      must be a numeric value in the range 0 to 100.  May be fractional although styles may round to
    '                       nearest integer at their discretion.
    '
    '   label-subtext:      text which is typically displayed inferior to the notification text.
    '   label-tagline:      a short piece of text typically displayed in a prominent way.
    '
    '   colour-background:  requested background colour.
    '   colour-title:       requested title colour.
    '   colour-text:        requested text colour.
    '   colour-nnn:         requested colour for component "label-nnn".
    '
    '                               V42     Val-%   Label-*
    '                               ---     -----   -------
    '    STYLE_NAME_JUST_BLACK
    '    STYLE_NAME_IPHONEY          X
    '    STYLE_NAME_MINIMAL
    '    STYLE_NAME_SONY
    '    STYLE_NAME_SMOOTH
    '    STYLE_NAME_METRO
    '    STYLE_NAME_CLOCK           n/a      n/a      n/a
    '    STYLE_NAME_MUSIC
    '    STYLE_NAME_DROID            X        X        X
    '    STYLE_NAME_CORP             X        X        X
    '    STYLE_NAME_OSX              X        X       n/a
    '
    ' */

Private Const LIB_DATE = "31-Aug-2011"

Private Const STYLE_NAME_JUST_BLACK = "Just Black"
Private Const STYLE_NAME_IPHONEY = "iPhoney"
Private Const STYLE_NAME_MINIMAL = "Minimal"
Private Const STYLE_NAME_SONY = "Sony"
Private Const STYLE_NAME_SMOOTH = "Smooth"
Private Const STYLE_NAME_METRO = "Metro"
Private Const STYLE_NAME_CLOCK = "Clock"
Private Const STYLE_NAME_MUSIC = "Slideshow"
Private Const STYLE_NAME_DROID = "Droid"
Private Const STYLE_NAME_CORP = "Corporate"
Private Const STYLE_NAME_OSX = "OSX-ish"

Dim mLastErr As String

Dim mStyle() As TStyle
Dim mCount As Long

Dim mPanel As BPrefsPanel
Dim mPage As BPrefsPage
Dim mNoWriteConfig As Boolean

    ' /* local copy of style's settings that we're configuring */
Dim mSettings As T_METER_STYLE_SETTINGS
Dim mCurrent As TStyle

Implements KPrefsPage
Implements KPrefsPanel
Implements IStyleEngine

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = mCount

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance
Dim pStyle As IMeterStyle

    Select Case StyleName
    Case STYLE_NAME_JUST_BLACK
        Set pStyle = New TJustBlackInstance

    Case STYLE_NAME_IPHONEY
        Set pStyle = New TiPhoneyInstance

    Case STYLE_NAME_MINIMAL
        Set pStyle = New TMinimalInstance

    Case STYLE_NAME_SONY
        Set pStyle = New TSonyInstance

    Case STYLE_NAME_SMOOTH
        Set pStyle = New TSmoothInstance

    Case STYLE_NAME_METRO
        Set pStyle = New TMetroInstance

    Case STYLE_NAME_CLOCK
        Set pStyle = New TClockInstance

    Case STYLE_NAME_MUSIC
        Set pStyle = New TMusicInstance

    Case STYLE_NAME_DROID
        Set pStyle = New TDroidInstance

    Case STYLE_NAME_CORP
        Set pStyle = New TCorpInstance

    Case STYLE_NAME_OSX
        Set pStyle = New TOSXInstance

    End Select

    If Not (pStyle Is Nothing) Then
        pStyle.SetStyle mStyle(uIndexOf(StyleName))
        Set IStyleEngine_CreateInstance = pStyle

    End If

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = App.FileDescription

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long

    If (mPanel Is Nothing) Then

        Set mCurrent = mStyle(uIndexOf(StyleName))

        If (mCurrent Is Nothing) Then
            MsgBox "Error setting style", vbCritical, App.Title
            Exit Function

        End If

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle StyleName & " Preferences"
            .SetWidth 300
            .SetWindow 1

            If (mCurrent.Flags And E_METER_STYLE_HAS_OWN_PAGE) Then
                .AddPage new_BPrefsPage("", Nothing, mCurrent.Page)

            Else
                .AddPage new_BPrefsPage("", Nothing, Me)

            End If

            .Go
'    g_SetWindowIconToAppResourceIcon mPanel.hWnd

        End With
    End If

    IStyleEngine_GetConfigWindow = mPanel.hWnd

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    ' /* add our styles */

    uAdd STYLE_NAME_JUST_BLACK, _
         "A simple notification inspired by Notify-OSD", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE, _
         "Standard|Big Icon|Big Icon (Whole)", _
         1, 21, _
         uCreatePackedStr("title_font::Verdana#?title_size::8#?text_font::Verdana#?text_size::8#?background::" & CStr(rgba(32, 32, 32)) & "#?foreground::" & CStr(rgba(255, 255, 255)) & "#?show_graphically::1#?show_percent::0#?meter_type=1#?bar_colour::" & CStr(rgba(255, 128, 0)) & "#?spectrum_type::1"), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_VARIABLE_METERS Or E_METER_STYLE_HAS_SPECTRUMS, _
         "Monochrome Bar|Coloured Bar|Spectrum Bar|Monochrome Graph|Coloured Graph|Spectrum Graph", _
         "Red/Yellow/Green|Magenta/White/Cyan"

    uAdd STYLE_NAME_IPHONEY, _
         "Mimics the Apple iPhone alert.  Inspired by 'iPhonesque' by mattrobs", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_PULSE_NEEDED Or S_STYLE_V42_CONTENT, _
         "Custom|Black|Original|PSP", _
         1, 1, _
         uCreatePackedStr("title_font::Arial#?title_size::10#?text_font::Arial#?text_size::9#?fg_alpha::70#?background::" & CStr(rgba(56, 89, 131)) & "#?foreground::" & CStr(rgba(255, 255, 255)) & "#?show_graphically::1#?show_icon::1#?icon_align::1#?show_percent::0"), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_CAN_CHANGE_BKGND Or E_METER_STYLE_CAN_SET_FG_ALPHA

'            .TitleFont = uGetDefaultFont(STYLE_NAME_IPHONEY, False)
'            .TextFont = uGetDefaultFont(STYLE_NAME_IPHONEY, True)
'            .BackgroundColour = rgba(56, 89, 131)
'            .TextColour = rgba(255, 255, 255)
'            .TextAlpha = 70

'             "Standard|Standard (Icon Only)|Standard (Title Only)|Standard (Bar)|" & _
'             "Black|Black (Icon Only)|Black (Title Only)|Black (Bar)|" & _
'             "Original|Original (Icon Only)|Original (Title Only)||Original (Bar)|" & _
'            "Sony PSP|Sony PSP (Icon Only)|Sony PSP (Title Only)|Sony PSP (Bar)", _


    uAdd STYLE_NAME_MINIMAL, _
         "A simplistic minimal look", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE, _
         "Minimalistique|Koka|Lyme|Charkoal", _
         1, 0, _
         uCreatePackedStr("title_font::Arial#?title_size::10#?text_font::Arial#?text_size::9#?fg_alpha::100#?show_graphically::1#?show_icon::1#?icon_align::1#?show_percent::0"), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_CAN_SET_ICON_POS Or E_METER_STYLE_CAN_HIDE_SHOW_ICON

'#?background::" & CStr(rgba(130, 138, 196)) & "#?foreground::" & CStr(rgba(255, 255, 255)) & "

    uAdd STYLE_NAME_SONY, _
         "Displays a notification with a background that changes colour based on the month of the year", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_V42_CONTENT, _
         "PS3|PS3 (Big Icon)|PSP|PSP (Big Icon)", _
         2, 0, uCreatePackedStr("title_font::Trebuchet MS#?title_size::10#?text_font::Trebuchet MS#?text_size::9#?fg_alpha::100#?show_graphically::1#?show_icon::1#?icon_align::0#?show_percent::1"), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_NO_LEGACY_METER

    uAdd STYLE_NAME_SMOOTH, _
         "A simple smooth style", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_V42_CONTENT, _
         "Smooth", _
         2, 0, _
         uCreatePackedStr("title_font::Arial#?title_size::10#?text_font::Arial#?text_size::9#?background::" & CStr(rgba(51, 102, 196)) & "#?foreground::" & CStr(rgba(255, 255, 255)) & "#?border::" & CStr(rgba(0, 0, 0))), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_CAN_CHANGE_BKGND Or E_METER_STYLE_LOAD_SCHEMES_FROM_FILE Or E_METER_STYLE_HAS_CUSTOM_SCHEME Or E_METER_STYLE_NO_LEGACY_METER Or E_METER_STYLE_CAN_CHANGE_FGND Or E_METER_STYLE_CAN_CHANGE_BORDER

    uAdd STYLE_NAME_METRO, _
         "A modern style", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE, _
         "Custom", _
         1, 0, _
         uCreatePackedStr("title_font::Arial#?title_size::11#?text_font::Arial#?text_size::9#?background::" & CStr(rgba(255, 255, 255)) & "#?border::" & CStr(rgba(255, 176, 0)) & "#?foreground::" & CStr(rgba(82, 98, 160)) & "#?show_graphically::1#?show_percent::1"), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_CAN_CHANGE_BKGND Or E_METER_STYLE_CAN_CHANGE_BORDER Or E_METER_STYLE_CAN_CHANGE_FGND Or E_METER_STYLE_HAS_CUSTOM_SCHEME Or E_METER_STYLE_LOAD_SCHEMES_FROM_FILE


    uAdd STYLE_NAME_CLOCK, _
         "Displays notification text formatted as 'YYYYMMDDHHMMSS' as a time", _
         S_STYLE_MULTI_INSTANCE Or S_STYLE_V42_CONTENT, _
         "", _
         1, 0, uCreatePackedStr(""), E_METER_STYLE_LOAD_SCHEMES_FROM_FILE Or E_METER_STYLE_HAS_OWN_PAGE, , , New TClockPage


    uAdd STYLE_NAME_MUSIC, _
         "Displays a fixed-size rectangular notification with the text part is limited to two lines.  Landscape icons with a height of more than 100px are recommended.", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE, _
         "Top and Bottom|Bottom|Top|None", _
         1, 0, _
         uCreatePackedStr("title_font::Arial#?title_size::8#?text_font::Arial#?text_size::7#?background::" & CStr(rgba(17, 17, 17)) & "#?show_graphically::1#?fg_alpha::100"), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_CAN_CHANGE_BKGND Or E_METER_STYLE_CAN_CHANGE_FGND Or E_METER_STYLE_CAN_SET_FG_ALPHA Or E_METER_STYLE_NO_SHOW_PERCENT


    uAdd STYLE_NAME_DROID, _
         "Droid Style", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_V42_CONTENT Or S_STYLE_PULSE_NEEDED, _
         "Standard", _
         1, 0, _
         uCreatePackedStr("title_font::Droid Sans#?title_size::13#?text_font::Droid Sans#?text_size::9#?background::" & CStr(rgba(87, 87, 87)) & "#?fg_alpha::100"), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_NO_SHOW_PERCENT


    uAdd STYLE_NAME_CORP, _
         "Corporate Style", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_V42_CONTENT Or S_STYLE_PULSE_NEEDED, _
         "Standard|Light|Standard framed|Light framed", _
         1, 1, _
         uCreatePackedStr("title_font::Droid Sans#?title_size::11#?text_font::Droid Sans#?text_size::9#?background::" & CStr(rgba(15, 15, 15)) & "#?foreground::" & CStr(rgba(255, 255, 255)) & "#?fg_alpha::100"), _
         E_METER_STYLE_CAN_CHANGE_FONT Or E_METER_STYLE_NO_SHOW_PERCENT


    uAdd STYLE_NAME_OSX, _
         "A single-instance style inspired by the OS X On Screen Display", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_V42_CONTENT Or S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_DONT_QUEUE Or S_STYLE_NO_DROPSHADOW Or S_STYLE_CLICK_THROUGH, _
         "Standard|Solid", _
         2, 0, _
         uCreatePackedStr("background::" & CStr(rgba(24, 24, 24)) & "#?bg_alpha::67" & "#?foreground::" & CStr(rgba(255, 255, 255))), _
         E_METER_STYLE_CAN_CHANGE_BKGND Or E_METER_STYLE_CAN_SET_BG_ALPHA Or E_METER_STYLE_NO_SHOW_PERCENT Or E_METER_STYLE_SINGLE_PREVIEW Or E_METER_STYLE_CAN_CHANGE_FGND Or E_METER_STYLE_NO_LEGACY_METER


Dim pPacked As BPackedData

    ' /* defaults */

    gDbgMinute = -1

    Set pPacked = New BPackedData
    With pPacked
        .Add "name", "Tahoma"
        .Add "size", "9"

    End With

'    With gSettings
'        .ColourGraphColour = rgba(255, 128, 0)
''        .Font = pPacked.AsString()
'        .ShowBasRelief = True
'        .ShowDarkShade = False
'        .ShowTitle = True
'        .ShowText = True
'        .SpectrumType = 1
'
'        .PriorityBackgroundColour = rgba(232, 92, 92)
'        .PriorityEmblem = ""
'        .CentreIcon = True
'
'        ' /* Just Black style */
'
'        With .JustBlack
'            .TitleFont = uGetDefaultFont(STYLE_NAME_JUST_BLACK, False)
'            .TextFont = uGetDefaultFont(STYLE_NAME_JUST_BLACK, True)
''            .BackgroundColour = rgba(56, 89, 131)
'            .TextColour = rgba(255, 255, 255)
''            .TextAlpha = 70
'
'        End With
'
'        ' /* iPhoney style  */
'
'        With .iPhoney

'
'        End With
'
'        ' /* Sony style  */
'
'        With .Sony
'            .TitleFont = uGetDefaultFont(STYLE_NAME_SONY, False)
'            .TextFont = uGetDefaultFont(STYLE_NAME_SONY, True)
'            .TextAlpha = 50     ' // actually the dropshadow strength (%)
'
'        End With
'
'        ' /* Minimal style */
'
''        With .Minimal
''            .TitleFont = uGetDefaultFont(STYLE_NAME_MINIMAL, False)
''            .TextFont = uGetDefaultFont(STYLE_NAME_MINIMAL, True)
''            .BackgroundColour = rgba(255, 255, 255)
''            .TextColour = rgba(168, 192, 225)
''            .ShowMeterPercent = True
''
''        End With
'
'    End With

    ' /* load config here */

Dim sz As String
Dim i As Long

    With New ConfigFile
        .File = style_GetSnarlConfigPath("meter")
        If .Load Then
            i = .FindSection("general")
            If i Then
'                With .SectionAt(i)
'                    If .Find("ColourGraphColour", sz) Then _
'                        gSettings.ColourGraphColour = Val(sz)
'
''                    If .Find("Font", sz) Then _
'                        gSettings.Font = sz
'
'                    If .Find("ShowBasRelief", sz) Then _
'                        gSettings.ShowBasRelief = (Val(sz) <> 0)
'
'                    If .Find("ShowDarkShade", sz) Then _
'                        gSettings.ShowDarkShade = (Val(sz) <> 0)
'
'                    If .Find("ShowTitle", sz) Then _
'                        gSettings.ShowTitle = (Val(sz) <> 0)
'
'                    If .Find("ShowText", sz) Then _
'                        gSettings.ShowText = (Val(sz) <> 0)
'
'                    If .Find("SpectrumType", sz) Then _
'                        gSettings.SpectrumType = Val(sz)
'
'                    If .Find("PriorityBackgroundColour", sz) Then _
'                        gSettings.PriorityBackgroundColour = Val(sz)
'
'                    If .Find("PriorityEmblem", sz) Then _
'                        gSettings.PriorityEmblem = sz
'
'                    If .Find("CentreIcon", sz) Then _
'                        gSettings.CentreIcon = (Val(sz) <> 0)
'
'                End With
            End If
        End If

    End With

    ' /* load the priority emblem image */

    g_UpdateEmblem

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

    If (Index < 1) Or (Index > mCount) Then _
        Exit Sub

    mStyle(Index).GetInfo Style

End Sub

Private Sub IStyleEngine_TidyUp()

End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function

Private Sub KPrefsPage_AllAttached()
Dim pc As BControl

    If mPanel.Find("fs>dbgTime", pc) Then _
        pc.SetValue "-1"

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg
Dim dw As Long
Dim sz As String

'    If mPanelStyle = STYLE_NAME_JUST_BLACK Then
'        ' /* Just Black */
'
'        With mPage
'            .Add new_BPrefsControl("banner", "", "General Settings")
''            .Add new_BPrefsControl("fancyfontpicker", "fnp>Font", "", "Font:", gSettings.Font)
'            .Add new_BPrefsControl("fancytoggle2", "ft>ShowTitle", "Display title in the notification?", , IIf(gSettings.ShowTitle, "1", "0"))
'            .Add new_BPrefsControl("fancytoggle2", "ft>ShowText", "Display text in the notification?", , IIf(gSettings.ShowText, "1", "0"))
'            .Add new_BPrefsControl("fancytoggle2", "ft>CentreIcon", "Centre icon vertically within the notification?", , IIf(gSettings.CentreIcon, "1", "0"))
'            .Add new_BPrefsControl("banner", "", "Bar and Graph Settings")
'            .Add new_BPrefsControl("fancytoggle2", "ft>ShowBasRelief", "Draw bas-relief effect?", , IIf(gSettings.ShowBasRelief, "1", "0"))
'            .Add new_BPrefsControl("fancytoggle2", "ft>ShowDarkShade", "Draw shadow effect?", , IIf(gSettings.ShowDarkShade, "1", "0"))
'            .Add new_BPrefsControl("fancycolourpicker", "fcp>ColourGraphColour", "", "Colour:", gSettings.ColourGraphColour)
'
'            dw = gSettings.ColourGraphColour
'            .Add new_BPrefsControl("label", "lbl>ColourGraphColour", get_red(dw) & "," & get_green(dw) & "," & get_blue(dw) & " (#" & g_HexStr(get_red(dw), 2) & g_HexStr(get_green(dw), 2) & g_HexStr(get_blue(dw), 2) & ")", "")
'
'            Set pm = New CTempMsg
'            pm.Add "item-height", 24&
'            Set pc = new_BPrefsControl("combo", "c>spectrum", "Red/Yellow/Green|Magenta/White/Cyan", "Spectrum:", CStr(gSettings.SpectrumType + 1), pm)
'            .Add pc
'            uSetSpectrumIcon pc, 1, 0
'            uSetSpectrumIcon pc, 2, 1
'
'            .Add new_BPrefsControl("banner", "", "Priority Notfications")
'            .Add new_BPrefsControl("fancycolourpicker", "fcp>PriorityBackgroundColour", "", "Colour:", CStr(gSettings.PriorityBackgroundColour))
'
'            Set pm = New CTempMsg
'            pm.Add "filter", "PNG Images#?png"
'
'            sz = g_GetPath(gSettings.PriorityEmblem)
'            If sz = "" Then _
'                sz = g_GetSystemFolderStr(CSIDL_MYPICTURES)
'
'            pm.Add "path", sz
'
'            .Add new_BPrefsControl("fancyfilepicker", "ffp>PriorityEmblem", gSettings.PriorityEmblem, "Emblem:", , pm)
'
'        End With
'
'        Exit Sub
'
'    End If

Dim fUseMeter As Boolean

    ' /* construct based on the style we're configuring */

    With mPage

'        .Add new_BPrefsControl("banner", "", "General Settings")

        ' /* font pickers */

        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_FONT) Then
            .Add new_BPrefsControl("fancyfontpicker", "title", "", "Title font:", uCreatePackedFontStr(mCurrent.ReadSetting("title_font"), mCurrent.ReadSetting("title_size")))
            .Add new_BPrefsControl("fancyfontpicker", "text", "", "Text font:", uCreatePackedFontStr(mCurrent.ReadSetting("text_font"), mCurrent.ReadSetting("text_size")))

        End If

        ' /* icon */
        If (mCurrent.Flags And E_METER_STYLE_CAN_HIDE_SHOW_ICON) Then
            .Add new_BPrefsControl("fancytoggle2", "show_icon", "Show icon?", "Icon:", mCurrent.ReadSetting("show_icon"))
            If (mCurrent.Flags And E_METER_STYLE_CAN_SET_ICON_POS) Then
                Set pm = New CTempMsg
                pm.Add "text", "Position:"
                pm.Add "align", 1&
                .Add new_BPrefsControl("fancytoggle", "icon_align", "Left|Right", " ", mCurrent.ReadSetting("icon_align"), pm)

            End If
        End If

        ' /* background colour picker */

Dim f As Boolean

        f = ((mCurrent.Name = STYLE_NAME_JUST_BLACK) And (g_IsPressed(VK_LCONTROL)))
        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_BKGND) Or (f) Then _
            .Add new_BPrefsControl("fancycolourpicker", "background", "", IIf(f, "Not just black:", "Background:"), mCurrent.ReadSetting("background"))

        ' /* foreground opacity */
        If (mCurrent.Flags And E_METER_STYLE_CAN_SET_BG_ALPHA) Then
            Set pm = New CTempMsg
            pm.Add "min", 0&
            pm.Add "max", 100&
            pm.Add "freq", 10&
            pm.Add "suffix", "%"
            .Add new_BPrefsControl("fancyslider", "bg_alpha", "", "Background opacity:", mCurrent.ReadSetting("bg_alpha"), pm)

        End If

        ' /* foreground colour picker */
        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_FGND) Then _
            .Add new_BPrefsControl("fancycolourpicker", "foreground", "", "Text:", mCurrent.ReadSetting("foreground"))

        ' /* border colour picker */
        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_BORDER) Then _
            .Add new_BPrefsControl("fancycolourpicker", "border", "", "Border:", mCurrent.ReadSetting("border"))

        ' /* foreground opacity */
        If (mCurrent.Flags And E_METER_STYLE_CAN_SET_FG_ALPHA) Then
            Set pm = New CTempMsg
            pm.Add "min", 0&
            pm.Add "max", 100&
            pm.Add "freq", 10&
            pm.Add "suffix", "%"
            .Add new_BPrefsControl("fancyslider", "fg_alpha", "", "Foreground opacity:", mCurrent.ReadSetting("fg_alpha"), pm)

        End If

        ' /* meter */
        If (mCurrent.Flags And E_METER_STYLE_NO_LEGACY_METER) = 0 Then
            .Add new_BPrefsControl("banner", "", "Meter")
            .Add new_BPrefsControl("fancytoggle2", "show_graphically", "Enable legacy meter support?", "", mCurrent.ReadSetting("show_graphically"))
            fUseMeter = (Val(mCurrent.ReadSetting("show_graphically")) <> 0)
    
            If (mCurrent.Flags And E_METER_STYLE_VARIABLE_METERS) Then
                .Add new_BPrefsControl("fancycycle", "meter_type", mCurrent.MeterTypes, "Type:", mCurrent.ReadSetting("meter_type"), , fUseMeter)
                If (mCurrent.Flags And E_METER_STYLE_HAS_SPECTRUMS) Then
                    .Add new_BPrefsControl("fancycolourpicker", "bar_colour", "", "Bar colour:", mCurrent.ReadSetting("bar_colour"))
                    Set pm = New CTempMsg
                    pm.Add "item-height", 24&
                    Set pc = new_BPrefsControl("combo", "spectrum_type", mCurrent.Spectrums, "Spectrum type:", mCurrent.ReadSetting("spectrum_type"), pm, (InStr(mCurrent.MeterType(Val(mCurrent.ReadSetting("meter_type"))), "Spectrum")))
                    .Add pc
                    uSetSpectrumIcon pc, 1, 0
                    uSetSpectrumIcon pc, 2, 1
    
                End If
            End If

            If (mCurrent.Flags And E_METER_STYLE_NO_SHOW_PERCENT) = 0 Then _
                .Add new_BPrefsControl("fancytoggle2", "show_percent", "Display value as percentage?", "Options:", mCurrent.ReadSetting("show_percent"))

        End If

        ' /* test and reset buttons - all */
        .Add new_BPrefsControl("separator", "", "")
'        .Add new_BPrefsControl("fancybutton2", "fb2>Test", "Test")
'        .Add new_BPrefsControl("fancybutton2", "fb2>Reset", "Reset to Factory Settings")
        .Add new_BPrefsControl("fancytoolbar", "main_toolbar", "Test Settings||Restore Defaults")

        ' /* debugging section - Sony only ATM (and only if CTRL held down) */
        If g_IsPressed(VK_LCONTROL) Then
'            If mPanelStyle = STYLE_NAME_SONY Then
'            .Add new_BPrefsControl("separator", "", "")
                .Add new_BPrefsControl("banner", "", "Debugging")
                Set pm = New CTempMsg
                pm.Add "min", -1&
                pm.Add "max", 1439&
                pm.Add "freq", 120&
                pm.Add "label_size", 50&
                .Add new_BPrefsControl("fancyslider", "fs>dbgTime", "", "Time of day:", "0", pm)

'            End If
        End If

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Debug.Print Control.GetName() & " > " & g_HexStr(Val(Value))

    ' /* write settings back to config file */

    Select Case Control.GetName()
    Case "main_toolbar"
        Select Case Val(Value)
        Case 1
            ' /* test settings */
            mCurrent.DoPreview

        Case 3
            ' /* reset to factory defaults */
            uResetToDefaults

        End Select

    Case "title", "text"
        ' /* fonts need a little extra tweaking... */
        With New BPackedData
            If .SetTo(Value) Then
                mCurrent.WriteSetting Control.GetName & "_font", .ValueOf("name")
                mCurrent.WriteSetting Control.GetName & "_size", .ValueOf("size")

            End If

        End With

    Case Else
        ' /* ...everything else is straightforward */
        mCurrent.WriteSetting Control.GetName, Value

    End Select

    ' /* deal with control setting */

Dim pc As BControl

    Select Case Control.GetName()

    Case "meter_type"
        If Control.Page.Find("spectrum_type", pc) Then _
            pc.SetEnabled (InStr(mCurrent.MeterType(Val(Value)), "Spectrum"))

        If Control.Page.Find("bar_colour", pc) Then _
            pc.SetEnabled (InStr(mCurrent.MeterType(Val(Value)), "Coloured"))

'    ' /* Sony only */
'
'    Case "fs>dbgTime"
'        ' /* debug mode minute slider */
'        gDbgMinute = Val(Value)
'        If Val(Value) = -1 Then
'            Control.SetText "Auto"
'
'        Else
'            Control.SetText Format$(TimeSerial(0, Val(Value), 0), "h:mm AMPM")
'
'        End If
'

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName()
    Case "fb2>Test"
        ' /* test settings */
        mCurrent.DoPreview

    Case "fb2>Reset"
        ' /* reset to factory defaults */
        uResetToDefaults

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 140

    ' /* get our settings */

'    Select Case mPanelStyle
''    Case STYLE_NAME_MINIMAL
''        LSet mSettings = gSettings.Minimal
'
'    Case STYLE_NAME_IPHONEY
'        LSet mSettings = gSettings.iPhoney
'
'    Case STYLE_NAME_SONY
'        LSet mSettings = gSettings.Sony
'
'    End Select

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
    Set mCurrent = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub uSetSpectrumIcon(ByRef Ctl As BControl, ByVal Index As Long, ByVal Spectrum As Long)
Dim pv As mfxView
Dim pm As CTempMsg

    Set pv = New mfxView
    With pv
        .SizeTo 24, 24
        .EnableSmoothing False
        If Spectrum = 0 Then
            .SetHighColour rgba(255, 0, 0)
            .SetLowColour rgba(255, 255, 0)

        ElseIf Spectrum = 1 Then
            .SetHighColour rgba(255, 0, 255)
            .SetLowColour rgba(255, 255, 255)

        End If

        .FillRect new_BRect(0, 0, 12, 24), MFX_HORIZ_GRADIENT

        If Spectrum = 0 Then
            .SetHighColour rgba(255, 255, 0)
            .SetLowColour rgba(0, 255, 0)

        ElseIf Spectrum = 1 Then
            .SetHighColour rgba(255, 255, 255)
            .SetLowColour rgba(0, 255, 255)

        End If

        .FillRect new_BRect(12, 0, 24, 24), MFX_HORIZ_GRADIENT

        .SetHighColour rgba(0, 0, 0, 102)
        .StrokeRect .Bounds

    End With

    Set pm = New CTempMsg
    pm.Add "index", Index
    pm.Add "image-object", pv.ConvertToBitmap()
    Ctl.DoExCmd B_SET_ITEM, pm

End Sub

'Private Sub uWriteConfig()
'
'    If mNoWriteConfig Then _
'        Exit Sub
'
'    ' /* store our settings */
'
''    With New ConfigFile
''        .File = style_GetSnarlConfigPath("meter")
''        With .AddSectionObj("general")
''            .Add "ColourGraphColour", CStr(gSettings.ColourGraphColour)
'''            .Add "Font", gSettings.Font
''            .Add "ShowBasRelief", IIf(gSettings.ShowBasRelief, "1", "0")
''            .Add "ShowDarkShade", IIf(gSettings.ShowDarkShade, "1", "0")
''            .Add "ShowTitle", IIf(gSettings.ShowTitle, "1", "0")
''            .Add "ShowText", IIf(gSettings.ShowText, "1", "0")
''            .Add "SpectrumType", CStr(gSettings.SpectrumType)
''            .Add "PriorityBackgroundColour", CStr(gSettings.PriorityBackgroundColour)
''            .Add "PriorityEmblem", gSettings.PriorityEmblem
''            .Add "CentreIcon", IIf(gSettings.CentreIcon, "1", "0")
''
''        End With
''
''        uWriteConfigSection .AddSectionObj(STYLE_NAME_JUST_BLACK), gSettings.Sony
''        uWriteConfigSection .AddSectionObj(STYLE_NAME_IPHONEY), gSettings.iPhoney
'''        uWriteConfigSection .AddSectionObj(STYLE_NAME_MINIMAL), gSettings.Minimal
''        uWriteConfigSection .AddSectionObj(STYLE_NAME_SONY), gSettings.Sony
''
''        .Save
''
''    End With
'
'End Sub

'Private Function uGetDefaultFont(ByVal Style As String, ByVal Text As Boolean) As String
'Dim szFace As String
'Dim dwSize As Long
'
'
'    Select Case Style
'    Case STYLE_NAME_JUST_BLACK
'        szFace = "Bitstream Vera Sans"
'        dwSize = 9
'        If Not g_FontExists(szFace) Then _
'            szFace = "Tahoma"
'
'
'
'    Case Else
'        g_Debug "StyleEngine.uGetDefaultFont(): '" & Style & "' is unknown", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End Select
'
'    With New BPackedData
'        .Add "name", szFace
'        .Add "size", CStr(dwSize)
'        uGetDefaultFont = .AsString()
'
'    End With
'
'End Function

Private Sub uAdd(ByVal StyleName As String, ByVal Description As String, ByVal Flags As S_STYLE_FLAGS, ByVal Schemes As String, ByVal Major As Long, ByVal Minor As Long, ByRef Defaults As BPackedData, ByVal IntFlags As E_METER_STYLE_FLAGS, Optional ByVal MeterTypes As String, Optional ByVal SpectrumTypes As String, Optional ByRef CustomPrefsPage As KPrefsPage)
Dim ps As TStyle

    Set ps = New TStyle
    ps.Init StyleName, Description, Flags, Schemes, Major, Minor, Defaults, LIB_DATE, IntFlags, MeterTypes, SpectrumTypes, CustomPrefsPage

    mCount = mCount + 1
    ReDim Preserve mStyle(mCount)
    Set mStyle(mCount) = ps

End Sub

Private Function uIndexOf(ByVal Name As String) As Long
Static i As Long

    For i = 1 To mCount
        If mStyle(i).Name = Name Then
            uIndexOf = i
            Exit Function

        End If
    Next i

End Function

Private Function uCreatePackedStr(ByVal Content As String) As BPackedData

    Set uCreatePackedStr = New BPackedData
    uCreatePackedStr.SetTo Content

End Function

Private Function uCreatePackedFontStr(ByVal Typeface As String, ByVal TypeSize As String) As String
Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.Add "name", Typeface
    ppd.Add "size", TypeSize
    uCreatePackedFontStr = ppd.AsString()

End Function

Private Sub uResetToDefaults()
Dim pc As BControl

    mNoWriteConfig = True

    ' /* set our title and text controls - the correct values are determined by
    '    uGetDefaultFont() and are stored in the right place as setting a
    '    control's value manually still generates a 'changed' notification */

Dim szn As String
Dim szv As String

    With mCurrent.Defaults
        .Rewind
        Do While .GetNextItem(szn, szv)
            If mPage.Find(szn, pc) Then _
                pc.SetValue szv

        Loop

    End With

    ' /* fonts */
    If mPage.Find("title", pc) Then _
        pc.SetValue uCreatePackedFontStr(mCurrent.GetDefault("title_font"), mCurrent.GetDefault("title_size"))

    If mPage.Find("text", pc) Then _
        pc.SetValue uCreatePackedFontStr(mCurrent.GetDefault("text_font"), mCurrent.GetDefault("text_size"))

'            Select Case mPanelStyle
'            Case STYLE_NAME_IPHONEY
'                pc.SetValue CStr(rgba(56, 89, 131))
'
'            Case STYLE_NAME_MINIMAL
'                pc.SetValue CStr(rgba(255, 255, 255))
'
'            End Select
'        End If
'
'        If mPage.Find("foreground", pc) Then
'            Select Case mPanelStyle
'            Case STYLE_NAME_IPHONEY
'                pc.SetValue CStr(rgba(255, 255, 255))
'
'            Case STYLE_NAME_MINIMAL
'                pc.SetValue CStr(rgba(168, 192, 225))
'
'            End Select
'        End If

    ' /* set style-specific controls here */

'        Select Case mPanelStyle
'        Case STYLE_NAME_SONY
'            ' /* Sony */
'
'            If mPage.Find("fs>DropshadowStrength", pc) Then _
'                pc.SetValue "50"

    mNoWriteConfig = False

End Sub
