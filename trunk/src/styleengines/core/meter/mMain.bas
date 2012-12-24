Attribute VB_Name = "mMain"
Option Explicit

Public Type T_LABEL
    Text As String
    Frame As BRect

End Type

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Dim mSettings As ConfigFile

Public Type T_METER_STYLE_SETTINGS
    TitleFont As String
    TextFont As String

    BackgroundColour As Long        ' // applies to 'standard' scheme only
'    TextColour As Long
'    TextAlpha As Long
'    ShowMeterPercent As Boolean

    ' /* not persistent */
    zzTitleFont As String
    zzTitleFontSize As Long
    zzTextFont As String
    zzTextFontSize As Long

End Type

Public Type T_SETTINGS

    ShowTitle As Boolean            ' // default=true
    ShowText As Boolean             ' // default=true
    ShowBasRelief As Boolean        ' // default=true
    ShowDarkShade As Boolean        ' // default=true

    ColourGraphColour As Long       ' // colour to use in "Colour Graph" scheme
    SpectrumType As Integer         ' // 0 (default) = red/yellow/green; 1 = magenta/while/cyan

    PriorityBackgroundColour As Long
    PriorityEmblem As String        ' // path to PNG

    CentreIcon As Boolean           ' // default=true

    JustBlack As T_METER_STYLE_SETTINGS
    iPhoney As T_METER_STYLE_SETTINGS
    Sony As T_METER_STYLE_SETTINGS
'    Minimal As T_METER_STYLE_SETTINGS

End Type

Public gSettings As T_SETTINGS
Public gPriorityEmblem As mfxBitmap
Public gEmblemSize As Long
Public gDbgMinute As Long

Public Sub g_UpdateEmblem()
Dim pbm As mfxBitmap

'    Set gPriorityEmblem = Nothing
'    If load_image(gSettings.PriorityEmblem, pbm) Then
'        If (pbm.Width > 0) And (pbm.Height > 0) Then
'            gEmblemSize = pbm.Height
'            Set gPriorityEmblem = pbm
'
'        End If
'    End If

End Sub

Public Function g_GetPSPColour() As Long
Dim dw As Long

    Select Case Month(Now())
    Case 1:     dw = rgba(192, 192, 192)
    Case 2:     dw = rgba(240, 230, 140)
    Case 3:     dw = rgba(173, 255, 46)
    Case 4:     dw = rgba(255, 105, 180)
    Case 5:     dw = rgba(51, 205, 51)
    Case 6:     dw = rgba(123, 103, 238)
    Case 7:     dw = rgba(0, 206, 209)
    Case 8:     dw = rgba(0, 0, 205)
    Case 9:     dw = rgba(138, 43, 226)
    Case 10:    dw = rgba(255, 215, 0)
    Case 11:    dw = rgba(139, 68, 21)
    Case 12:    dw = rgba(255, 0, 0)

    End Select

    g_GetPSPColour = dw

End Function

Public Function g_GetPS3Colour() As Long
Dim fInMonth As Boolean
Dim dw As Long

    fInMonth = ((Day(Now) >= 15) And (Day(Now) <= 24))

    If fInMonth Then
        ' /* between 15th and 24th of the month */
        Select Case Month(Now())
        Case 1
            dw = rgba(240, 230, 140)

        Case 2
            dw = rgba(107, 142, 34)

        Case 3
            dw = rgba(255, 192, 203)

        Case 4
            dw = rgba(192, 192, 192)

        Case 5
            dw = rgba(216, 191, 216)

        Case 6
            dw = rgba(173, 216, 230)

        Case 7
            dw = rgba(0, 0, 255)

        Case 8
            dw = rgba(128, 0, 128)

        Case 9
            dw = rgba(128, 0, 0)

        Case 10
            dw = rgba(184, 115, 51)

        Case 11
            dw = rgba(206, 34, 43)

        Case 12
            dw = rgba(224, 176, 255)

        End Select

    Else
        ' /* before 15th of month or after 24th of month */

Dim i As Integer

        i = Month(Now)

        If Day(Now) > 24 Then
            dw = uGetPS3OffsetColour(i)

        Else
            i = i - 1
            If i = 0 Then _
                i = 12

            dw = uGetPS3OffsetColour(i)
        
        End If

    End If

    g_GetPS3Colour = dw

End Function

Private Function uGetPS3OffsetColour(ByVal Month As Integer)
Dim dw As Long

    ' /* return colour for after 24th of specified month */

    Select Case Month
    Case 1:     dw = rgba(165, 43, 43)
    Case 2:     dw = rgba(0, 128, 0)
    Case 3:     dw = rgba(255, 105, 180)
    Case 4:     dw = rgba(0, 128, 0)
    Case 5:     dw = rgba(128, 0, 128)
    Case 6:     dw = rgba(0, 128, 128)
    Case 7:     dw = rgba(0, 0, 139)
    Case 8:     dw = rgba(238, 130, 238)
    Case 9:     dw = rgba(240, 230, 140)
    Case 10:    dw = rgba(165, 43, 43)
    Case 11:    dw = rgba(255, 0, 0)
    Case 12:    dw = rgba(192, 192, 192)

    End Select

    uGetPS3OffsetColour = dw

End Function

Public Function g_FontExists(ByVal Typeface As String) As Boolean
Dim hr As Long

    hr = CreateFont(-13, 0, 0, 0, 400, 0, 0, 0, 0&, 1, 0, 0, 0, Typeface)
    If hr Then
        DeleteObject hr
        g_FontExists = True

    End If

End Function

Public Function g_GetSettings() As Boolean

    Set mSettings = New ConfigFile

    With mSettings
        .File = style_GetSnarlConfigPath("meter")
        .Load

'        if .FindSection()


    End With
        
'        If .Load Then
'            i = .FindSection("general")
'            If i Then
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
'            End If
'
'            i = .FindSection(STYLE_NAME_1)
'            If i Then _
'                uGetSettings .SectionAt(i), gSettings.JustBlack
'
'            i = .FindSection(STYLE_NAME_IPHONEY)
'            If i Then _
'                uGetSettings .SectionAt(i), gSettings.iPhoney
'
'            i = .FindSection(STYLE_NAME_SONY)
'            If i Then _
'                uGetSettings .SectionAt(i), gSettings.Sony
'
'            i = .FindSection(STYLE_NAME_MINIMAL)
'            If i Then _
'                uGetSettings .SectionAt(i), gSettings.Minimal
'
'            ' /* validate loaded settings */
'
'
'            With gSettings.iPhoney
'                If .TextAlpha < 0 Then
'                    .TextAlpha = 0
'
'                ElseIf .TextAlpha > 255 Then
'                    .TextAlpha = 255
'
'                End If
'
'            End With
'
'            With gSettings.Sony
'                If .TextAlpha < 0 Then
'                    .TextAlpha = 0
'
'                ElseIf .TextAlpha > 100 Then
'                    .TextAlpha = 100
'
'                End If
'            End With
'        End If


End Function

'Public Function g_GetSetting(ByRef Settings As ConfigSection, ByVal Name As String, ByRef Defaults As BPackedData) As String
'
'    If (Settings Is Nothing) Or (Defaults Is Nothing) Then _
'        Exit Function
'
'Dim sz As String
'
'    If Settings.Find(Name, sz) Then
'        g_GetSetting = sz
'
'    Else
'        g_GetSetting = Defaults.ValueOf(Name)
'
'    End If
'
'End Function



Public Function g_CreateMarker(ByVal Colour As Long, Optional ByVal Height As Long = 24) As MImage
Dim pp(3) As BPoint

    With New mfxView
        .SizeTo 20, Height + 2
        .EnableSmoothing True
        .SetHighColour Colour

        Set pp(0) = new_BPoint(0, 0)
        Set pp(1) = new_BPoint(0, Height - 1)
        Set pp(2) = new_BPoint(0 + 9, Height - 5 - 1)
        Set pp(3) = new_BPoint(0 + 9, 0)
        .FillShape pp(), True

        Set pp(0) = new_BPoint(0 + 9, 0)
        Set pp(1) = new_BPoint(0 + 9, Height - 5 - 1)
        Set pp(2) = new_BPoint(0 + 9 + 9, Height - 1)
        Set pp(3) = new_BPoint(0 + 9 + 9, 0)
        .FillShape pp(), True

        .EnableSmoothing False
        .StrokeLine new_BRect(0 + 9, 0, 0 + 9, Height - 5 - 1)

        Set g_CreateMarker = .ConvertToBitmap()

    End With

End Function
