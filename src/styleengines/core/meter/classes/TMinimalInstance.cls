VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMinimalInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TMinimalInstance.cls
'/
'/  Description:    Represents a single instance of a Minimal style notification
'/
'/  © 2010 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Const EM_TITLE_TEXT = 0
Const EM_MARGIN = 10

Dim mInfo As notification_info
Dim mIconSize As Long
Dim mIcon As mfxBitmap
Dim mView As mfxView

Dim mDrawBar As Boolean
Dim mIconOnly As Boolean
Dim mTitleRect As BRect
Dim mTextRect As BRect

Dim mStyle As TStyle

Implements IStyleInstance
Implements IMeterStyle


Private Sub Class_Initialize()

    Set mView = New mfxView

    With mView
        .TextMode = MFX_TEXT_ANTIALIAS

    End With

End Sub

Private Sub IMeterStyle_SetStyle(Style As TStyle)

    Set mStyle = Style

End Sub

'Private Sub uDrawBar()
'Dim xOffset As Long
'Dim pr As brect
'Dim prTray As brect
'Dim cy As Long
'Dim n As Long
'
'    uDrawBackground cy, xOffset
'
'    With mView
'
'        ' /* meter tray */
'
'        Set pr = new_brect(0, 0, .Width - cy, 6)
'        pr.InsetBy 10, 0
'        pr.OffsetBy xOffset, Fix((.Height - pr.Height) / 2)
'
'        Set prTray = pr.Duplicate
'
'        .EnableSmoothing True
'        pr.InsetBy -2, -2
'        .SetHighColour rgba(28, 28, 28)
'        .FillRoundRect pr, 2, 2
'        .EnableSmoothing False
'        pr.InsetBy 2, 2
'
'        .SetHighColour rgba(100, 100, 100)
'        .FillRect pr, MFX_SOLID_HIGH            ' // tray itself
'
'        .SetHighColour rgba(0, 0, 0, 0)
'        .SetLowColour rgba(0, 0, 0, 64)
'        .FillRect pr, MFX_VERT_GRADIENT         ' // darken lower half with an alpha shade
'
'        .SetHighColour rgba(0, 0, 0, 102)
''        .StrokeRect pr.insetbycopy(-1, -1), 2   ' // tray border
'
'
'        ' /* should be a percentage */
'
'        If g_ispercent(mInfo.Text) Then
'            n = g_safelong(minfo.text)
'            If n > 100 Then
'                n = 100
'
'            ElseIf n < 0 Then
'                n = 0
'
'            End If
'
'        Else
'            ' /* not a number so display as 50% to indicate how style looks (especially
'            '    useful for the scheme preview in Snarl) */
'
'            n = 50
'
'        End If
'
'        ' /* meter bar */
'
'        If n > 0 Then
'
'            Set pr = new_brect(0, 0, Fix((pr.Width - 1) * (n / 100)), 6)
'            pr.OffsetBy xOffset + 10, Fix((.Height - pr.Height) / 2)
'
'            If InStr(mInfo.Scheme, "spectrum") Then
'                Select Case n
'                Case Is >= 80
'                    n = rgba(0, 255, 0)
'
'                Case Is >= 60
'                    n = rgba(128, 255, 0)
'
'                Case Is >= 40
'                    n = rgba(255, 255, 0)
'
'                Case Is >= 20
'                    n = rgba(255, 128, 0)
'
'                Case Else
'                    n = rgba(255, 0, 0)
'
'                End Select
'
'            ElseIf InStr(mInfo.Scheme, "colour") Then
'                n = gSettings.ColourGraphColour
'
'            Else
'                ' /* monochrome */
'                n = rgba(199, 197, 198)
'
'            End If
'
'            .SetHighColour n
'            .FillRect pr
'
'            ' /* darken the lower part of the bar */
'
'            If gSettings.ShowDarkShade Then
'                .SetHighColour rgba(0, 0, 0, 0)
'                .SetLowColour rgba(0, 0, 0, 102)
'                .FillRect pr, MFX_VERT_GRADIENT
'
'            End If
'
'            ' /* bas relief */
'
'            If gSettings.ShowBasRelief Then
'                .SetHighColour rgba(255, 255, 255, 62)
'                .SetLowColour rgba(0, 0, 0, 62)
'                .StrokeFancyRect pr
'
'            End If
'
'            ' /* round the corners */
'
'            .SetHighColour rgba(0, 0, 0, 102)
'            .SetPixel prTray.TopLeft
'            .SetPixel prTray.TopRight
'            .SetPixel prTray.BottomLeft
'            .SetPixel prTray.BottomRight
'
'        End If
'
'        ' /* title */
'
'        .SetHighColour rgba(255, 255, 255)
'        .SetLowColour rgba(0, 0, 0)
'
'Dim cyTitle As Long
'
'        If (mInfo.Title <> "") And (gSettings.ShowTitle) Then
'            .SetFont gSettings.zzFontName, gSettings.zzFontSize, True
'            .MeasureString mInfo.Title, .Bounds, pr
'
'            cyTitle = pr.Height
'            Set pr = prTray.OffsetToCopy(0, -(pr.Height + 4))
'            pr.Bottom = pr.Top + cyTitle
'            .DrawString mInfo.Title, pr, MFX_SIMPLE_OUTLINE
'
'        End If
'
'    End With
'
''Dim pb As mfxBitmap
''
''    Set pb = create_dropshadow(mView, , 10)
''
''    mView.SizeTo pb.Width, pb.Height
''    mView.DrawScaledImage pb
'
'End Sub
'
'Private Sub uDrawGraph()
''Dim prOrigin As brect
''Dim pr As brect
''Dim n As Long
''Dim i As Long
''Dim cx As Long
''
''    uDrawBackground2 80             ' // create and draw background with a fixed size of 80px
''    uDrawForeground False           ' // draw icon
''
''    With mView
''
''        ' /* graph */
''
''        ' /* 10 bars @ 10px each with 3px gap */
''
''        cx = (10 * 10) + (9 * 3)
''
''        ' /* figure out where the graph should go */
''
''        Set pr = new_brect(0, 0, cx, 38)
''        pr.OffsetBy mIconSize + Fix(((.Width - mIconSize) - pr.Width) / 2), Fix((.Height - pr.Height) / 2)
''
''        ' /* offset it from the vertical centre so we can fit the title in */
''
''        If mInfo.Title <> "" Then _
''            pr.OffsetBy 0, 4
''
'''        .SetHighColour rgba(255, 0, 0)
'''        .StrokeRect pr
''
''        ' /* draw the individual bar backgrounds */
''
''        Set pr = new_brect(pr.Left, pr.Top + 36, pr.Left + 10, pr.Bottom)
''        Set prOrigin = pr.Duplicate
''
''        For i = 1 To 10
''
''            ' /* bar dark outline */
''            .EnableSmoothing True
''            pr.InsetBy -2, -2
''            .SetHighColour rgba(28, 28, 28)
''            .FillRoundRect pr, 2, 2
''            .EnableSmoothing False
''            pr.InsetBy 2, 2
''
''            ' /* bar background */
''            .SetHighColour rgba(88, 80, 78)
''            .SetLowColour rgba(69, 61, 59)
''            .FillRect pr, MFX_VERT_GRADIENT
''
''            ' /* shade bar using a black alpha gradient */
''
'''            .SetHighColour rgba(0, 0, 0, 0)
'''            .SetLowColour rgba(0, 0, 0, 102)
'''            .FillRect pr, MFX_VERT_GRADIENT
''
''            ' /* round the corners */
''
''            .SetHighColour rgba(0, 0, 0, 102)
''            .SetPixel pr.TopLeft
''            .SetPixel pr.TopRight
''            .SetPixel pr.BottomLeft
''            .SetPixel pr.BottomRight
''
''            pr.OffsetBy 10 + 3, 0
''            pr.Top = pr.Top - 4
''
''        Next i
''
''        ' /* meter bar */
''
''        If g_ispercent(mInfo.Text) Then
''            n = g_safelong(minfo.text)
''            If n > 100 Then
''                g_Debug "StyleInstance.uDrawGraph(): value '" & mInfo.Text & "' constrained to 100", LEMON_LEVEL_WARNING
''                n = 100
''
''            ElseIf n < 0 Then
''                g_Debug "StyleInstance.uDrawGraph(): value '" & mInfo.Text & "' constrained to 0", LEMON_LEVEL_WARNING
''                n = 0
''
''            End If
''
''        Else
''            g_Debug "StyleInstance.uDrawGraph(): value '" & mInfo.Text & "' is not a number", LEMON_LEVEL_WARNING
''            n = 50
''
''        End If
''
''        If n > 0 Then
''
''            If (n Mod 10) >= 5 Then
''                ' /* round up */
''                n = ROUNDUP(n / 10)
''
''            Else
''                ' /* round down */
''                n = Fix(n / 10)
''
''            End If
''
''            Set pr = prOrigin.Duplicate
''            For i = 1 To n
''
''                If InStr(mInfo.Scheme, "spectrum") Then
''
''                    Select Case i
''                    Case Is >= 9
''                        If gSettings.SpectrumType = 0 Then
''                            .SetHighColour rgba(0, 255, 0)
''
''                        Else
''                            .SetHighColour rgba(255, 0, 255)
''
''                        End If
''
''                    Case Is >= 7
''                        If gSettings.SpectrumType = 0 Then
''                            .SetHighColour rgba(128, 255, 0)
''
''                        Else
''                            .SetHighColour rgba(255, 128, 255)
''
''                        End If
''
''                    Case Is >= 5
''                        If gSettings.SpectrumType = 0 Then
''                            .SetHighColour rgba(255, 255, 0)
''
''                        Else
''                            .SetHighColour rgba(255, 255, 255)
''
''                        End If
''
''                    Case Is >= 3
''                        If gSettings.SpectrumType = 0 Then
''                            .SetHighColour rgba(255, 128, 0)
''
''                        Else
''                            .SetHighColour rgba(128, 255, 255)
''
''                        End If
''
''                    Case Else
''                        If gSettings.SpectrumType = 0 Then
''                            .SetHighColour rgba(255, 0, 0)
''
''                        Else
''                            .SetHighColour rgba(0, 255, 255)
''
''                        End If
''
''                    End Select
''
''                    ' /* solid-fill the bar in the selected colour */
''
''                    .FillRect pr, MFX_SOLID_HIGH
''
''                    ' /* shade bar using a black alpha gradient */
''
''                    If gSettings.ShowDarkShade Then
''                        .SetHighColour rgba(0, 0, 0, 0)
''                        .SetLowColour rgba(0, 0, 0, 102)
''                        .FillRect pr, MFX_VERT_GRADIENT
''
''                    End If
''
''                ElseIf InStr(mInfo.Scheme, "colour") Then
''
''                    ' /* fill each bar with the same solid colour */
''
''                    .SetHighColour gSettings.ColourGraphColour
''                    .FillRect pr, MFX_SOLID_HIGH
''
''                    ' /* shade bar using a solid black fill with the alpha reducing on each bar */
''
''                    .SetHighColour rgba(0, 0, 0, 100 - (i * 10))
''                    .FillRect pr, MFX_SOLID_HIGH
''
''                Else
''
''                    ' /* fill with a grey gradient */
''
''                    .SetHighColour rgba(199, 197, 198)
''                    .SetLowColour rgba(154, 150, 149)
''                    .FillRect pr, MFX_HORIZ_GRADIENT
''
''                    ' /* shade bar using a black alpha gradient */
''
''                    If gSettings.ShowDarkShade Then
''                        .SetHighColour rgba(0, 0, 0, 0)
''                        .SetLowColour rgba(0, 0, 0, 102)
''                        .FillRect pr, MFX_VERT_GRADIENT
''
''                    End If
''
''                End If
''
''                ' /* bar bas-relief */
''
''                If gSettings.ShowBasRelief Then
''                    .SetHighColour rgba(255, 255, 255, 32)
''                    .SetLowColour rgba(0, 0, 0, 32)
''                    .StrokeFancyRect pr
''
''                End If
''
''                ' /* round the corners */
''
''                .SetHighColour rgba(0, 0, 0, 102)
''                .SetPixel pr.TopLeft
''                .SetPixel pr.TopRight
''                .SetPixel pr.BottomLeft
''                .SetPixel pr.BottomRight
''
''
''                pr.OffsetBy 10 + 3, 0
''                pr.Top = pr.Top - 4
''
''            Next i
''
''        End If
''
''        ' /* title */
''
''        If (mInfo.Title <> "") And (gSettings.ShowTitle) Then
''            Set pr = new_brect(0, 0, .Width - mIconSize, 20)
''            pr.InsetBy 10, 0
''            pr.OffsetBy mIconSize, 8
''
'''            .SetHighColour rgba(255, 0, 0)
'''            .StrokeRect pr
''
''            .SetHighColour rgba(255, 255, 255)
''            .SetLowColour rgba(0, 0, 0)
''            .SetFont gSettings.zzFontName, gSettings.zzFontSize, True
''            .DrawString mInfo.Title, pr, MFX_ALIGN_H_CENTER Or MFX_SIMPLE_OUTLINE
''
''        End If
''
''    End With
'
'End Sub

'Private Sub uDrawBackground(ByRef Height As Long, ByRef MeterOffset As Long)
'
'    Height = 0
'    MeterOffset = 0
'
'    With mView
'        .SizeTo 250, 80             '// 48 + 20
'        .EnableSmoothing True
'
'        ' /* background */
'
'        .SetHighColour rgba(42, 42, 42)
'        .FillRoundRect .Bounds, 20, 20
'
'        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'            .SetHighColour rgba(0, 0, 0, 0)
'            .SetLowColour gSettings.PriorityBackgroundColour
'            .FillRoundRect .Bounds, 20, 20, MFX_VERT_GRADIENT
'
'        End If
'
'        ' /* edge */
'
'        .SetHighColour rgba(0, 0, 0, 102)
'        .StrokeRoundRect .Bounds, 20, 20
'
''        .SetHighColour rgba(255, 255, 255, 102)
''        .SetLowColour rgba(0, 0, 0, 102)
''        .StrokeFancyRoundRect .Bounds, 20, 20, 2
'
'        .EnableSmoothing False
'
'        If Not (mIcon Is Nothing) Then
'            Height = mIcon.Height
'            If Height > 0 Then
'                Height = MIN(Height, 48)
'
''                .SetHighColour rgba(255, 0, 0)
''                .FillRect new_brect(10, Fix((.Height - height) / 2), 10 + height, Fix((.Height - height) / 2) + height)
'
'                ' /* centre icon? */
'
'                If (mInfo.Title = "") Or (mInfo.Text = "") Then
'                    .DrawScaledImage mIcon, new_bpoint(Fix((.Width - Height) / 2), Fix((.Height - Height) / 2)), new_bpoint(Height, Height)
'
'                Else
'                    .DrawScaledImage mIcon, new_bpoint(10, Fix((.Height - Height) / 2)), new_bpoint(Height, Height)
'
'                End If
'
'                Height = Height + 10
'                MeterOffset = MeterOffset + Height
'
'            End If
'        End If
'
''        .SetHighColour rgba(255, 255, 255)
''        .DrawString Scheme, .Bounds
'
'    End With
'
'End Sub

Private Sub uCalcSize()

    If (mView Is Nothing) Then
        g_Debug "TMinimalInstance.uCalcSize(): view not created", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim cyNotification As Long

    cyNotification = 80
    Set mTitleRect = new_BRect(0, 0, 0, 0)
    Set mTextRect = new_BRect(0, 0, 0, 0)

    ' /* icon-only notifications are a fixed width and height */

    If mIconOnly Then
        mView.SizeTo 250, cyNotification
        Exit Sub

    End If

Dim prTemp As BRect
Dim pr As BRect

    ' /* measure title and text to see how big the notification needs to be */

    With mView

        ' /* create the base bounding rect - this is used to get the title and text height
        '    based on a fixed width */

        Set pr = new_BRect(0, 0, 250, cyNotification)
        pr.InsetBy EM_MARGIN, EM_MARGIN

        ' /* if we have an icon then shrink accordingly */
        If mIconSize > 0 Then _
            pr.Right = pr.Right - (mIconSize + 4)

        ' /* measure title */
        If mInfo.Title <> "" Then
            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size"))
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Title, prTemp, prTemp
            Set mTitleRect = new_BRect(0, 0, prTemp.Width + 2, prTemp.Height + 2)

            If Val(mStyle.ReadSetting("icon_align")) = 0 Then _
                mTitleRect.OffsetBy mIconSize + EM_MARGIN, 0

        End If

        ' /* measure text */
        If (mInfo.Text <> "") And (Not mDrawBar) Then
            .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size"))
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Text, prTemp, prTemp
            Set mTextRect = new_BRect(0, 0, prTemp.Width + 2, prTemp.Height + 2)

            If Val(mStyle.ReadSetting("icon_align")) = 0 Then _
                mTextRect.OffsetBy mIconSize + EM_MARGIN, 0

        End If

        ' /* size the view */

        .SizeTo 250, MAX(mTitleRect.Height + mTextRect.Height + (2 * EM_MARGIN), cyNotification)

    End With

End Sub

Private Sub IStyleInstance_AdjustPosition(X As Long, Y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    LSet mInfo = NotificationInfo

    ' /* clear cached icon */

    Set mIcon = Nothing
    mIconSize = 0

Dim pi As MImage

    ' /* cache icon */

    If Val(mStyle.ReadSetting("show_icon")) Then
        Set pi = load_image_obj(NotificationInfo.Icon)
        If is_valid_image(pi) Then
        Set mIcon = style_MakeSquareImage(pi)
            mIconSize = MIN(mIcon.Height, 48)

        End If

    End If

    mDrawBar = (g_IsPercent(NotificationInfo.Text)) And (Val(mStyle.ReadSetting("show_graphically")) <> 0)

    ' /* pre-configuration */

    mIconOnly = ((mInfo.Title = "") And (mInfo.Text = "")) 'Or (InStr(mInfo.Scheme, "icon only") > 0))

    uCalcSize                       '// calculate notifcation size and internal measurements
    
Const RX = 3

    With mView
        .EnableSmoothing True
        .SetHighColour uGetSchemeBackground(NotificationInfo.Scheme)
        'Val(mstyle.ReadSetting("background"))
        .FillRoundRect .Bounds, RX, RX

    End With

    uDrawIcon

    With mView

        .SetHighColour rgba(255, 255, 255, 102)
        .SetLowColour rgba(255, 255, 255, 0)
        .SetHighlightPosition Fix(.Bounds.Width / 2), Fix(.Bounds.Height / 1)
        .SetHighlightFocusScale 0.1, 0.1
        .FillEllipse .Bounds.InsetByCopy(-48, -48).OffsetByCopy(0, .Bounds.Height / 2), MFX_CENTER_GRADIENT

        .SetHighColour IIf((mInfo.Flags And S_NOTIFICATION_IS_PRIORITY), rgba(255, 0, 0, 252), rgba(255, 255, 255, 252))
        .StrokeRoundRect .Bounds, RX, RX, 2

    End With

    If mIconOnly Then _
        Exit Sub


    uDrawForeground

End Sub

Private Sub uDrawIcon()

    If (mIconSize < 1) Or (Val(mStyle.ReadSetting("show_icon")) = 0) Then _
        Exit Sub

Dim pt As BPoint
Dim pb As mfxBitmap
Dim lCol As Long

    With mView
        If mIconOnly Then
            ' /* centre of the notification */
            Set pt = new_BPoint(Fix((.Width - mIconSize) / 2), Fix((.Height - mIconSize) / 2))

        Else
            If Val(mStyle.ReadSetting("icon_align")) = 0 Then
                ' /* top-left corner */
                Set pt = new_BPoint(EM_MARGIN, EM_MARGIN)

            Else
                ' /* top-right corner */
                Set pt = new_BPoint(.Width - (mIconSize + EM_MARGIN), EM_MARGIN)

            End If
        End If

        ' /* draw the icon */
        .DrawScaledImage mIcon, pt, new_BPoint(mIconSize, mIconSize), 128

        ' /* flip it along the y-axis */
        Set pb = mIcon.Duplicate
        GdipImageRotateFlip pb.Handle, RotateNoneFlipY
        pt.OffsetBy 0, mIconSize

        .DrawScaledImage pb, pt, new_BPoint(mIconSize, mIconSize), 128

'        lCol = Val(mstyle.ReadSetting("background"))
        lCol = uGetSchemeBackground(mInfo.Scheme)
        .SetHighColour rgba(get_red(lCol), get_green(lCol), get_blue(lCol), 128)
        .SetLowColour lCol

        .EnableSmoothing False
        .FillRect new_BRect(pt.X, pt.Y, pt.X + mIconSize, pt.Y + MIN((.Height - pt.Y - 1), mIconSize)), MFX_VERT_GRADIENT
        .EnableSmoothing True

    End With

End Sub

Private Sub uDrawForeground()
Dim lCol As Long

    With mView
        
'        lCol = Val(mstyle.ReadSetting("foreground"))
        lCol = rgba(255, 255, 255)

        .SetHighColour rgba(get_red(lCol), get_green(lCol), get_blue(lCol), (Val(mStyle.ReadSetting("fg_alpha")) / 100) * 255)
        .SetLowColour rgba(102, 102, 102, (Val(mStyle.ReadSetting("fg_alpha")) / 100) * 160)

        ' /* title */

        If mTitleRect.Bottom > 0 Then
            If (mTextRect.Height = 1) And (Not mDrawBar) Then
                ' /* no text so centre vertically */
                mTitleRect.OffsetBy EM_MARGIN, Fix((.Height - mTitleRect.Height) / 2)

            Else
                mTitleRect.OffsetBy EM_MARGIN, EM_MARGIN

            End If

            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size"))
            .DrawString mInfo.Title, mTitleRect, MFX_SIMPLE_OUTLINE

        End If

        ' /* text or bar */

Dim pr As BRect
Dim prLabel As BRect
Dim dw As Single

        If mDrawBar Then
            ' /* text is assumed to represent a value between 0 and 100 (i.e. a percent) */

            .EnableSmoothing False

            Set pr = new_BRect(0, 0, .Width, 4)
            pr.InsetBy EM_MARGIN, 0                             '// left/right margins
            pr.OffsetBy 0, Fix((.Height - pr.Height) / 2)       '// center vertically

            If mIconSize > 0 Then
                ' /* compensate for icon */
                If Val(mStyle.ReadSetting("icon_align")) Then
                    pr.Right = pr.Right - (mIconSize + EM_MARGIN)

                Else
                    pr.Left = pr.Left + (mIconSize + EM_MARGIN)

                End If
            End If

            ' /* draw meter background */

            .SetHighColour rgba(96, 96, 96, 63)
            .FillRect pr

            ' /* get the value */

            dw = MAX(MIN(g_SafeLong(mInfo.Text), 100), 0)
            Set prLabel = pr.Duplicate

            ' /* figure out how big the meter should be and draw it */

            pr.Right = pr.Left + Fix(pr.Width * (dw / 100))
            .SetHighColour rgba(255, 255, 255)
            .FillRect pr.InsetByCopy(1, 1)

            ' /* draw the actual value (not the passed text) below */

            If Val(mStyle.ReadSetting("show_percent")) Then
                .SetHighColour rgba(get_red(lCol), get_green(lCol), get_blue(lCol), (Val(mStyle.ReadSetting("fg_alpha")) / 100) * 255)
                .SetLowColour rgba(102, 102, 102, (Val(mStyle.ReadSetting("fg_alpha")) / 100) * 160)

                .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size"))

                prLabel.OffsetBy 0, prLabel.Height
                prLabel.ExpandBy 0, .StringHeight("A")
'                .StrokeRect prLabel
                .DrawString CStr(dw) & "%", prLabel, MFX_ALIGN_H_CENTER Or MFX_ALIGN_TOP Or MFX_SIMPLE_OUTLINE

            End If

        Else
            ' /* draw text as text */

            If mTitleRect.Height = 1 Then
                ' /* no title so centre vertically */
                mTextRect.OffsetBy EM_MARGIN, Fix((.Height - mTextRect.Height) / 2)

            Else
                mTextRect.OffsetBy EM_MARGIN, EM_MARGIN + mTitleRect.Height

            End If

            .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size"))
            .DrawString mInfo.Text, mTextRect, MFX_SIMPLE_OUTLINE

        End If

    End With

End Sub

Private Function uGetSchemeBackground(ByVal Scheme As String) As Long

    Select Case Scheme

    Case "koka"
        uGetSchemeBackground = rgba(190, 105, 65)

    Case "lyme"
        uGetSchemeBackground = rgba(123, 216, 92)

    Case "charkoal"
        uGetSchemeBackground = rgba(92, 92, 92)

    Case Else
        uGetSchemeBackground = rgba(130, 138, 196)

    End Select

End Function


