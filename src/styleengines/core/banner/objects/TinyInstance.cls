VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TinyInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim theView As mfxView
Dim mShowing As Boolean
Dim mYPos As Single
Dim mInfo As notification_info

Public Style As TStyle

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
Const STEP_SIZE = 1 / 4

    ' /* on entry:
    '       x, y are set to the position Snarl has determined is best for the notification
    '       Alpha is set to the maximum opacity this notification can reach
    '       Done is set to True
    '
    '   In order for Snarl to keep polling AdjustPosition we must set 'Done' to False each
    '   time until we're done positioning the notification
    '
    ' */

    If InStr(mInfo.Scheme, "left") Then
        x = 64

    ElseIf InStr(mInfo.Scheme, "right") Then
        x = GetSystemMetrics(SM_CXSCREEN) - (theView.Width + 64)

    Else
        x = Fix((GetSystemMetrics(SM_CXSCREEN) - theView.Width) / 2)

    End If



    If InStr(mInfo.Scheme, "top") Then

        If mShowing Then
            If mYPos < 0 Then
                ' /* keep going... */
                mYPos = mYPos + STEP_SIZE
                y = ROUNDUP(mYPos)
                Done = False

            End If

        Else
            ' /* hiding */
            If mYPos < theView.Height Then
                ' /* keep going... */
                mYPos = mYPos + (STEP_SIZE * 2)
                y = -ROUNDUP(mYPos)
                Done = False

            End If

        End If

    Else

        If mShowing Then
            If mYPos > (g_ScreenHeight - theView.Height) Then
                ' /* keep going... */
                mYPos = mYPos - STEP_SIZE
                y = ROUNDUP(mYPos)
                Done = False

            End If

        Else
            ' /* hiding */
            If mYPos < (g_ScreenHeight + theView.Height) Then
                ' /* keep going... */
                mYPos = mYPos + (STEP_SIZE * 2)
                y = ROUNDUP(mYPos)
                Done = False

            End If

        End If
    
    End If


End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = theView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
Dim rc As RECT

    mShowing = Visible
    g_GetWorkArea rc

    If InStr(mInfo.Scheme, "bottom") Then
        mYPos = IIf(mShowing, g_ScreenHeight() + theView.Height, g_ScreenHeight() - theView.Height)

    Else
        mYPos = IIf(mShowing, -theView.Height, 0)

    End If

End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Const MARGIN_SIZE = 3
Const STAR_SIZE = 12
Const TEXT_GAP = 0
Const RX = 9

    LSet mInfo = NotificationInfo

    ' /* we're (now) a V42 style so notificationinfo->title = App name; notificationinfo->text = content */

Dim pContent As BPackedData
Dim szTitle As String
Dim szText As String
Dim szSubText As String
Dim cStars As Integer
Dim s As Single

    Set pContent = New BPackedData
    With pContent
        .SetTo NotificationInfo.Text

        szTitle = .ValueOf("title")
        szText = .ValueOf("text")
        szSubText = .ValueOf("label-subtext")

        If .Exists("value-percent") Then
            s = Val(.ValueOf("value-percent"))
            If (s < 0) Or (s > 100) Then _
                s = -1

        Else
            s = -1

        End If

        If s > -1 Then
            cStars = Fix(s / 20)

        Else
            cStars = -1

        End If

    End With

    ' /* get icon and calc text offsets from it */

Dim pIcon As mfxBitmap
Dim cxIcon As Long
Dim xOffset As Long

    Set pIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))
    If is_valid_image(pIcon) Then
        cxIcon = Min(pIcon.Width, g_SafeLong(Style.ReadSetting("icon_size")))
        xOffset = cxIcon + 3                    ' // horizontal text offset

    End If

Dim xIcon As Long
Dim pr As BRect
Dim cy As Long
Dim yMargin As Long
Dim cyContent As Long               ' // the sum total height of title, text, subtext and meter with corresponding gaps

    Set theView = New mfxView
    With theView

        ' /* determine the required height */

        ' /* font mode */
        If Style.ReadSetting("use_cleartype") = "1" Then
            .TextMode = MFX_TEXT_CLEARTYPE

        Else
            .TextMode = MFX_TEXT_ANTIALIAS

        End If

        ' /* title */
        If szTitle <> "" Then
            .SetFont Style.ReadSetting("title_font"), Val(Style.ReadSetting("title_size")), (Style.ReadSetting("title_bold") = "1")
            cyContent = cyContent + .StringHeight("A") + TEXT_GAP

        End If

        ' /* text */
        If szText <> "" Then
            .SetFont Style.ReadSetting("text_font"), Val(Style.ReadSetting("text_size")), (Style.ReadSetting("text_bold") = "1")
            cyContent = cyContent + .StringHeight("A") + TEXT_GAP

        End If

        ' /* meter */
        If cStars > -1 Then _
            cyContent = cyContent + STAR_SIZE + TEXT_GAP

        ' /* subtext - uses same font as text but drawn in 90% alpha */
        If szSubText <> "" Then
            .SetFont Style.ReadSetting("text_font"), Val(Style.ReadSetting("text_size")), (Style.ReadSetting("text_bold") = "1")
            cyContent = cyContent + .StringHeight("A") + TEXT_GAP

        End If

        ' /* tidy content size up */
        If cyContent > 0 Then _
            cyContent = cyContent - TEXT_GAP

        cyContent = cyContent + (2 * MARGIN_SIZE)

        ' /* must be at least big enough for the icon */
        cy = Max(cyContent, cxIcon)
        .SizeTo 320, cy

        If (szText = "") And (szTitle = "") Then
            ' /* special case: icon only */
            xIcon = Fix((.Width - cxIcon) / 2)

        Else
            xIcon = 4

        End If

        If cyContent < cy Then
            yMargin = Fix((cy - cyContent) / 2)

        Else
            yMargin = MARGIN_SIZE

        End If

        ' /* --- draw --- */


        ' /* background */
        .EnableSmoothing True
        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(240, 0, 0)

        Else
            .SetHighColour g_SafeLong(Style.ReadSetting("col-background"))

        End If
        .FillRoundRect .Bounds, RX, RX

'        .SetHighColour rgba(255, 255, 255, 127)
'        .SetLowColour rgba(0, 0, 0, 127)
'        .StrokeFancyRoundRect .Bounds, RX, RX

        ' /* border */
        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(255, 255, 255)

        Else
            .SetHighColour g_SafeLong(Style.ReadSetting("col-border"))

        End If
        .StrokeRoundRect .Bounds, RX, RX, 1

Dim dCol As Long

        ' /* shading */
        dCol = g_SafeLong(Style.ReadSetting("col-shading"))
        If Style.ReadSetting("invert_shading") = "1" Then
            .SetHighColour g_SetRGB32Alpha(dCol, 128)
            .SetLowColour g_SetRGB32Alpha(dCol, 0)

        Else
            .SetHighColour g_SetRGB32Alpha(dCol, 0)
            .SetLowColour g_SetRGB32Alpha(dCol, 128)

        End If
        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT

'        .SetHighColour rgba(0, 0, 0)
'        .FillRoundRect .Bounds, RX, RX
'        .FillRect new_BRect(0, 0, .Width - 1, (RX / 2) - 1)         ' // square-up the top

'        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'            .SetHighColour rgba(255, 0, 0, 0)
'            .SetLowColour rgba(255, 0, 0)
'            .FillRoundRect .Bounds, rx, rx, MFX_VERT_GRADIENT
''            .SetHighColour rgba(255, 0, 0)
''            .StrokeRoundRect .Bounds, RX, RX
'
'        End If

        ' /* icon */
        If Not (pIcon Is Nothing) Then _
            .DrawScaledImage pIcon, new_BPoint(xIcon, Fix((.Height - cxIcon) / 2)), new_BPoint(cxIcon, cxIcon), 240

        ' /* labels */
        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(255, 255, 255)

        Else
            .SetHighColour g_SafeLong(Style.ReadSetting("col-foreground"))

        End If

        ' /* tint */
        If pContent.Exists("colour-tint") Then
            dCol = g_SafeLong(pContent.ValueOf("colour-tint"))
            If dCol <> 0 Then _
                .DrawScaledImage g_CreateMarker(dCol, 16), new_BPoint(.Width - 24, 0), , 230

        End If

        ' /* create the base rect */
        Set pr = .Bounds.InsetByCopy(MARGIN_SIZE, 0)
        pr.Left = pr.Left + xOffset
        pr.OffsetBy 0, yMargin

        ' /* title */
        If szTitle <> "" Then
            .SetFont Style.ReadSetting("title_font"), Val(Style.ReadSetting("title_size")), (Style.ReadSetting("title_bold") = "1")
            .DrawString .GetFormattedText(szTitle, pr.Width, True), pr, MFX_ALIGN_LEFT
            pr.OffsetBy 0, .StringHeight("A") + TEXT_GAP

        End If

        ' /* text */
        If szText <> "" Then
            .SetFont Style.ReadSetting("text_font"), Val(Style.ReadSetting("text_size")), (Style.ReadSetting("text_bold") = "1")
            .DrawString .GetFormattedText(szText, pr.Width, True), pr, MFX_ALIGN_LEFT
            pr.OffsetBy 0, .StringHeight("A") + TEXT_GAP

        End If

Dim pImg As mfxBitmap
Dim i As Integer

        ' /* meter */
        If cStars > -1 Then
            Set pImg = load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "star.png")
            If NOTNULL(pImg) Then
                For i = 0 To 4
                    If i > (cStars - 1) Then _
                        pImg.MakeGreyscale

                    .DrawScaledImage pImg, pr.TopLeft.OffsetByCopy(i * (STAR_SIZE + 1), 0), new_BPoint(STAR_SIZE, STAR_SIZE)
        
                Next i
            End If
            pr.OffsetBy 0, STAR_SIZE + TEXT_GAP

        End If

        ' /* subtext */
        If szSubText <> "" Then
            .SetHighColour g_SetRGB32Alpha(g_SafeLong(Style.ReadSetting("col-foreground")), 230)
            .SetFont Style.ReadSetting("text_font"), Val(Style.ReadSetting("text_size")), (Style.ReadSetting("text_bold") = "1")
            .DrawString .GetFormattedText(szSubText, pr.Width, True), pr, MFX_ALIGN_LEFT
            pr.OffsetBy 0, .StringHeight("A") + TEXT_GAP

        End If

        ' /* glass effect */
        If Style.ReadSetting("use_glass") = "1" Then
            .SetHighColour rgba(255, 255, 255, 63)
            .FillRoundRect new_BRect(0, 0, .Bounds.Right, Fix(.Bounds.Height / 2)), RX, RX

        End If

    End With

End Sub

