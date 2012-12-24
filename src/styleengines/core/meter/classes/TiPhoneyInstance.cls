VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TiPhoneyInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    ' /*
    '
    '   TiPhoneyInstance.cls -- Represents a single instance of an iPhoney Style notification
    '
    '   © 2010 full phat products.  Some Rights Reserved.
    '   Licenced as per the Simplified BSD (http://www.opensource.org/licenses/bsd-license.php)
    '
    '   Notes
    '   -----
    '
    '   This style utilises the new (introduced with Snarl R2.4) S_STYLE_V42_CONTENT flag
    '   to allow for more advanced notification content.  Enabling this flag causes Snarl
    '   to send the entire unabridged packed data created by sn41[EZ]Notify() and
    '   sn41[EZ]Update() in notification_info.Text.
    '
    ' */

Const EM_TITLE_TEXT = 0
Const EM_MARGIN = 8
Const EM_METER = EM_MARGIN / 2

Dim mStyle As TStyle
Dim mScheme As String
Dim mFlags As S_NOTIFICATION_FLAGS

'Dim mInfo As notification_info
Dim mData As BPackedData
Dim mIcon As MImage
Dim mIconSize As Long
Dim mTitleRect As BRect
Dim mTextRect As BRect

Dim mView As mfxView
Dim mRed As Integer

Dim mIconOnly As Boolean
Dim mIsNumeric As Boolean
Dim mTintCol As Long

Implements IStyleInstance
Implements IMeterStyle

Private Sub Class_Initialize()

    Set mView = New mfxView
    mRed = 160

    With mView
        .TextMode = MFX_TEXT_ANTIALIAS

    End With

End Sub

Private Sub IMeterStyle_SetStyle(Style As TStyle)

    Set mStyle = Style

End Sub

Private Sub IStyleInstance_AdjustPosition(X As Long, Y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
Static nStep As Integer
Const STEP_SIZE = 8

    If (mFlags And S_NOTIFICATION_IS_PRIORITY) Then

        If nStep = 0 Then _
            nStep = STEP_SIZE

        mRed = mRed + nStep
        If mRed > 254 Then
            nStep = -STEP_SIZE

        ElseIf mRed < 161 Then
            nStep = STEP_SIZE

        End If

        uDraw
        IStyleInstance_Pulse = True

    End If

End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    ' /* note: notificationinfo->scheme is translated to lowercase by Snarl */

    If (mView Is Nothing) Then
        Set mView = New mfxView
        mView.TextMode = MFX_TEXT_ANTIALIAS

    End If

    Set mData = New BPackedData
    mData.SetTo NotificationInfo.Text
    mScheme = NotificationInfo.Scheme
    mFlags = NotificationInfo.Flags

    ' /* cache icon */

    mIconSize = 0
    If mData.Exists("icon") Then
        Set mIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))
        If is_valid_image(mIcon) Then _
            mIconSize = MIN(mIcon.Height, 48)

    End If

    If mData.Exists("colour-tint") Then
        mTintCol = g_SafeLong(mData.ValueOf("colour-tint"))

    Else
        mTintCol = 0

    End If

    ' /* pre-configuration */

    mIconOnly = ((mData.ValueOf("title") = "") And (mData.ValueOf("text") = ""))
    mIsNumeric = ((g_IsPercent(mData.ValueOf("text"))) And (Val(mStyle.ReadSetting("show_graphically")) <> 0))

    ' /* calculate notifcation size and internal measurements */

    uCalcSize
    uDraw

End Sub

Private Sub uCalcSize()

    If (mView Is Nothing) Then
        g_Debug "TSonyInstance.uCalcSize(): view not created", LEMON_LEVEL_CRITICAL
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
Dim cy As Long

    ' /* measure title and text to see how big the notification needs to be */

    With mView

        ' /* create the base bounding rect - this is used to get the title and text height
        '    based on a fixed width */

        Set pr = new_BRect(0, 0, 250 - 1, cyNotification - 1)
        pr.InsetBy EM_MARGIN, EM_MARGIN
        pr.Right = pr.Right - mIconSize

        ' /* measure title */

        If mData.ValueOf("title") <> "" Then
            .SetFont mStyle.ReadSetting("title_font"), mStyle.ReadSetting("title_size"), True
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mData.ValueOf("title"), prTemp, prTemp
            Set mTitleRect = new_BRect(0, 0, prTemp.Width, prTemp.Height)

        End If

        ' /* measure text */

        If (mData.ValueOf("text") <> "") And (Not mIsNumeric) Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mData.ValueOf("text"), prTemp, prTemp
            Set mTextRect = new_BRect(0, 0, prTemp.Width, prTemp.Height)

        ElseIf mIsNumeric Then
            cy = EM_MARGIN + EM_METER

        End If

        ' /* size the view */

        .SizeTo 250, MAX(mTitleRect.Height + mTextRect.Height + (2 * EM_MARGIN) + EM_TITLE_TEXT + cy, cyNotification)

    End With

End Sub

Private Sub uDraw()
Dim pr As BRect
Const RX = 20

    With mView
        .Clear
        .EnableSmoothing True

        ' /* background */

        If (mFlags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(mRed, 0, 0)
            .FillRoundRect .Bounds, RX, RX

        Else

            Select Case mScheme
            Case "black"
                .SetHighColour rgba(24, 24, 24)

            Case "original"
                .SetHighColour rgba(39, 51, 76)

            Case "glass"
                .SetHighColour rgba(0, 0, 0, 32)

            Case "psp"
                .SetHighColour g_GetPSPColour()

            Case Else
                ' /* assume "custom" */
                .SetHighColour Val(mStyle.ReadSetting("background"))

            End Select

            .FillRoundRect .Bounds, RX, RX

        End If

        If mTintCol <> 0 Then _
            .DrawScaledImage g_CreateMarker(mTintCol, 36), new_BPoint(12, 0), , 245

'        .SetHighColour rgba(255, 0, 0)
'        .StrokeRect mTitleRect
'
'        .SetHighColour rgba(0, 255, 0)
'        .StrokeRect mTextRect


Dim pbm As mfxBitmap
Dim pt As BPoint
Dim dw As Long

        ' /* icon */

        If mIconSize Then
            If mIconOnly Then
                Set pt = new_BPoint(Fix((.Width - mIconSize) / 2), Fix((.Height - mIconSize) / 2))

            Else
                Set pt = new_BPoint(.Width - (mIconSize + 10), IIf((gSettings.CentreIcon) Or (mIsNumeric), Fix((.Height - mIconSize) / 2), 10))

            End If

            .DrawScaledImage mIcon, pt, new_BPoint(mIconSize, mIconSize)

            ' /* draw inverted (reflected) image */

            Set pbm = create_bitmap_from_image(mIcon)
            If is_valid_image(pbm) Then
                Set pbm = pbm.SizeTo(mIconSize, mIconSize)
                GdipImageRotateFlip pbm.Handle, RotateNoneFlipY
                pt.OffsetBy 0, mIconSize
                dw = MIN(mIconSize, .Height - pt.Y - 1)
                .DrawImage pbm, new_BRect(pt.X, pt.Y, mIconSize, dw), new_BRect(0, 0, mIconSize, dw)

                ' /* draw shading over image */

                .EnableSmoothing False
                .SetLowColour .HighColour
                .SetHighColour rgba(get_red(.HighColour), get_green(.HighColour), get_blue(.HighColour), 190)

'                .SetHighColour rgba(255, 0, 0)
'                .SetLowColour rgba(0, 0, 255)

                .FillRect new_BRect(pt.X, pt.Y, pt.X + mIconSize, MIN(pt.Y + mIconSize, .Height - 2)), MFX_VERT_GRADIENT

            End If

            .EnableSmoothing True

        End If

        ' /* background lighting */

        Set pr = .Bounds.Duplicate
        pr.Right = Fix(pr.Width / 2) - 1
        .SetHighColour rgba(255, 255, 255, 0)
        .SetLowColour rgba(255, 255, 255, 76)
        .FillRect pr, MFX_HORIZ_GRADIENT

        pr.OffsetBy Fix(.Bounds.Width / 2), 0
        .SetHighColour rgba(255, 255, 255, 76)
        .SetLowColour rgba(255, 255, 255, 0)
        .FillRect pr, MFX_HORIZ_GRADIENT

        ' /* edge */

        .SetHighColour rgba(255, 255, 255, 160)
        .StrokeRoundRect .Bounds, RX, RX, 2

        ' /* title */

        If mTitleRect.Bottom > 0 Then
            Set pr = mTitleRect.Duplicate
            If (mTextRect.Height = 1) And (Not mIsNumeric) Then
                ' /* no text (or meter) so centre vertically */
                pr.OffsetBy EM_MARGIN, Fix((.Height - mTitleRect.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN, EM_MARGIN

            End If

            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
            .SetHighColour rgba(255, 255, 255)
'            .SetLowColour rgba(0, 0, 0)
            
            dw = Val(mStyle.ReadSetting("fg_alpha"))
            .SetHighColour rgba(255, 255, 255, Fix((dw / 100) * 255))
            
            .DrawString mData.ValueOf("title"), pr ', MFX_SIMPLE_OUTLINE

        End If

        ' /* text */

Const RXBAR = 6
Dim prLabel As BRect

        If mIsNumeric Then
            ' /* draw a bar instead */

            Set pr = new_BRect(0, 0, .Width - (mIconSize + (3 * EM_MARGIN)), EM_MARGIN - 1) ' // 3*10 = 2x10 margins + 1x10 icon gap
            
            If (mTitleRect.Height = 1) Or (mTitleRect.Bottom + EM_METER + EM_MARGIN < Fix(.Height / 2)) Then
                ' /* no title or title is smaller than the centre of the notification so center vertically */
                pr.OffsetBy EM_MARGIN, Fix((.Height - pr.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN, mTitleRect.Height + EM_MARGIN + EM_METER

            End If

            ' /* background */

            .SetHighColour rgba(255, 255, 255, 60)
            .FillRoundRect pr, RXBAR, RXBAR
            .SetHighColour rgba(0, 0, 0, 60)
            .StrokeRoundRect pr, RXBAR, RXBAR

            ' /* value part */

            dw = g_SafeLong(mData.ValueOf("text"))
            If dw > 0 Then
                Set prLabel = pr.Duplicate
                dw = Fix((dw / 100) * pr.Width)
                If dw > RXBAR Then
                    pr.Right = pr.Left + dw
                    .SetHighColour rgba(255, 255, 255, 140)
                    .FillRoundRect pr, RXBAR, RXBAR
                    .SetHighColour rgba(0, 0, 0, 60)
                    .StrokeRoundRect pr, RXBAR, RXBAR

                End If
            End If

            ' /* display value? */

            If Val(mStyle.ReadSetting("show_percent")) Then
                .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
                .SetHighColour rgba(255, 255, 255, Fix((Val(mStyle.ReadSetting("fg_alpha")) / 100) * 255))

                prLabel.OffsetBy 0, prLabel.Height
                prLabel.Bottom = prLabel.Top + .StringHeight("A")
                .DrawString CStr(CInt(g_SafeLong(mData.ValueOf("text")))) & "%", prLabel, MFX_ALIGN_RIGHT Or MFX_ALIGN_V_CENTER

            End If

        ElseIf mTextRect.Bottom > 1 Then
            ' /* have some text to draw */

            Set pr = mTextRect.Duplicate
            If mTitleRect.Height = 1 Then
                ' /* no title so centre vertically */
                pr.OffsetBy EM_MARGIN, Fix((.Height - mTextRect.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN, EM_MARGIN + mTitleRect.Height + EM_TITLE_TEXT

            End If

            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            dw = Val(mStyle.ReadSetting("fg_alpha"))
            .SetHighColour rgba(255, 255, 255, Fix((dw / 100) * 255))
            .DrawString mData.ValueOf("text"), pr

        End If

        ' /* glass reflection effect */

Dim pv As mfxView

        Set pv = New mfxView
        With pv
            .SizeTo mView.Width, 30
            .EnableSmoothing True
            .SetHighColour rgba(255, 255, 255, 140)
            .SetLowColour rgba(255, 255, 255, 10)
            .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT

        End With

        .DrawImage pv.ConvertToBitmap()

    End With

End Sub



