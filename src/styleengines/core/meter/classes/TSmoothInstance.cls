VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSmoothInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const EM_TITLE_TEXT = 3
Const EM_MARGIN = 10
Const EM_LEFT_MARGIN = 32

Dim mStyle As TStyle

Dim mInfo As notification_info
Dim mIconSize As Long
Dim mIcon As mfxBitmap
Dim mView As mfxView

Dim mIconOnly As Boolean
Dim mTint As Long                   ' // colour-tint
Dim mValue As Long                  ' // value-percent

Dim mTitleRect As BRect
Dim mTextRect As BRect

Implements IStyleInstance
Implements IMeterStyle

Private Sub IMeterStyle_SetStyle(Style As TStyle)

    Set mStyle = Style

End Sub

Private Sub IStyleInstance_AdjustPosition(X As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    If ISNULL(mView) Then
        Set mView = New mfxView
        mView.TextMode = MFX_TEXT_ANTIALIAS

    End If

    LSet mInfo = NotificationInfo

    With New BPackedData
        .SetTo NotificationInfo.Text
        mInfo.Title = .ValueOf("title")
        mInfo.Text = .ValueOf("text")

        mValue = -1
        If .Exists("value-percent") Then
            mValue = g_SafeLong(.ValueOf("value-percent"))
            If (mValue < 0) Or (mValue > 100) Then _
                mValue = -1

        End If

        mTint = 0
        If .Exists("colour-tint") Then _
            mTint = g_SafeLong(.ValueOf("colour-tint"))


    End With

Dim pi As MImage

    ' /* cache icon */

    Set mIcon = Nothing
    mIconSize = 0
    Set pi = load_image_obj(NotificationInfo.Icon)
    If is_valid_image(pi) Then _
        Set mIcon = create_bitmap_from_image(pi)

    ' /* pre-configuration */

    mIconOnly = ((mInfo.Title = "") And (mInfo.Text = "")) 'Or (InStr(mInfo.Scheme, "icon only") > 0))

    If mIconOnly Then
        ' /* icon-only notifications are a fixed width and height */
        mView.SizeTo 250, 80

    Else
        ' /* calculate notifcation size and internal measurements */
        uCalcSize

    End If

    uDraw

End Sub

Private Sub uCalcSize()
Dim prTemp As BRect
Dim pr As BRect
Dim cy As Long

    ' /* measure title and text to see how big the notification needs to be */

    With mView
        .SizeTo 250, 900
        cy = EM_MARGIN * 2

        ' /* create the base bounding rect - this is used to get the title and text height based on a fixed width */
        Set pr = new_BRect(0, 0, .Bounds.Right, 900 - 1).InsetByCopy(EM_MARGIN, 0)
        pr.Left = EM_LEFT_MARGIN

        ' /* measure title */
        If mInfo.Title <> "" Then
            .SetFont mStyle.ReadSetting("title_font"), mStyle.ReadSetting("title_size"), True
            .MeasureString mInfo.Title, pr.Duplicate, prTemp
            Set mTitleRect = new_BRect(0, 0, prTemp.Width, prTemp.Height)
            cy = cy + mTitleRect.Height

        End If

        ' /* measure text */
        If mInfo.Text <> "" Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size"), True
            .MeasureString mInfo.Text, pr.Duplicate, prTemp
            Set mTextRect = new_BRect(0, 0, prTemp.Width, prTemp.Height)
            cy = cy + mTextRect.Height

        End If

        ' /* if we have title and text, add separator */
        If (mInfo.Title <> "") And (mInfo.Text <> "") Then _
            cy = cy + EM_TITLE_TEXT

        ' /* add on space for meter, if required */
        If mValue > -1 Then _
            cy = cy + 10 + EM_TITLE_TEXT

        ' /* size the view */
        .SizeTo 250, Max(cy, 80)

    End With

End Sub

Private Sub uDraw()
Const RX = 5

    With mView
        ' /* now we have the final notification height, calculate the size of the icon */
        If is_valid_image(mIcon) Then _
            mIconSize = Min(mIcon.Height, .Height + (2 * EM_MARGIN))

        ' /* draw the background */
        .Clear
        .EnableSmoothing True

        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            ' /* priority notifications always get a red background */
            .SetHighColour rgba(223, 0, 0)

        Else
            .SetHighColour mStyle.GetSchemeColour(mInfo.Scheme, "background")

        End If

        .FillRoundRect .Bounds, RX, RX
        .SetHighColour rgba(255, 255, 255, 96)
        .SetLowColour rgba(255, 255, 255, 0)
        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT
        .SetHighColour rgba(0, 0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 96)
        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT
        .SetHighColour rgba(255, 255, 255, 64)
        .SetLowColour rgba(0, 0, 0, 0)
        .StrokeFancyRoundRect .Bounds.InsetByCopy(1, 1), RX, RX
        .SetHighColour rgba(0, 0, 0, 64)
        .StrokeRoundRect .Bounds, RX, RX

        If mIconSize > 0 Then _
            .DrawScaledImage mIcon, new_BPoint(-(mIconSize * 0.25), Fix((mView.Height - mIconSize) / 2)), new_BPoint(mIconSize, mIconSize), 170

        ' /* draw the foreground */

        If Not mIconOnly Then _
            uDrawForeground

        ' /* draw tint *

        If mTint <> 0 Then _
            .DrawScaledImage g_CreateMarker(mTint, 18), new_BPoint(.Width - 28, 0), , 180

    End With

End Sub

Private Sub uDrawForeground()
Dim pr As BRect

    With mView
        .SetLowColour g_SetRGB32Alpha(g_SafeLong(mStyle.ReadSetting("border")), 48)

        ' /* title */
        If NOTNULL(mTitleRect) Then
            Set pr = mTitleRect.Duplicate
            If ISNULL(mTextRect) Then
                ' /* no text so centre vertically */
                pr.OffsetBy EM_LEFT_MARGIN, Fix((.Height - mTitleRect.Height) / 2)

            Else
                pr.OffsetBy EM_LEFT_MARGIN, EM_MARGIN

            End If

            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
            .SetHighColour g_SafeLong(mStyle.ReadSetting("foreground"))
            .DrawString mInfo.Title, pr, MFX_SIMPLE_OUTLINE

        End If

        ' /* text */
        If NOTNULL(mTextRect) Then
            Set pr = mTextRect.Duplicate
            If ISNULL(mTitleRect) Then
                ' /* no title so centre vertically */
                pr.OffsetBy EM_LEFT_MARGIN, Fix((.Height - mTextRect.Height) / 2)

            Else
                ' /* has title (and text) so draw separator */
                pr.OffsetBy EM_LEFT_MARGIN, EM_MARGIN + mTitleRect.Height
                .SetHighColour rgba(255, 255, 255, 170)
                .StrokeLine new_BRect(pr.Left, pr.Top - 0, .Bounds.Right - EM_MARGIN, pr.Top - 0)
                pr.OffsetBy 0, EM_TITLE_TEXT

            End If

            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size"), True
            .SetHighColour g_SetRGB32Alpha(g_SafeLong(mStyle.ReadSetting("foreground")), 250)
            .DrawString mInfo.Text, pr, MFX_SIMPLE_OUTLINE

        End If

        ' /* meter */
        If mValue > -1 Then
            Set pr = new_BRect(EM_LEFT_MARGIN, 0, .Width - EM_MARGIN, 10 - 1).OffsetByCopy(0, (.Height - (10 + EM_MARGIN)))
            .SetHighColour rgba(0, 0, 0, 96)
            .EnableSmoothing True
            .FillRoundRect pr, 6, 6
            If mValue > 4 Then
                .SetHighColour rgba(255, 255, 255, 202)
                pr.InsetBy 2, 2
                pr.Right = pr.Left + Fix(pr.Width * (mValue / 100))
                .FillRoundRect pr, 4, 4

            End If
        End If

    End With

End Sub


