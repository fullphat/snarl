VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDroidInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    ' /*
    '
    '   TDroidInstance.cls -- Represents a single instance of a Droid Style notification
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

Const EM_MARGIN = 8
Const EM_TITLE_TEXT = 5

Dim mStyle As TStyle
Dim mInfo As notification_info
Dim mView As mfxView
Dim mRed As Integer

Dim mTitle As String
Dim mTitleRect As BRect
Dim mText As String
Dim mTextRect As BRect
Dim mIcon As mfxBitmap
Dim mIconSize As Long
Dim mIconPos As BPoint
Dim mPercent As Single
Dim mMeterRect As BRect

Dim mBkCol As Long
Dim mTintCol As Long

Dim mContent As BPackedData
Dim mLabel() As T_LABEL
Dim mLabels As Long

Implements IStyleInstance
Implements IMeterStyle

Private Sub Class_Initialize()

    mRed = 160
    mBkCol = rgba(64, 64, 64)

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

    If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then

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

    ReDim mLabel(0)
    mLabels = 0

    If (mView Is Nothing) Then
        Set mView = New mfxView
        mView.TextMode = MFX_TEXT_ANTIALIAS
        mView.TextMode = MFX_TEXT_CLEARTYPE

    End If

    ' /* notificationinfo->scheme is translated to lowercase by Snarl */

    LSet mInfo = NotificationInfo

    ' /* cache icon */

    mIconSize = 0
    Set mIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))
    If is_valid_image(mIcon) Then _
        mIconSize = MIN(mIcon.Height, 24)

    ' /* pre-configuration */

Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo NotificationInfo.Text

    If ppd.Exists("title") Then _
        mTitle = ppd.ValueOf("title")

    If ppd.Exists("text") Then _
        mText = ppd.ValueOf("text")


    ' /* colour */

    If ppd.Exists("colour-background") Then _
        mBkCol = g_RGBAFromString(ppd.ValueOf("colour-background"))

    If get_alpha(mBkCol) = 0 Then _
        mBkCol = rgba(64, 64, 64)           ' // zero alpha means reset to default


    If ppd.Exists("colour-tint") Then
        mTintCol = g_SafeLong(ppd.ValueOf("colour-tint"))

    Else
        mTintCol = 0

    End If

    ' /* percent */

    If ppd.Exists("value-percent") Then
        mPercent = Val(ppd.ValueOf("value-percent"))
        If (mPercent < 0#) Or (mPercent > 100#) Then _
            mPercent = -1#

    ElseIf (mStyle.ReadSetting("show_graphically") = "1") And (g_IsNumeric(mText)) Then
        mPercent = mText
        mText = ""
        If (mPercent < 0#) Or (mPercent > 100#) Then _
            mPercent = -1#

    Else
        mPercent = -1#                          ' // means we don't need a meter...

    End If

    Set mContent = ppd

    ' /* calculate notifcation size and internal measurements */

    uCalcSize
    uDraw

End Sub

Private Sub uCalcSize()

    If (mView Is Nothing) Then _
        Exit Sub

Dim cyNotification As Long

    cyNotification = 3 * EM_MARGIN
    Set mTitleRect = new_BRect(0, 0, 0, 0)
    Set mTextRect = new_BRect(0, 0, 0, 0)

    mView.SizeTo 210, cyNotification

Dim Y As Long
Dim prTemp As BRect
Dim pr As BRect
Dim fNeedGap As Boolean

    ' /* measure title and text to see how big the notification needs to be */

    With mView

        Y = EM_MARGIN
        .TextMode = MFX_TEXT_ANTIALIAS

        ' /* create the base bounding rect - this is used to get the title and text height
        '    based on a fixed width */

        Set pr = new_BRect(0, 0, .Width - 1, cyNotification - 1).InsetByCopy(EM_MARGIN, EM_MARGIN)

        ' /* measure title */

        If mTitle <> "" Then
            .SetFont mStyle.ReadSetting("title_font"), mStyle.ReadSetting("title_size")
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mTitle, prTemp, prTemp
            Set mTitleRect = new_BRect(EM_MARGIN, Y, .Width - EM_MARGIN - 1, Y + prTemp.Height - 1)
            fNeedGap = True

            Y = Y + prTemp.Height

        End If

        ' /* measure text */

        If mText <> "" Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mText, prTemp, prTemp

            If fNeedGap Then _
                Y = Y + EM_TITLE_TEXT

            Set mTextRect = new_BRect(EM_MARGIN, Y, .Width - EM_MARGIN - 1, Y + prTemp.Height - 1)

            fNeedGap = True
            Y = Y + prTemp.Height

        End If

        ' /* icon */

        If mIconSize > 0 Then
            If fNeedGap Then _
                Y = Y + EM_TITLE_TEXT

            Set mIconPos = new_BPoint(Fix((.Width - mIconSize) / 2), Y)

            fNeedGap = True
            Y = Y + mIconSize

        End If

        ' /* meter */

        If mPercent > -1 Then
            If fNeedGap Then _
                Y = Y + EM_TITLE_TEXT

            Set mMeterRect = new_BRect(EM_MARGIN, Y, .Width - EM_MARGIN - 1, Y + 10 - 1)
            fNeedGap = True
            Y = Y + mMeterRect.Height

        End If


        ' /* label-* */

        .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")

Dim szn As String
Dim szv As String

        With mContent
            .Rewind
            Do While .GetNextItem(szn, szv)
                If (g_SafeLeftStr(szn, 6) = "label-") And (szv <> "") Then
                    Set prTemp = pr.Duplicate
                    prTemp.Left = EM_MARGIN
                    prTemp.Bottom = 16384
                    mView.MeasureString szv, prTemp, prTemp

                    If fNeedGap Then _
                        Y = Y + EM_TITLE_TEXT

                    mLabels = mLabels + 1
                    ReDim Preserve mLabel(mLabels)
                    With mLabel(mLabels)
                        .Text = szv
                        Set .Frame = new_BRect(EM_MARGIN, Y, mView.Width - EM_MARGIN - 1, Y + prTemp.Height - 1)

                    End With

                    fNeedGap = True
                    Y = Y + prTemp.Height

                End If
            Loop

        End With


        ' /* size the view */

        Y = Y + EM_MARGIN
        .SizeTo .Width, MAX(Y, cyNotification)

    End With

End Sub

Private Sub uDraw()
Const RX = 20

    With mView
        .Clear
        .EnableSmoothing True
        .TextMode = MFX_TEXT_ANTIALIAS

        ' /* background */

        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(mRed, 0, 0)
            .FillRoundRect .Bounds, RX, RX

        Else
            .SetHighColour mBkCol
            .FillRoundRect .Bounds, RX, RX

        End If

        ' /* inner gradient fill */

        .SetHighColour rgba(0, 0, 0, 90)
        .SetLowColour rgba(0, 0, 0, 0)
        .FillRoundRect new_BRect(.Bounds.Left, .Bounds.Top, .Bounds.Right, MIN(.Bounds.Bottom, 60 - 1)), RX, RX, MFX_VERT_GRADIENT

        If mTintCol <> 0 Then _
            .DrawScaledImage g_CreateMarker(mTintCol), new_BPoint(.Width - 30, 0), , 200

        ' /* outer white border */

        .SetHighColour rgba(255, 255, 255, 240)
        .StrokeRoundRect .Bounds, RX, RX, 2

        ' /* inner dark border */

        .SetHighColour rgba(0, 0, 0, 130)
        .StrokeRoundRect .Bounds.InsetByCopy(1, 1), RX, RX

        ' /* title */

        If mTitle <> "" Then
'            .SetHighColour rgba(255, 0, 0)
'            .StrokeRect mTitleRect

            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size"))
            .SetHighColour rgba(255, 255, 255, 190)

'            dw = Val(mStyle.ReadSetting("fg_alpha"))
'            .SetHighColour rgba(255, 255, 255, Fix((dw / 100) * 255))

            .DrawString mTitle, mTitleRect, MFX_ALIGN_H_CENTER

        End If

        ' /* text */

        If mText <> "" Then
'            .SetHighColour rgba(255, 255, 0)
'            .StrokeRect mTextRect
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            .SetHighColour rgba(255, 255, 255, 190)
            .DrawString mText, mTextRect, MFX_ALIGN_H_CENTER

        End If

        ' /* icon */

        If Not (mIcon Is Nothing) Then _
            .DrawScaledImage mIcon, mIconPos, new_BPoint(mIconSize, mIconSize)

        ' /* meter */

Dim pr As BRect

        If mPercent > -1# Then
            .SetHighColour rgba(0, 0, 0, 60)
            .FillRoundRect mMeterRect, 4, 4

            .SetHighColour rgba(255, 255, 255, 90)
            Set pr = mMeterRect.InsetByCopy(1, 1)
            pr.Right = pr.Left + ((mPercent / 100) * pr.Width) - 1
            If pr.Width > 3 Then _
                .FillRoundRect pr, 4, 4

        End If

        ' /* label-* */

Dim i As Long

        If mLabels > 0 Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size") - 1
            .SetHighColour rgba(255, 255, 255, 190)

            For i = 1 To mLabels
                Set pr = mLabel(i).Frame.Duplicate
                If (mTitle = "") And (mText = "") And (mMeterRect Is Nothing) Then _
                    pr.OffsetTo pr.Left, Fix((.Height - pr.Height) / 2)

                .DrawString mLabel(i).Text, pr, MFX_ALIGN_H_CENTER

            Next i

        End If

    End With

End Sub

