VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMetroInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TMetroInstance.cls
'/
'/  Description:    Represents a single instance of a Metro style notification
'/
'/  © 2010 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Const EM_TITLE_TEXT = 8
Const EM_MARGIN = 16
Const RX = 32

Dim mStyle As TStyle

Dim mInfo As notification_info
Dim mIconSize As Long
Dim mIcon As mfxBitmap
Dim mView As mfxView

Dim mBackground As Long
Dim mForeground As Long
Dim mBorder As Long

Dim mIconOnly As Boolean
Dim mIsNumeric As Boolean
Dim mTitleRect As BRect
Dim mTextRect As BRect

Implements IStyleInstance
Implements IMeterStyle

Private Sub IMeterStyle_SetStyle(Style As TStyle)

    Set mStyle = Style

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

    If (mView Is Nothing) Then
        Set mView = New mfxView
        mView.TextMode = MFX_TEXT_ANTIALIAS

    End If

    LSet mInfo = NotificationInfo

    ' /* clear cached icon */

    Set mIcon = Nothing
    mIconSize = 0

Dim pi As MImage

    ' /* cache icon */

    Set pi = load_image_obj(NotificationInfo.Icon)
    If is_valid_image(pi) Then
        Set mIcon = style_MakeSquareImage(pi)
        mIconSize = mIcon.Height

    End If

    ' /* pre-configuration */

    mIconOnly = ((mInfo.Title = "") And (mInfo.Text = "")) 'Or (InStr(mInfo.Scheme, "icon only") > 0))
    mIsNumeric = (g_IsPercent(mInfo.Text) And (Val(mStyle.ReadSetting("show_graphically"))))

    uCalcSize
    
    ' /* now we have the final notification height, calculate the size of the icon */

    If is_valid_image(mIcon) Then _
        mIconSize = MIN(mIcon.Height, mView.Height + (2 * EM_MARGIN))
    
    uDraw

End Sub

Private Sub uDraw()

    ' /* draw the background */

    With mView
        .Clear
        .EnableSmoothing True

        ' /* background */

'        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'            ' /* priority notifications always get a red background */
'            .SetHighColour rgba(223, 0, 0)
'
'        Else

'        End If

        mBackground = Val(mStyle.ReadSetting("background"))
        mForeground = Val(mStyle.ReadSetting("foreground"))
        mBorder = Val(mStyle.ReadSetting("border"))

        If mInfo.Scheme <> "custom" Then
            mBackground = mStyle.GetSchemeColour(mInfo.Scheme, "background")
            mForeground = mStyle.GetSchemeColour(mInfo.Scheme, "foreground")
            mBorder = mStyle.GetSchemeColour(mInfo.Scheme, "border")

        End If

        .SetHighColour mBackground
        .FillRoundRect .Bounds, RX, RX

    End With

    ' /* draw the foreground */

    uDrawForeground

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

    ' /* measure title and text to see how big the notification needs to be */

    With mView

        ' /* create the base bounding rect - this is used to get the title and text height
        '    based on a fixed width */

        Set pr = new_BRect(0, 0, 250 - 1, cyNotification - 1)
        pr.InsetBy EM_MARGIN, EM_MARGIN

        ' /* measure title */

        If mInfo.Title <> "" Then
            .SetFont mStyle.ReadSetting("title_font"), mStyle.ReadSetting("title_size"), True
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Title, prTemp, prTemp
            Set mTitleRect = new_BRect(0, 0, prTemp.Width + 0, prTemp.Height + 0)

        End If

        ' /* measure text */

        If (mInfo.Text <> "") And (Not mIsNumeric) Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size"), True
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Text, prTemp, prTemp
            Set mTextRect = new_BRect(0, 0, prTemp.Width + 0, prTemp.Height + 0)

        End If

        ' /* size the view */

        .SizeTo 250, MAX(mTitleRect.Height + mTextRect.Height + (2 * EM_MARGIN) + EM_TITLE_TEXT, cyNotification)

    End With

End Sub

Private Sub uDrawForeground()

    ' /* icon */

    If mIconSize > 0 Then _
        mView.DrawScaledImage mIcon, _
                              new_BPoint(Fix((mView.Width - mIconSize) / 2), Fix((mView.Height - mIconSize) / 2)), _
                              new_BPoint(mIconSize, mIconSize), _
                              96

    If mIconOnly Then _
        Exit Sub

Dim pr As BRect
Dim prLabel As BRect
Dim dw As Single
Const OFFS = 4

    With mView

        ' /* title */

        If mTitleRect.Bottom > 0 Then
            Set pr = mTitleRect.Duplicate
            If (mTextRect.Height = 1) And (Not mIsNumeric) Then
                ' /* no text so centre vertically */
                pr.OffsetBy EM_MARGIN, Fix((.Height - mTitleRect.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN, EM_MARGIN

            End If

            .SetHighColour rgba(0, 0, 0, 48)
            .FillRect new_BRect(3, pr.Top - OFFS, .Width - 3 - 1, pr.Bottom + OFFS)

            .SetHighColour rgba(0, 0, 0, 48)
            .StrokeLine new_BRect(3, pr.Top - OFFS, .Width - 3 - 1, pr.Top - OFFS)
            .StrokeLine new_BRect(3, pr.Bottom + OFFS, .Width - 3 - 1, pr.Bottom + OFFS)

            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
            .SetHighColour mForeground
            .DrawString mInfo.Title, pr

        End If

        ' /* text */

        If mIsNumeric Then

            Set pr = new_BRect(EM_MARGIN, 0, .Width - EM_MARGIN - 1, 24 - 1)
            pr.OffsetBy 0, EM_MARGIN + mTitleRect.Height + EM_TITLE_TEXT + 3
            .SetHighColour rgba(0, 0, 0, 102)
            .FillRoundRect pr, 16, 16

            Set prLabel = pr.Duplicate

            dw = g_SafeLong(mInfo.Text)
            dw = MAX(MIN(dw, 100), 0)
            If dw > 4 Then
                .SetHighColour mBackground
                pr.InsetBy 2, 2
                pr.Right = pr.Left + Fix(pr.Width * (dw / 100))
                .FillRoundRect pr, 16, 16

            End If

            If Val(mStyle.ReadSetting("show_percent")) Then
                .SetHighColour mForeground
                .SetFont "Arial", 9, True
                .DrawString CStr(dw) & "%", prLabel, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

            End If

        ElseIf mTextRect.Bottom > 1 Then
            Set pr = mTextRect.Duplicate
            If mTitleRect.Height = 1 Then
                ' /* no title so centre vertically */
                pr.OffsetBy EM_MARGIN, Fix((.Height - mTextRect.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN, EM_MARGIN + mTitleRect.Height
'                .StrokeLine new_BRect(1, pr.Top - 1, .Width - 1, pr.Top - 1), 3
                pr.OffsetBy 0, EM_TITLE_TEXT

            End If

            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size"), True
            .SetHighColour mForeground
            .DrawString mInfo.Text, pr

        End If

        ' /* border */

        .SetHighColour mBorder
        .StrokeRoundRect .Bounds.InsetByCopy(1, 1), RX, RX, 3

        .SetHighColour rgba(0, 0, 0)
        .StrokeRoundRect .Bounds, RX, RX, 1

    End With

End Sub



'         "Custom|Urban|Blue|Dust|White|Nuclear", _

'        Select Case mInfo.Scheme
'        Case "neo"
'            mBackground = rgba(154, 154, 154)
'            mForeground = rgba(255, 255, 255)
'            mBorder = rgba(244, 112, 32)
'
'        Case "uniform"
'            mBackground = rgba(129, 141, 190)
'            mForeground = rgba(255, 255, 255)
'            mBorder = rgba(82, 98, 160)
'
'        Case "dust"
'            mBackground = rgba(255, 255, 255)
'            mForeground = rgba(198, 57, 106)
'            mBorder = rgba(129, 141, 190)
'
'        Case "pavement"
'            mBackground = rgba(255, 255, 255)
'            mForeground = rgba(156, 156, 156)
'            mBorder = rgba(133, 133, 133)
'
'        Case "nuclear"
'            mBackground = rgba(69, 69, 69)
'            mForeground = rgba(0, 0, 0)
'            mBorder = rgba(242, 225, 13)
'
'        Case "scape"
'            mBackground = rgba(87, 155, 174)
'            mForeground = rgba(254, 189, 72)
'            mBorder = rgba(199, 84, 84)
'
'        Case "metric"
'            mBackground = rgba(217, 237, 255)
'            mForeground = rgba(235, 74, 135)
'            mBorder = rgba(140, 213, 89)
'
'

