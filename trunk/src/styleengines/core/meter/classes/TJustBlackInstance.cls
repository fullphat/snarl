VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TJustBlackInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TJustBlackInstance.cls
'/
'/  Description:    Represents a single instance of a Just Black style notification
'/
'/  © 2010 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Const EM_TITLE_TEXT = 0
Const EM_MARGIN = 8
Const EM_METER = EM_MARGIN / 2
Const RX = 20

Dim mStyle As TStyle

Dim mInfo As notification_info
Dim mIconSize As Long
Dim mIcon As mfxBitmap
Dim mView As mfxView

Dim mBackground As Long
Dim mTintCol As Long

Dim mIconOnly As Boolean
Dim mIsNumeric As Boolean
Dim mIsBar As Boolean           ' // only applies if mIsNumeric is TRUE.  FALSE represents a graph.

Private Enum E_COLOUR_TYPE
    E_MONO
    E_COLOUR
    E_SPECTRUM

End Enum

Dim mColType As E_COLOUR_TYPE

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

'    If mData.Exists("colour-tint") Then
'        mTintCol = g_SafeLong(mData.ValueOf("colour-tint"))
'
'    Else
'        mTintCol = 0
'
'    End If


    ' /* pre-configuration */

Dim szMeterType As String

    mIconOnly = ((mInfo.Title = "") And (mInfo.Text = "")) 'Or (InStr(mInfo.Scheme, "icon only") > 0))
    mIsNumeric = (g_IsPercent(mInfo.Text) And (Val(mStyle.ReadSetting("show_graphically"))))

    szMeterType = LCase$(mStyle.MeterType(Val(mStyle.ReadSetting("meter_type"))))
    mIsBar = CBool(InStr(szMeterType, "bar"))

    If InStr(szMeterType, "coloured") Then
        mColType = E_COLOUR

    ElseIf InStr(szMeterType, "spectrum") Then
        mColType = E_SPECTRUM

    Else
        mColType = E_MONO

    End If

    uCalcSize
    
    ' /* now we have the final notification height, calculate the size of the icon */

    If is_valid_image(mIcon) Then _
        mIconSize = MIN(mIcon.Height, IIf(InStr(mInfo.Scheme, "big icon"), 128, 48))

    uDraw

End Sub

Private Sub uDraw()

    ' /* draw the background */

    With mView
        .Clear
        .EnableSmoothing True

        ' /* background */

        .SetHighColour Val(mStyle.ReadSetting("background"))
        .FillRoundRect .Bounds, RX, RX

        ' /* priority notifications always get a red background */

        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(223, 0, 0)
            .StrokeRoundRect .Bounds, RX, RX, 3

        End If

    End With

    ' /* draw the foreground */

    uDrawForeground

End Sub

Private Sub uCalcSize()

    If (mView Is Nothing) Then
        g_Debug "TJustBlackInstance.uCalcSize(): view not created", LEMON_LEVEL_CRITICAL
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

        ' /* create the base bounding rect - this is used to get the title and text height based on a fixed width */

        Set pr = new_BRect(0, 0, 250 - 1, cyNotification - 1)
        pr.InsetBy EM_MARGIN, EM_MARGIN
        pr.Left = pr.Left + IIf(mInfo.Scheme = "big icon (whole)", 0, 48) + EM_MARGIN

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
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Text, prTemp, prTemp
            Set mTextRect = new_BRect(0, 0, prTemp.Width + 0, prTemp.Height + 0)

        ElseIf mIsNumeric Then
            cy = IIf(mIsBar, 6, 38) + EM_METER

        End If

        ' /* size the view */

        .SizeTo 250, MAX(mTitleRect.Height + mTextRect.Height + (2 * EM_MARGIN) + EM_TITLE_TEXT + cy, cyNotification)

    End With

End Sub

Private Sub uDrawForeground()
Dim pt As BPoint

    ' /* icon */

    If mIconSize > 0 Then
        If mIconOnly Then
            ' /* center icon horizontally and vertically */
            Set pt = new_BPoint(Fix((mView.Width - mIconSize) / 2), Fix((mView.Height - mIconSize) / 2))

        Else
            ' /* center icon vertically and offset off-notification if using the 'big icon' scheme */
            Set pt = new_BPoint(EM_MARGIN, Fix((mView.Height - mIconSize) / 2))
            
            If mInfo.Scheme = "big icon (whole)" Then
                pt.OffsetBy -EM_MARGIN, 0

            ElseIf mInfo.Scheme = "big icon" Then
                pt.OffsetBy -Fix(mIconSize / 2), 0

            End If

        End If

        mView.DrawScaledImage mIcon, pt, new_BPoint(mIconSize, mIconSize)

    End If

    If mIconOnly Then _
        Exit Sub

Dim pr As BRect
Dim prLabel As BRect
Dim dw As Single
Dim xOff As Long

    With mView

        xOff = EM_MARGIN
        If mIconSize > 0 Then _
            xOff = xOff + IIf(mInfo.Scheme = "big icon (whole)", 0, 48) + EM_MARGIN

        ' /* title */

        If mTitleRect.Bottom > 0 Then
            Set pr = mTitleRect.Duplicate
            If (mTextRect.Height = 1) And (Not mIsNumeric) Then
                ' /* no text so centre vertically */
                pr.OffsetBy xOff, Fix((.Height - mTitleRect.Height) / 2)

            Else
                pr.OffsetBy xOff, EM_MARGIN

            End If

            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
            .SetHighColour rgba(255, 255, 255)
            .DrawString mInfo.Title, pr

        End If

        ' /* text */

        If mIsNumeric Then

            dw = g_SafeLong(mInfo.Text)
            dw = MAX(MIN(dw, 100), 0)

            ' /* bar or graph? */

            If mIsBar Then
                Set pr = new_BRect(EM_MARGIN, 0, .Width - EM_MARGIN - 1, 6 - 1)
                If (mTitleRect.Height = 1) Or (mTitleRect.Bottom + EM_METER + EM_MARGIN < Fix(.Height / 2)) Then
                    ' /* no title or title is smaller than the centre of the notification so center vertically */
                    pr.OffsetBy 0, Fix((.Height - pr.Height) / 2)

                Else
                    pr.OffsetBy 0, mTitleRect.Height + EM_MARGIN + EM_METER

                End If

                If mIconSize > 0 Then _
                    pr.Left = pr.Left + 48 + EM_MARGIN

                .EnableSmoothing False
                .SetHighColour rgba(100, 100, 100)
                .FillRect pr

                Set prLabel = pr.Duplicate

                ' /* draw coloured area */

                pr.Right = pr.Left + Fix(pr.Width * (dw / 100)) - 1

                If pr.Width > 1 Then
                    If mColType = E_COLOUR Then
                        ' /* solid colour */
                        .SetHighColour Val(mStyle.ReadSetting("bar_colour"))

                    ElseIf mColType = E_SPECTRUM Then

                        Select Case dw
                        Case Is >= 80
                            .SetHighColour rgba(0, 255, 0)
        
                        Case Is >= 60
                            .SetHighColour rgba(128, 255, 0)
        
                        Case Is >= 40
                            .SetHighColour rgba(255, 255, 0)
        
                        Case Is >= 20
                            .SetHighColour rgba(255, 128, 0)
        
                        Case Else
                            .SetHighColour rgba(255, 0, 0)

                        End Select

                    Else
                        .SetHighColour rgba(199, 197, 198)

                    End If

                    .FillRect pr

                    ' /* bas relief */

                    .SetHighColour rgba(255, 255, 255, 62)
                    .SetLowColour rgba(0, 0, 0, 62)
                    .StrokeFancyRect pr

                End If

                ' /* darken the lower part of the bar */

                .SetHighColour rgba(0, 0, 0, 0)
                .SetLowColour rgba(0, 0, 0, 64)
                .FillRect prLabel, MFX_VERT_GRADIENT    ' // darken lower half with an alpha shade

                ' /* round the corners */

                .SetHighColour rgba(0, 0, 0, 102)
                .SetPixel prLabel.TopLeft
                .SetPixel prLabel.TopRight
                .SetPixel prLabel.BottomLeft
                .SetPixel prLabel.BottomRight

                If Val(mStyle.ReadSetting("show_percent")) Then
                    .SetHighColour rgba(255, 255, 255)
                    .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size")) - 1, True
                    prLabel.OffsetBy 0, prLabel.Height + 1
                    prLabel.Bottom = prLabel.Top + .StringHeight("A")
                    .DrawString CStr(dw) & "%", prLabel, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

                End If

            Else
                ' /* meter type must be a graph */
                uDrawGraph dw

            End If


        ElseIf mTextRect.Bottom > 1 Then
            Set pr = mTextRect.Duplicate
            If mTitleRect.Height = 1 Then
                ' /* no title so centre vertically */
                pr.OffsetBy xOff, Fix((.Height - mTextRect.Height) / 2)

            Else
                pr.OffsetBy xOff, EM_MARGIN + mTitleRect.Height + EM_TITLE_TEXT

            End If

            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            .SetHighColour rgba(255, 255, 255)
            .DrawString mInfo.Text, pr

        End If

'                .SetHighColour rgba(255, 255, 255)
'                .SetFont "Arial", 9, True
'                .DrawString mstyle.MeterType(Val(mstyle.ReadSetting("meter_type")))

    End With

End Sub

Private Sub uDrawGraph(ByVal Value As Long)
Dim prOrigin As BRect
Dim pr As BRect
Dim cx As Long
Dim n As Long
Dim i As Long
Dim dw As Long

Const BAR_WIDTH = 10
Const BAR_GAP = 4

    With mView

        If Val(mStyle.ReadSetting("show_percent")) Then
           Set pr = .Bounds
            pr.InsetBy EM_MARGIN, EM_MARGIN
            If mTitleRect.Bottom > 1 Then _
                pr.OffsetBy 0, mTitleRect.Height

            .SetHighColour rgba(255, 255, 255)
            pr.OffsetBy 48 + EM_MARGIN, 0
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            .DrawString CStr(Value) & "%", pr

        End If

        ' /* 10 bars each with a gap */

        cx = (10 * BAR_WIDTH) + (9 * BAR_GAP)

        ' /* figure out where the graph should go */

        Set pr = new_BRect(0, 0, cx - 1, 38 - 1)            ' // this is the overall size of the graph

'        Debug.Print mTitleRect.Bottom & " / " & mTitleRect.Height

        If (mIconSize = 0) And (mTitleRect.Bottom = 0) Then
            ' /* if there's no icon and no title then center horizontally and vertically */
            pr.OffsetBy Fix((.Width - pr.Width) / 2), Fix((.Height - pr.Height) / 2)
            
        ElseIf mTitleRect.Bottom = 0 Then
            ' /* icon but no title */
            
            Debug.Print "icon but no title"

            pr.OffsetBy 48 + EM_MARGIN + EM_MARGIN, Fix((.Height - pr.Height) / 2)

        Else
            ' /* icon and title *
            pr.OffsetBy 48 + EM_MARGIN + EM_MARGIN, mTitleRect.Height + EM_MARGIN

        End If

        ' /* offset it from the vertical centre so we can fit the title in */
        
'        If mInfo.Title <> "" Then _
            pr.OffsetBy 0, 4

        ' /* draw the individual bar backgrounds */

        Set pr = new_BRect(pr.Left, pr.Top + 36, pr.Left + 10 - 1, pr.Bottom - 1)
        Set prOrigin = pr.Duplicate

        If Value > 0 Then _
            Value = Fix(Value / 10)

        For i = 1 To 10
            ' /* bar dark outline */
            .EnableSmoothing True
            .SetHighColour rgba(28, 28, 28, 102)
            .FillRoundRect pr.InsetByCopy(-2, -2), 2, 2
            .EnableSmoothing False

            ' /* bar background */
            If i <= Value Then
                If mColType = E_COLOUR Then
                    .SetHighColour Val(mStyle.ReadSetting("bar_colour"))

                ElseIf mColType = E_SPECTRUM Then
                    .SetHighColour uGetSpectrumColour(i)

                Else
                    ' /* assume monochrome */
                    .SetHighColour rgba(199, 197, 198)

                End If

                .FillRect pr

                ' /* shade bar using a black alpha diagonal gradient */

                .SetHighColour rgba(0, 0, 0, 0)
                .SetLowColour rgba(0, 0, 0, 102)
                .FillRect pr, MFX_DIAG_GRADIENT

                ' /* round the corners */
    
                .SetHighColour rgba(0, 0, 0, 102)
                .SetPixel pr.TopLeft
                .SetPixel pr.TopRight
                .SetPixel pr.BottomLeft
                .SetPixel pr.BottomRight

            End If

            pr.OffsetBy BAR_WIDTH + BAR_GAP, 0
            pr.Top = pr.Top - 4

        Next i

    End With

End Sub

Private Function uGetSpectrumColour(ByVal Index As Long)

    Select Case Val(mStyle.ReadSetting("spectrum_type"))
    Case 2
        Select Case Index
        Case Is >= 9:       uGetSpectrumColour = rgba(255, 0, 255)
        Case Is >= 7:       uGetSpectrumColour = rgba(255, 128, 255)
        Case Is >= 5:       uGetSpectrumColour = rgba(255, 255, 255)
        Case Is >= 3:       uGetSpectrumColour = rgba(128, 255, 255)
        Case Else:          uGetSpectrumColour = rgba(0, 255, 255)
        End Select

    Case Else
        Select Case Index
        Case Is >= 9:       uGetSpectrumColour = rgba(0, 255, 0)
        Case Is >= 7:       uGetSpectrumColour = rgba(128, 255, 0)
        Case Is >= 5:       uGetSpectrumColour = rgba(255, 255, 0)
        Case Is >= 3:       uGetSpectrumColour = rgba(255, 128, 0)
        Case Else:          uGetSpectrumColour = rgba(255, 0, 0)
        End Select

    End Select

End Function
