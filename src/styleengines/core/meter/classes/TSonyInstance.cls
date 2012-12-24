VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSonyInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const EM_TITLE_TEXT = 0
Const EM_MARGIN = 10

Dim mStyle As TStyle
Dim mInfo As notification_info
Dim mIconOnly As Boolean
Dim mIconSize As Long
Dim mIcon As mfxBitmap
Dim mView As mfxView

Dim mTitleRect As BRect
Dim mTextRect As BRect
Dim mValue As Long
Dim mTint As Long

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

Private Sub IStyleInstance_AdjustPosition(X As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
'Static nStep As Integer
'Const STEP_SIZE = 16
'
'    If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'
'        If nStep = 0 Then _
'            nStep = STEP_SIZE
'
'        mRed = mRed + nStep
'        If mRed > 254 Then
'            nStep = -STEP_SIZE
'
'        ElseIf mRed < 1 Then
'            nStep = STEP_SIZE
'
'        End If
'
'        uDraw
'
'        IStyleInstance_Pulse = True
'
'    End If

End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    LSet mInfo = NotificationInfo
'    LSet mSettings = gSettings.Sony

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
    mIconSize = 0
    Set mIcon = Nothing
    Set pi = load_image_obj(NotificationInfo.Icon)
    If is_valid_image(pi) Then
        Set mIcon = style_MakeSquareImage(pi)
        mIconSize = Min(mIcon.Height, IIf(InStr(mInfo.Scheme, "big icon"), 128, 48))

    End If

    ' /* pre-configuration */
    mIconOnly = ((mInfo.Title = "") And (mInfo.Text = ""))
    If mIconOnly Then
        ' /* icon-only notifications are a fixed width and height */
        mView.SizeTo 250, 80

    Else
        ' /* calculate notifcation size and internal measurements */
        uCalcSize

    End If

    uDraw

End Sub

Private Sub uDraw()
Dim m As Long

    ' /* draw the background */

    With mView
        .Clear
        .EnableSmoothing False

        ' /* background */
        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(255, 0, 0, 232)

        Else
            .SetHighColour IIf(InStr(mInfo.Scheme, "psp"), g_GetPSPColour(), g_GetPS3Colour())

        End If
        .FillRect .Bounds

        ' /* for the PS3 darken background based on time of day */
        If InStr(mInfo.Scheme, "ps3") Then
            If gDbgMinute = -1 Then
                ' /* get time of day in minutes */
                m = (Hour(Now) * 60) + Minute(Now)

            Else
                Debug.Print "using debug setting"
                m = gDbgMinute

            End If

            Debug.Print "minute: " & m

            If m > 1080 Then
                ' /* evening (after 6pm) */
                .SetHighColour rgba(0, 0, 0, 255)
                .SetLowColour rgba(0, 0, 0, m * 0.15)

            ElseIf m < 360 Then
                ' /* early morning (before 6am) */
                .SetHighColour rgba(0, 0, 0, 255)
                .SetLowColour rgba(0, 0, 0, 255 - (m * 0.15))

            Else
                ' /* daytime (0600-1800) */
                m = Abs(720 - m)                        ' // brightness strength is based on distance from midday
                .SetHighColour rgba(0, 0, 0, m * 0.7)
                .SetLowColour rgba(0, 0, 0, 0)

            End If

            .FillRect .Bounds, MFX_VERT_GRADIENT

        End If

        ' /* light effect */
'        .SetHighColour rgba(255, 255, 255, 100)
'        .SetLowColour rgba(0, 0, 0, 0)
'        .StrokeFancyRoundRect .Bounds.InsetByCopy(2, 2), 4, 4

Dim pr As BRect
Dim c As Long

        ' /* bokeh effect */
        .EnableSmoothing True
        For m = 1 To 32
            Randomize Timer
            c = (Rnd * 72)
            Set pr = new_BRect(0, 0, c - 1, c - 1).OffsetByCopy(Rnd * (.Width * 2) - (.Width / 2), Rnd * (.Height * 2) - (.Height / 2))
            .SetHighColour rgba(255, 255, 255, (Rnd * 40) + 10)
            .StrokeEllipse pr
            .SetLowColour .HighColour
            .SetHighColour g_SetRGB32Alpha(.HighColour, 2)
            .FillEllipse pr, MFX_CENTER_GRADIENT

        Next m

Dim pt As BPoint

        ' /* icon */
        If mIconSize > 0 Then
            If mIconOnly Then
                ' /* centre of the notification */
                Set pt = new_BPoint(Fix((.Width - mIconSize) / 2), Fix((.Height - mIconSize) / 2))

            ElseIf InStr(mInfo.Scheme, "big icon") Then
                Set pt = new_BPoint(0, Fix((.Height - mIconSize) / 2))

            Else
                Set pt = new_BPoint(EM_MARGIN, EM_MARGIN)

            End If

            .DrawScaledImage mIcon, pt, new_BPoint(mIconSize, mIconSize), IIf(InStr(mInfo.Scheme, "big icon"), 170, 240)

        End If

        If InStr(mInfo.Scheme, "ps3") Then
            ' /* PS3 schemes get reflection effect */
            .EnableSmoothing True
            .SetHighColour rgba(255, 255, 255, 63)
            .FillEllipse new_BRect(-300, -30, .Bounds.Right + 90 - 1, .Bounds.Top + 47)
            .SetHighColour rgba(255, 255, 255, 102)
            .StrokeEllipse new_BRect(-300, -30, .Bounds.Right + 90 - 1, .Bounds.Top + 47)
            .EnableSmoothing False

        End If

        ' /* edge */
        .EnableSmoothing False
        .SetHighColour rgba(0, 0, 0, 220)
        .StrokeRect .Bounds

    End With

    ' /* draw the foreground */

    If Not mIconOnly Then _
        uDrawForeground

End Sub

Private Sub uCalcSize()

    If (mView Is Nothing) Then
        g_Debug "TSonyInstance.uCalcSize(): view not created", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim prTemp As BRect
Dim pr As BRect
Dim cy As Long

    cy = 2 * EM_MARGIN

    ' /* measure title and text to see how big the notification needs to be */

    With mView
        .SizeTo 250, 900
        ' /* create the base bounding rect - this is used to get the title and text height
        '    based on a fixed width */
        Set pr = new_BRect(0, 0, .Bounds.Right, .Bounds.Bottom)
        pr.InsetBy EM_MARGIN, EM_MARGIN
        ' /* if we have an icon then shrink accordingly */
        If (mIconSize > 0) And (mIconSize < 128) Then _
            pr.Right = pr.Right - (mIconSize + EM_MARGIN)

        ' /* measure title */
        If mInfo.Title <> "" Then
            .SetFont mStyle.ReadSetting("title_font"), mStyle.ReadSetting("title_size"), True
            .MeasureString mInfo.Title, pr.Duplicate, prTemp
            Set mTitleRect = new_BRect(0, 0, prTemp.Width + 2, prTemp.Height + 2)
            cy = cy + mTitleRect.Height

        End If

        ' /* measure text */
        If mInfo.Text <> "" Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size"), True
            .MeasureString mInfo.Text, pr.Duplicate, prTemp
            Set mTextRect = new_BRect(0, 0, prTemp.Width + 2, prTemp.Height + 2)
            cy = cy + mTextRect.Height

        End If

        ' /* meter */
        If mValue > -1 Then _
            cy = cy + 10

        ' /* size the view */
        .SizeTo 250, Max(cy, 80)

    End With

End Sub

Private Sub uDrawForeground()
Dim xOffset As Long
Dim pr As BRect

    If mIconSize > 0 Then _
        xOffset = IIf(mIconSize = 128, 0, 48) + EM_MARGIN

    With mView
        ' /* title */
        If NOTNULL(mTitleRect) Then
            Set pr = mTitleRect.Duplicate
            If ISNULL(mTextRect) Then
                ' /* no text so centre vertically */
                pr.OffsetBy EM_MARGIN + xOffset, Fix((.Height - mTitleRect.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN + xOffset, EM_MARGIN

            End If

            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
            .SetHighColour rgba(0, 0, 0, 102)
            .DrawString mInfo.Title, pr.OffsetByCopy(2, 2)                  '// shadow text
            .SetHighColour rgba(255, 255, 255)
            .SetLowColour rgba(96, 96, 96, 127)
            .DrawString mInfo.Title, pr, MFX_SIMPLE_OUTLINE                 '// actual text

        End If

        ' /* text */
        If NOTNULL(mTextRect) Then
            Set pr = mTextRect.Duplicate
            If ISNULL(mTitleRect) Then
                ' /* no title so centre vertically */
                pr.OffsetBy EM_MARGIN + xOffset, Fix((.Height - mTextRect.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN + xOffset, EM_MARGIN + mTitleRect.Height

            End If

            .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size"))
            .SetHighColour rgba(0, 0, 0, 102)
            .DrawString mInfo.Text, pr.OffsetByCopy(2, 2)                   '// shadow text
            .SetHighColour rgba(255, 255, 255)
            .SetLowColour rgba(96, 96, 96, 127)
            .DrawString mInfo.Text, pr, MFX_SIMPLE_OUTLINE                  '// actual text

        End If

        ' /* meter */
        If mValue > -1 Then
            Set pr = .Bounds.InsetByCopy(EM_MARGIN, EM_MARGIN)
            pr.Top = pr.Bottom - 9
            pr.Left = pr.Left + xOffset
            .EnableSmoothing True
            .SetHighColour rgba(0, 0, 0, 160)
            .FillRoundRect pr, 6, 6
            pr.InsetBy 1, 1
            pr.Right = pr.Left + (pr.Width * (mValue / 100))
            .SetHighColour rgba(255, 255, 255, 200)
            .FillRoundRect pr, 4, 4

        End If
    End With

End Sub

