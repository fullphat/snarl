VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfoBarInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Style As TStyle

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const FONT_NAME = "Tahoma"
Private Const FONT_SIZE_TITLE = 10
Private Const FONT_SIZE_TEXT = 10

Private Const FONT_NAME_OSX = "Lucida Grande"
Private Const FONT_SIZE_TITLE_OSX = 10
Private Const FONT_SIZE_TEXT_OSX = 10

Private Const FONT_NAME_GS = "Tahoma"
Private Const FONT_SIZE_TITLE_GS = 9
Private Const FONT_SIZE_TEXT_GS = 9

Dim theView As mfxView
Dim mShowing As Boolean
Dim mPos As Single

Dim mInfo As notification_info
Dim mIcon As mfxBitmap
Dim mPriorityIcon As mfxBitmap
Dim mValue As Long

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)

    ' /* on entry:
    '       x, y are set to the position Snarl has determined is best for the notification
    '       Alpha is set to the maximum opacity this notification can reach
    '       Done is set to True
    '
    '   In order for Snarl to keep polling AdjustPosition we must set 'Done' to False each
    '   time until we're done positioning the notification */

Const STEP_SIZE = 0.25
Const FAST_STEP_SIZE = 2

Dim d As Single

    d = IIf(mInfo.Scheme = "amiga guru", FAST_STEP_SIZE, STEP_SIZE)
    x = 0

    If mShowing Then
        ' /* counter is set to view height so decrease by one on each call */
        If mPos = 0 Then
            y = 0
            Exit Sub

        End If

        mPos = mPos - d
        y = 0 - mPos
        If mPos > 0 Then _
            Done = False        ' // keep going

    Else
        ' /* counter is set to view height so decrease by one on each call */

        If mPos = 0 Then _
            Exit Sub

        mPos = mPos - d
        y = mPos - theView.Height
        If mPos > 0 Then _
            Done = False        ' // keep going

    End If

End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = theView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)

    mShowing = Visible
    mPos = theView.Height

End Sub


Private Sub uDrawAmiga()
Dim xOffset As Long
Dim pr As BRect

    With theView
        .SizeTo GetSystemMetrics(SM_CXSCREEN), 21
        .EnableSmoothing False
        .TextMode = MFX_TEXT_CLEARTYPE

        ' /* background */

        .SetHighColour rgba(255, 255, 255)
        .FillRect .Bounds

        ' /* priority icon (set/unset during IStyleInstance.UpdateContent) */

        .DrawScaledImage mPriorityIcon, new_BPoint(.Width - 22, 0), new_BPoint(20, 20)

        ' /* shadow */

        Set pr = .Bounds.Duplicate
        pr.Top = pr.Bottom
        .SetHighColour rgba(0, 0, 0)
        .StrokeLine pr

        xOffset = 8

        ' /* notification icon */

Dim cxIcon As Long

        If Not (mIcon Is Nothing) Then
            cxIcon = Min(mIcon.Height, (.Height - 2))
            .DrawScaledImage mIcon, new_BPoint(xOffset, Fix((.Height - cxIcon) / 2)), new_BPoint(cxIcon, cxIcon)
            xOffset = xOffset + cxIcon + 4

        End If

        ' /* title */

'        .EnableSmoothing True

        Set pr = .Bounds.Duplicate
        .SetHighColour rgba(0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 16)

        If mInfo.Title <> "" Then
            pr.Left = xOffset
            .SetFont FONT_NAME, FONT_SIZE_TITLE
            .DrawString mInfo.Title, pr, MFX_ALIGN_V_CENTER Or MFX_SIMPLE_OUTLINE
            xOffset = xOffset + .StringWidth(mInfo.Title) + 4

        End If

        ' /* text */

        pr.Left = xOffset
        .SetFont FONT_NAME, FONT_SIZE_TEXT
        .DrawString mInfo.Text & IIf(mValue > -1, "%", ""), pr, MFX_ALIGN_V_CENTER

    End With

End Sub

Private Sub uDrawOSX()
Dim xOffset As Long
Dim pr As BRect

    With theView
        .SizeTo GetSystemMetrics(SM_CXSCREEN), 24
        .Clear
        .EnableSmoothing False
        .TextMode = MFX_TEXT_CLEARTYPE

        ' /* background */

        .SetHighColour rgba(255, 255, 255)
        .FillRect .Bounds

        .SetHighColour rgba(0, 0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 16)
        .FillRect .Bounds, MFX_VERT_GRADIENT

        ' /* priority icon (set/unset during IStyleInstance.UpdateContent) */

        .DrawScaledImage mPriorityIcon, new_BPoint(.Width - (.Height + 4), 0), new_BPoint(.Height, .Height)

        xOffset = 8

        ' /* notification icon */

Dim cxIcon As Long

        If Not (mIcon Is Nothing) Then
            cxIcon = Min(mIcon.Height, (.Height - 2))
            .DrawScaledImage mIcon, new_BPoint(xOffset, Fix((.Height - cxIcon) / 2)), new_BPoint(cxIcon, cxIcon)
            xOffset = xOffset + cxIcon + 4

        End If

        ' /* title */

        Set pr = .Bounds.Duplicate
        .SetHighColour rgba(0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 24)

        If mInfo.Title <> "" Then
            pr.Left = xOffset
            .SetFont FONT_NAME_OSX, FONT_SIZE_TITLE_OSX
            .DrawString mInfo.Title, pr, MFX_ALIGN_V_CENTER Or MFX_SIMPLE_OUTLINE
            xOffset = xOffset + .StringWidth(mInfo.Title) + 4

        End If

        ' /* text */

Dim sz As String

        sz = mInfo.Text
        If mValue > -1 Then _
            sz = sz & "%"

        If mInfo.Text <> "" Then
            pr.Left = xOffset
            If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then _
                pr.Right = pr.Right - .Height - 4

            .SetFont FONT_NAME_OSX, FONT_SIZE_TEXT_OSX
            .DrawString .GetFormattedText(sz, pr.Width), pr, MFX_ALIGN_V_CENTER

        End If

    End With

End Sub

Private Sub uDrawC9()
Dim xOffset As Long
Dim pr As BRect

    With theView
        .SizeTo GetSystemMetrics(SM_CXSCREEN), 28
        .EnableSmoothing False
        .TextMode = MFX_TEXT_CLEARTYPE

        ' /* background */

        .SetHighColour rgba(216, 216, 216)
        .FillRect .Bounds

        .SetHighColour rgba(0, 0, 0)
        .StrokeRect .Bounds

        Set pr = .Bounds.InsetByCopy(1, 1)
        .SetHighColour rgba(253, 253, 253)
        .SetLowColour rgba(167, 167, 167)
        .StrokeFancyRect pr

        ' /* gripper */

        Set pr = new_BRect(0, 0, 14 - 1, .Height - 1)
        .SetHighColour rgba(0, 128, 192)
        .FillRect pr

        .SetHighColour rgba(0, 0, 0)
        .StrokeRect pr

        pr.InsetBy 1, 1
        .SetHighColour rgba(0, 172, 255)
        .SetLowColour rgba(0, 86, 128)
        .StrokeFancyRect pr

        Set pr = new_BRect(0, 0, 6 - 1, 6 - 1)
        pr.OffsetBy 5, 5
        .SetHighColour rgba(0, 172, 255)
        .StrokeRect pr

        pr.OffsetBy -2, -2
        .SetHighColour rgba(0, 86, 128)
        .StrokeRect pr
        
        pr.OffsetBy 1, 1
        .SetHighColour rgba(0, 0, 0)
        .StrokeRect pr

        uDrawC9Gripper 5, 13
        uDrawC9Gripper 8, 13
        uDrawC9Gripper 5, 16
        uDrawC9Gripper 8, 16

        ' /* priority icon (set/unset during IStyleInstance.UpdateContent) */

        .DrawScaledImage mPriorityIcon, new_BPoint(.Width - (.Height + 4), 0), new_BPoint(.Height, .Height)

        xOffset = 14 + 4

        ' /* notification icon */

'Dim cxIcon As Long

        If Not (mIcon Is Nothing) Then
'            cxIcon = MIN(mIcon.Height, (.Height - 4))
            .DrawScaledImage mIcon, new_BPoint(xOffset, 0), new_BPoint(.Height, .Height)
            xOffset = xOffset + .Height + 4

        End If

        ' /* title */

Dim sz As String
Dim cx As Long

        If mInfo.Title <> "" Then
            Set pr = .Bounds.Duplicate
            pr.Left = xOffset

            ' /* set up font and calc size of title for tray */

            .SetFont FONT_NAME, FONT_SIZE_TITLE
            cx = .StringWidth(mInfo.Title) + 8
            pr.Right = pr.Left + cx
            pr.InsetBy 0, 2

            ' /* tray */

            .SetHighColour rgba(0, 0, 0, 32)            ' // inset tray back
            .FillRect pr

            .SetHighColour rgba(0, 0, 0, 72)
            .SetLowColour rgba(255, 255, 255, 72)
            .StrokeFancyRect pr                         ' // inset shadow and shine

            .SetHighColour rgba(0, 0, 0, 99)
            .StrokeRect pr.InsetByCopy(1, 1)            ' // inner shadow

            .SetHighColour rgba(0, 0, 0)
            .SetLowColour rgba(0, 0, 0, 16)
            .DrawString mInfo.Title, pr, MFX_ALIGN_V_CENTER Or MFX_ALIGN_H_CENTER Or MFX_SIMPLE_OUTLINE
            xOffset = xOffset + cx + 4

        End If

        ' /* text */
        
        sz = mInfo.Text
        If mValue > -1 Then _
            sz = sz & "%"

        If mInfo.Text <> "" Then
            .SetHighColour rgba(0, 0, 0)
            Set pr = .Bounds.Duplicate
            pr.Left = xOffset
            If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then _
                pr.Right = pr.Right - .Height - 4

            .SetFont FONT_NAME, FONT_SIZE_TEXT
            sz = .GetFormattedText(sz, pr.Width)
            .DrawString sz, pr, MFX_ALIGN_V_CENTER

        End If

    End With

End Sub

Private Sub uDrawC9Gripper(ByVal x As Long, ByVal y As Long)
Dim pr As BRect

    Set pr = new_BRect(x, y, x + 2 - 1, y + 2 - 1)
    With theView
        .SetHighColour rgba(0, 172, 255)
        .StrokeRect pr

        pr.OffsetBy -1, -1
        .SetHighColour rgba(0, 86, 128)
        .StrokeRect pr

        .SetHighColour rgba(0, 0, 0)
        .SetPixel new_BPoint(x, y)

    End With

End Sub

Private Sub uDrawGS()
Dim xOffset As Long
Dim pr As BRect

    With theView
        .SizeTo GetSystemMetrics(SM_CXSCREEN), 30
        .EnableSmoothing False
        .TextMode = MFX_TEXT_CLEARTYPE

        ' /* background */

'"Bar Highlight"="132,134,132"
'"Bar Lowlight"="189,190,189"
'"Bar High Edge"="25,25,25"
'"Bar Low Edge"="0,0,0"
'"Bar Background"="50,50,50"
'"Bar Inset Background"="75,75,75"
'"Bar Selected Background"="100,100,100"
'"Bar Inset Text"="255,255,255"
'"Bar Text"="255,255,255"
'"Bar Disabled Text"="150,150,150"
'"Bar Selected Text"="200,200,200"
'"Bar Flashing Background"="255,255,255"
'"Gripper"="0,130,198"
'"Gripper Highlight"="0,178,255"
'"Gripper Lowlight"="0,77,115"

        .SetHighColour rgba(213, 207, 202)
        .FillRect .Bounds                           ' // background

'        .SetHighColour rgba(213, 207, 202)
'        .SetLowColour rgba(120, 120, 120)
'        uDrawGSFancyRect .Bounds

        .SetHighColour rgba(255, 255, 255)
        .SetLowColour rgba(163, 163, 163)
        uDrawGSFancyRect .Bounds                    ' // outer fancy rect

        .SetHighColour rgba(240, 240, 240)
        uDrawGSFancyRect .Bounds.InsetByCopy(1, 1)  ' // outer fancy rect pt 2

        ' /* gripper */

        Set pr = new_BRect(0, 0, 8 - 1, 7 - 1)
        pr.OffsetBy 3, 3

Dim i As Integer

        ' /* gripper - single 'close' button */

        .SetHighColour rgba(255, 255, 255)

        For i = 1 To 3
            uDrawGSFancyRect pr
            .SetHighColour rgba(240, 240, 240)
            pr.InsetBy 1, 1

        Next i

        Set pr = new_BRect(0, 0, 4 - 1, 4 - 1)
        pr.OffsetBy 3, 12

        For i = 1 To 3
            uDrawGSFancyRect pr
            uDrawGSFancyRect pr.OffsetByCopy(4, 0)
            pr.OffsetBy 0, 5

        Next i

        xOffset = 10 + 6

        ' /* notification icon */

'Dim cxIcon As Long
Dim sz As String

        If Not (mIcon Is Nothing) Then
'            cxIcon = MIN(mIcon.Height, (.Height - 4))
            .DrawScaledImage mIcon, new_BPoint(xOffset, 0), new_BPoint(.Height, .Height)
            xOffset = xOffset + .Height + 4

        End If

        ' /* priority icon (set/unset during IStyleInstance.UpdateContent) */

        .DrawScaledImage mPriorityIcon, new_BPoint(.Width - (.Height + 4), 0), new_BPoint(.Height, .Height)

        ' /* title */

        Set pr = .Bounds.Duplicate
        pr.Right = pr.Right - 4
        .SetHighColour rgba(0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 24)

        If mInfo.Title <> "" Then
            pr.Left = xOffset
            .SetFont FONT_NAME_GS, FONT_SIZE_TITLE_GS
            .DrawString mInfo.Title, pr, MFX_ALIGN_V_CENTER Or MFX_SIMPLE_OUTLINE
            xOffset = xOffset + .StringWidth(mInfo.Title) + 4

        End If

        ' /* text */

        sz = mInfo.Text
        If mValue > -1 Then _
            sz = sz & "%"

        If mInfo.Text <> "" Then
            pr.Left = xOffset
            If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then _
                pr.Right = pr.Right - .Height - 4

            .SetFont FONT_NAME_GS, FONT_SIZE_TEXT_GS
            .DrawString .GetFormattedText(sz, pr.Width), pr, MFX_ALIGN_V_CENTER

        End If

    End With

End Sub

Private Sub uDrawGSFancyRect(ByRef aRect As BRect)

    ' /* draws a 1px by 1px fancy rect but in a geOShell / Windows 95 style */

    With theView
        .StrokeLine new_BRect(aRect.Left, aRect.Top, aRect.Left, aRect.Bottom - 1)                      ' // left
        .StrokeLine new_BRect(aRect.Left, aRect.Top, aRect.Right, aRect.Top)                            ' // top
        .StrokeLine new_BRect(aRect.Right, aRect.Top + 1, aRect.Right, aRect.Bottom), , MFX_SOLID_LOW   ' // right
        .StrokeLine new_BRect(aRect.Left, aRect.Bottom, aRect.Right, aRect.Bottom), , MFX_SOLID_LOW     ' // bottom

    End With

End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Dim i As Long

'
'    fSoftIcon = (NotificationInfo.Scheme = SCHEME_1)
'
'            cxIcon = MIN(pIcon.Height, 96)
'
'        Else
'            cxIcon = MIN(pIcon.Height, 128)
'
'        End If
'
'        xOffset = cxIcon + 8
'
'    End If


'Debug.Print g_HexStr(NotificationInfo.Flags)

    mValue = -1
    LSet mInfo = NotificationInfo

    With New BPackedData
        .SetTo NotificationInfo.Text
        mInfo.Title = .ValueOf("title")
        mInfo.Text = .ValueOf("text")

        i = -1
        If .Exists("value-percent") Then
            i = g_SafeLong(.ValueOf("value-percent"))

        ElseIf g_IsNumeric(mInfo.Text) Then
            i = Val(mInfo.Text)
            mInfo.Text = ""

        End If

        If (i >= 0) And (i <= 100) Then _
            mValue = i

    End With

    Set mIcon = style_MakeSquareImage(load_image_obj(mInfo.Icon))
    If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
        Set mPriorityIcon = load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "priority.png")

    Else
        Set mPriorityIcon = Nothing

    End If

    Set theView = New mfxView

    Select Case LCase$(mInfo.Scheme)
    Case "amiga"
        uDrawAmiga

    Case "os x"
        uDrawOSX

    Case "geoshell"
        uDrawGS

    Case "cloud:9ine"
        uDrawC9

    Case "amiga guru"
        uDrawAmigaGuru

    Case Else
        Debug.Print LCase$(mInfo.Scheme)

    End Select

End Sub

Private Sub uDrawAmigaGuru()
Dim cxIcon As Long
Dim pr As BRect

    With theView
        .SizeTo GetSystemMetrics(SM_CXSCREEN), 78
        .EnableSmoothing False
        .TextMode = MFX_TEXT_CLEARTYPE

        ' /* background */
        .SetHighColour rgba(0, 0, 0)
        .FillRect .Bounds

        ' /* icon */
        If Not (mIcon Is Nothing) Then
            cxIcon = Min(mIcon.Height, 128)
            Set pr = new_BRect(0, 0, cxIcon - 1, cxIcon - 1).OffsetByCopy(0, Fix(.Height - cxIcon) / 2)

            If (mInfo.Title = "") And (mInfo.Text = "") And (mValue = -1) Then
                ' /* single, centered icon */
                pr.OffsetBy Fix((.Width - cxIcon) / 2), 0

            Else
                ' /* offset */
                pr.OffsetBy -Fix(cxIcon / 4), 0

            End If

            .DrawScaledImage mIcon, pr.TopLeft, pr.Size, 180, -45

        End If

        ' /* border */
        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(255, 0, 0)
            .SetLowColour rgba(255, 0, 0, 32)

        ElseIf g_IsPressed(VK_CONTROL) Then
            .SetHighColour rgba(0, 255, 0)
            .SetLowColour rgba(0, 255, 0, 32)

        Else
            .SetHighColour rgba(255, 192, 0)
            .SetLowColour rgba(255, 192, 0, 32)

        End If
        .StrokeRect .Bounds, 8

Dim sz As String

        .SetFont "Lucida Console", 12
        .SetLowColour rgba(0, 0, 0, 172)

        If mInfo.Text <> "" Then
            If mValue > -1 Then _
                mInfo.Title = mInfo.Title & " #" & g_HexStr(mValue) & "." & g_HexStr(Year(Now), 4) & g_HexStr(Month(Now), 2) & g_HexStr(Day(Now), 2)

            ' /* title */
            .DrawString mInfo.Title, .Bounds.InsetByCopy(8, 16), MFX_ALIGN_H_CENTER Or MFX_ALIGN_TOP Or MFX_SIMPLE_OUTLINE
            ' /* text */
            .DrawString .GetFormattedText(mInfo.Text, .Bounds.InsetByCopy(8, 16).OffsetByCopy(0, 28).Width), .Bounds.InsetByCopy(8, 16).OffsetByCopy(0, 28), MFX_ALIGN_H_CENTER Or MFX_ALIGN_TOP Or MFX_SIMPLE_OUTLINE

        Else
            ' /* title */
            .DrawString mInfo.Title, .Bounds.InsetByCopy(8, 16), MFX_ALIGN_H_CENTER Or MFX_ALIGN_TOP Or MFX_SIMPLE_OUTLINE
            ' /* text */
            If mValue > -1 Then
                sz = "#" & g_HexStr(mValue) & "." & g_HexStr(Year(Now), 4) & g_HexStr(Month(Now), 2) & g_HexStr(Day(Now), 2)
                .DrawString sz, .Bounds.InsetByCopy(8, 16).OffsetByCopy(0, 28), MFX_ALIGN_H_CENTER Or MFX_ALIGN_TOP Or MFX_SIMPLE_OUTLINE

            End If
        End If
    End With

End Sub

