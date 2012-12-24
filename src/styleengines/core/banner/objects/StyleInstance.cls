VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public StyleName As String
Private mScheme As String

Private Const MTV_FONT_NAME = "Bauhaus 93"
Private Const TINY_FONT_NAME = "Tahoma"
Private Const TINY_FONT_SIZE = 7

Dim theView As mfxView
Dim mIcon As mfxBitmap

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)

    Debug.Print "ADJUST POSITION: " & x & "," & y

    If StyleName = "music video" Then
        x = -8
        y = g_ScreenHeight() - theView.Height

    Else
        x = g_ScreenWidth() - (theView.Width + 32)
        y = -9

    End If

End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = theView.ConvertToBitmap()

End Function

Private Sub uUpdateMTV(ByVal Title As String, ByVal Text As String)
Dim cy As Long
Dim cxIcon As Long
Dim xOffset As Long
Dim pr As mfxRect
Dim hf As Long

Dim fSoftIcon As Boolean

    fSoftIcon = (mScheme = "1998")

    If Not (mIcon Is Nothing) Then
        cxIcon = MIN(mIcon.Width, 128)
        If Not fSoftIcon Then _
            xOffset = cxIcon + 8

    End If

    Set theView = New mfxView

    With theView
        .SizeTo g_ScreenWidth(), 128 + 16
'        .SmoothingMode = MFX_SMOOTHING_NONE
        .TextMode = MFX_TEXT_ANTIALIAS

        ' /* background */

        If mScheme = "1998" Then
            .SetHighColour rgba(128, 0, 128, 190)

        Else
            .SetHighColour rgba(0, 0, 128, 190)

        End If

        .FillRect .Bounds

        ' /* icon */

        .DrawScaledImage mIcon, create_point(8, 8), create_point(cxIcon, cxIcon), IIf(fSoftIcon, 160, 220)

        ' /* title */

        Set pr = .Bounds.InsetToCopy(8, 8)
        pr.Left = pr.Left + xOffset

        .SetHighColour rgba(255, 255, 255, 220)
        .SetLowColour rgba(0, 0, 0)
        .SetFont MTV_FONT_NAME, 14, True
        .DrawString Title, pr, MFX_SIMPLE_OUTLINE

        ' /* measure the amount of space the title used - graphics.library doesn't
        '    have a way to do this so we must do it manually */

'        hf = gdiplusCreateFont(MTV_FONT_NAME, 14, FontStyleRegular)
'        cy = g_gdipGetStringHeight(theView, Title, .Width - 16, hf) + 8 + 4
'        GdipDeleteFont hf

        ' /* text */

        Set pr = .Bounds.InsetToCopy(8, 8)
        pr.Left = pr.Left + xOffset
        pr.Top = cy

        .SetFont MTV_FONT_NAME, 12
        .DrawString Text, pr, MFX_SIMPLE_OUTLINE


'        .SetFont "Verdana", 7, True
'        cy = cy + (mInfo.MX * 2) + .StringHeight("W") + 3
'        If cy < 64 Then _
'            cy = 64                     ' // minimum allowable size
'
'        ' /* size to the correct height based on the text to display */
'
'        .SizeTo 250, cy

    End With

End Sub

Private Sub uUpdateTiny(ByVal Title As String, ByVal Text As String)
Dim cxIcon As Long
Dim xOffset As Long
Dim pr As mfxRect
Dim hf As Long
Dim cyTitle As Long
Const RX = 20

Dim nColBack As Long
Dim nColOutline As Long
Dim nColText As Long
Dim nColTextOutline As Long

    If Not (mIcon Is Nothing) Then
        cxIcon = MIN(mIcon.Width, 24)
        xOffset = cxIcon + 4

    End If

    Set theView = New mfxView

        ' /* measure the amount of space the text needs - graphics.library doesn't
        '    have a way to do this so we must do it manually */

'    hf = gdiplusCreateFont(TINY_FONT_NAME, TINY_FONT_SIZE, FontStyleBold)

Dim nHeight As Long
Dim nWidth As Long

'    cyTitle = g_gdipGetStringHeight(theView, "A", 0, hf)    ' // height of single character...

    Select Case mScheme
    Case "extra small"
        nHeight = (2 * 2) + (cyTitle * 2) + 2    ' // height of title plus one line of text plus 4px border plus 2px margin
        nHeight = nHeight - 1

'        nWidth = gdiplusGetStringWidth(theView, Title, hf, cyTitle)
'        nWidth = MAX(nWidth, gdiplusGetStringWidth(theView, Text, hf, cyTitle)) + xOffset + 12

    Case "small"
        nWidth = 150
'        nHeight = g_gdipGetStringHeight(theView, Title, nWidth - (xOffset + 4), hf) + 2
'        nHeight = nHeight + g_gdipGetStringHeight(theView, Text, nWidth - (xOffset + 4), hf)
'        nHeight = nHeight + (2 * 2)

    End Select

    ' /* set the colours based on the scheme */

'    Select Case mScheme
'    Case "black"
        nColBack = rgba(0, 0, 0)
        nColText = rgba(255, 255, 255, 220)
        nColTextOutline = rgba(0, 0, 0)
        nColOutline = 0

'    Case "white"
'        nColBack = rgba(255, 255, 255)
'        nColText = rgba(0, 0, 0, 190)
'        nColTextOutline = 0
'        nColOutline = rgba(0, 0, 0, 190)
'
'    Case "yellow"
'        nColBack = rgba(255, 255, 190)
'        nColText = rgba(0, 0, 0, 190)
'        nColTextOutline = 0
'        nColOutline = rgba(0, 0, 0, 190)
'
'    End Select

    With theView
'        .SizeTo nWidth, MAX(24 + 8, cyTitle + cyText + 8)
        .SizeTo nWidth, nHeight

'        .SmoothingMode = MFX_SMOOTHING_ANTIALIAS
        .TextMode = MFX_TEXT_ANTIALIAS

        ' /* background */

        .SetHighColour nColBack
        .FillRoundRect .Bounds, RX, RX
        .FillRect create_rect(0, 0, .Bounds.Right, RX)

'        .SetHighColour rgba(204, 204, 204, 255)
        .SetHighColour nColOutline
'        .StrokeRoundRect .Bounds, RX, RX, 2
'        .SetHighColour rgba(0, 0, 0)
'        .StrokeRoundRect .Bounds, RX, RX

        ' /* icon */

        If Not (mIcon Is Nothing) Then _
            .DrawScaledImage mIcon, create_point(4, 2), create_point(cxIcon, cxIcon), 220

        ' /* title */

'        .SetFont TINY_FONT_NAME, TINY_FONT_SIZE, True
        .SetHighColour nColText
        .SetLowColour nColTextOutline

        Set pr = .Bounds.InsetToCopy(2, 2)
        pr.Left = pr.Left + xOffset
'        gdiplusDrawString theView, Title, pr, nColText, hf, , 4

        ' /* text */

        Set pr = .Bounds.InsetToCopy(2, 2)
        pr.Left = pr.Left + xOffset
        pr.Top = cyTitle + 2
'        gdiplusDrawString theView, Text, pr, nColText, hf, , 4

'        gdiplusDeleteFont hf

    End With

End Sub

Private Sub IStyleInstance_Pulse()
End Sub

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    Set mIcon = load_image_obj(NotificationInfo.Icon)
    mScheme = LCase$(NotificationInfo.Scheme)

    Select Case StyleName
    Case "mtv"
        uUpdateMTV NotificationInfo.Title, NotificationInfo.Text

    Case "tiny"
        uUpdateTiny NotificationInfo.Title, NotificationInfo.Text

    End Select

End Sub



Public Function g_ScreenWidth(Optional ByVal VirtualScreen As Boolean = False) As Long
    g_ScreenWidth = GetSystemMetrics(SM_CXSCREEN)
End Function

Public Function g_ScreenHeight(Optional ByVal VirtualScreen As Boolean = False) As Long
    g_ScreenHeight = GetSystemMetrics(SM_CYSCREEN)
End Function

