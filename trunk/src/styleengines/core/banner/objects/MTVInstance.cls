VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MTVInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Style As TStyle

'Private Const FONT_NAME = "Arial Rounded MT Bold"
'Private Const FONT_SIZE_TITLE = 26
'Private Const FONT_SIZE_TEXT = 12

Const TEXT_MARGIN = 4
Const TEXT_FIX = 2
Const MARGIN = 16
Const H_SPLIT = 500

Dim theView As mfxView
Dim mShowing As Boolean
Dim mYPos As Long

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)

    ' /* on entry:
    '
    '       x, y are set to the position Snarl has determined is best for the notification
    '       Alpha is set to the maximum opacity this notification can reach
    '       Done is set to True
    '
    '   on exit:
    '       In order for Snarl to keep polling AdjustPosition() we must set 'Done' to False
    '       each time until we're done
    '
    ' */

    x = 0

    If mShowing Then

        y = GetSystemMetrics(SM_CYSCREEN) - mYPos

        If mYPos < theView.Height Then
            mYPos = mYPos + 6
            Done = False

        Else
            ' /* done */
            y = GetSystemMetrics(SM_CYSCREEN) - theView.Height
            Done = True

        End If

    Else
        y = GetSystemMetrics(SM_CYSCREEN) - mYPos
        
        If mYPos > 0 Then
            mYPos = mYPos - 6
            Done = False

'        Else
'            mYPos = 0

        End If

    End If

End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = theView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)

    mShowing = Visible

    If Visible Then
        mYPos = 0

    Else
        mYPos = theView.Height

    End If

End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo NotificationInfo.Text

'Dim fAdvanced As Boolean

'    fAdvanced =

Dim pIcon As mfxBitmap
Dim xOffset As Long
Dim cxIcon As Long

    xOffset = MARGIN
    cxIcon = 128 '// IIf(NotificationInfo.Scheme = LCase$(SCHEME_2), 64, 128)
    Set pIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))
    If Not (pIcon Is Nothing) Then _
        xOffset = xOffset + cxIcon + MARGIN


'Dim pb As mfxBitmap
'
'    With New mfxView
'        .EnableSmoothing True
'        .SizeTo 500, 500
'        .SetHighColour rgba(255, 255, 255, 10)
'        .SetHighColour rgba(255, 255, 255, 255)
'        .FillRect .Bounds ', MFX_VERT_GRADIENT
'        Set pb = .ConvertToBitmap
'
'    End With

    Set theView = New mfxView
    With theView
        .SizeTo GetSystemMetrics(SM_CXSCREEN), cxIcon + 16
        .EnableSmoothing False
        .TextMode = MFX_TEXT_ANTIALIAS

    End With

'        ' /* background */
'        .SetHighColour .LowColour
'        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'            .SetHighColour rgba(230, 0, 0, 230)
'
'        Else
'            .SetHighColour rgba(0, 0, 0, 230)
'
'        End If
'        Set pr = new_BRect(0, 0, .Bounds.Right, Fix(.Height / 2) - 1)
'        pr.OffsetBy 0, pr.Height
'        .FillRect pr

'        For i = 0 To .Bounds.Bottom Step 4
'            .StrokeLine new_BRect(0, i, .Bounds.Right, i)
'
'        Next i

        ' /* TESTING TESTING TESTING */
'fAdvanced = True
'        ppd.SetTo "title::Last Train To Trancentral#?value-rating::4/5#?label-artist::The KLF#?label-album::The White Room#?label-duration::3:47#?label-copyright::© 1991 KLF COMMUNICATIONS"
        ' /* TESTING TESTING TESTING */

    If (ppd.Exists("value-rating")) Or (ppd.Exists("label-artist")) Or (ppd.Exists("label-album")) Then
        uDrawAdvanced xOffset, ppd

    Else
        uDrawNormal xOffset, ppd, (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY), pIcon, cxIcon

    End If

End Sub

Private Function uRatingPercent(ByVal Rating As String) As Integer
Dim s() As String

    uRatingPercent = -1

    On Error Resume Next
    
    s = Split(Rating, "/")
    If UBound(s) <> 1 Then _
        Exit Function

Dim a As Single
Dim b As Single

    a = Val(s(0))
    b = Val(s(1))

    If b <= 0 Then _
        Exit Function

    a = a / b
    Debug.Print "rating = " & a

    If a > 1 Then _
        Exit Function

    uRatingPercent = a * 100

End Function

Private Sub uDrawNormal(ByVal xOffset As Long, ByRef Content As BPackedData, ByVal IsPriority As Boolean, ByRef Icon As MImage, ByVal IconSize As Long)
Dim pb As MImage
Dim pr As BRect
Dim cy As Long
Dim i As Long
Dim j As Long

    With theView

        ' /* calc background size */

        cy = 0
        i = -1
        If Content.Exists("value-percent") Then
            i = g_SafeLong(Content.ValueOf("value-percent"))
            If (i < 0) And (i > 100) Then _
                i = -1

        End If

        ' /* title */
        If Content.ValueOf("title") <> "" Then
            .SetFont Style.ReadSetting("title_font"), Style.ReadSetting("title_size"), (Style.ReadSetting("title_bold") = "1"), (Style.ReadSetting("title_italic") = "1")
            cy = cy + .StringHeight("A")
            
        End If

        ' /* text */
        If Content.ValueOf("text") <> "" Then
            .SetFont Style.ReadSetting("text_font"), Style.ReadSetting("text_size"), (Style.ReadSetting("text_bold") = "1"), (Style.ReadSetting("text_italic") = "1")
            cy = cy + .StringHeight("A")

        End If

        ' /* stars */
        If i > -1 Then _
            cy = cy + 16 + 2


        ' /* --- draw --- */


        ' /* background */
        If IsPriority Then
            .SetHighColour rgba(230, 0, 0, 230)

        Else
            .SetHighColour g_SafeLong(Style.ReadSetting("col-background"))

        End If
        Set pr = new_BRect(0, .Bounds.Bottom - cy, .Bounds.Right, .Bounds.Bottom)
        .FillRect pr

        ' /* icon */
        .DrawScaledImage Icon, new_BPoint(MARGIN, 8), new_BPoint(IconSize, IconSize), 250

        ' /* now work from top down */
        cy = .Bounds.Bottom - cy

        ' /* title */
        If Content.ValueOf("title") <> "" Then
            .SetFont Style.ReadSetting("title_font"), Style.ReadSetting("title_size"), (Style.ReadSetting("title_bold") = "1"), (Style.ReadSetting("title_italic") = "1")
            Set pr = new_BRect(xOffset, 0, .Bounds.Right - MARGIN, .StringHeight("A") - 1).OffsetByCopy(0, cy)
            .SetHighColour rgba(255, 255, 255)
            .SetLowColour rgba(0, 0, 0, 64)
            .DrawString .GetFormattedText(Content.ValueOf("title"), pr.Width), pr, MFX_ALIGN_V_CENTER Or MFX_SIMPLE_OUTLINE
            cy = cy + .StringHeight("A")

        End If

        ' /* text */
        If Content.ValueOf("text") <> "" Then
            cy = cy - 2
            .SetFont Style.ReadSetting("text_font"), Style.ReadSetting("text_size"), (Style.ReadSetting("text_bold") = "1"), (Style.ReadSetting("text_italic") = "1")
            Set pr = new_BRect(xOffset, 0, .Bounds.Right - MARGIN, .StringHeight("A") - 1).OffsetByCopy(0, cy)
            .SetHighColour rgba(0, 204, 254)
            .SetLowColour rgba(0, 0, 0, 64)
            .DrawString .GetFormattedText(Content.ValueOf("text"), pr.Width), pr, MFX_ALIGN_V_CENTER Or MFX_SIMPLE_OUTLINE
            cy = cy + .StringHeight("A") + 2

        End If

        ' /* rating stars */
        If i > 0 Then
            Set pr = new_BRect(0, 0, 16 - 1, 16 - 1).OffsetByCopy(xOffset, cy)
            Set pb = load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "star.png")
            i = Fix(i / 20)
            For j = 1 To i
                .DrawScaledImage pb, pr.TopLeft, pr.Size
                pr.OffsetBy 16 + 1, 0

            Next j

        End If

        ' /* redraw at constant alpha */
        Set pb = .ConvertToBitmap()
        .Clear
        .DrawScaledImage pb, , , 240

    End With

End Sub

Private Sub uDrawAdvanced(ByVal xOffset As Long, ByRef Content As BPackedData)
Dim sz As String
Dim pb As MImage
Dim pr As BRect
Dim cy As Long
Dim i As Long
Dim j As Long

    With theView
        i = -1
        If Content.Exists("value-rating") Then
            i = uRatingPercent(Content.ValueOf("value-rating"))
            Debug.Print "rating % = " & CStr(i)

        ElseIf Content.Exists("value-percent") Then
            i = g_SafeLong(Content.ValueOf("value-percent"))
            If (i < 0) Or (i > 100) Then _
                i = -1

        End If

        ' /* starting point */
        Set pr = new_BRect(xOffset, .Bounds.Bottom - TEXT_MARGIN - 16 - 1, .Bounds.Right - MARGIN - 1, .Bounds.Bottom - TEXT_MARGIN - 1)

        ' /* rating stars */
        If i > 0 Then
            Set pb = load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "star.png")
            i = Fix(i / 20)
            For j = 0 To (i - 1)
                .DrawScaledImage pb, new_BPoint(pr.Left + (17 * j), pr.Top), new_BPoint(16, 16)

            Next j
            pr.OffsetBy 0, -17

        End If

        .SetFont Style.ReadSetting("text_font"), Style.ReadSetting("text_size")
        .SetHighColour rgba(0, 204, 254)

        ' /* copyright */
        If Content.ValueOf("label-copyright") <> "" Then
            Set pr = new_BRect(xOffset, pr.Bottom - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, pr.Bottom - 1)
            pr.OffsetBy 0, -cy
            .DrawString Content.ValueOf("label-copyright"), pr, MFX_ALIGN_TOP
            cy = pr.Height - TEXT_FIX

        End If

        .SetFont Style.ReadSetting("text_font"), Style.ReadSetting("text_size")

        ' /* album */
        If Content.ValueOf("label-album") <> "" Then
            Set pr = new_BRect(xOffset, pr.Bottom - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, pr.Bottom - 1)
            pr.OffsetBy 0, -cy
            .DrawString Content.ValueOf("label-album"), pr, MFX_ALIGN_TOP
            cy = pr.Height - TEXT_FIX

        End If

        .SetFont Style.ReadSetting("text_font"), Style.ReadSetting("text_size"), (Style.ReadSetting("text_bold") = "1")

        ' /* artist */
        If Content.ValueOf("label-artist") <> "" Then
            Set pr = new_BRect(xOffset, pr.Bottom - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, pr.Bottom - 1)
            pr.OffsetBy 0, -cy
            .DrawString Content.ValueOf("label-artist"), pr, MFX_ALIGN_TOP
            cy = pr.Height - TEXT_FIX

        End If

        ' /* title */
        If Content.ValueOf("title") <> "" Then _
            sz = Content.ValueOf("title")

        If Content.ValueOf("label-duration") <> "" Then _
            sz = sz & IIf(sz <> "", " ", "") & "(" & Content.ValueOf("label-duration") & ")"

        If sz <> "" Then
            .SetFont Style.ReadSetting("title_font"), Style.ReadSetting("title_size"), (Style.ReadSetting("title_bold") = "1"), (Style.ReadSetting("title_italic") = "1")
            Set pr = new_BRect(xOffset, pr.Bottom - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, pr.Bottom - 1)
            pr.OffsetBy 0, -cy
            .SetHighColour rgba(255, 255, 255)
            .DrawString sz, pr, MFX_ALIGN_TOP

        End If

    End With

End Sub
