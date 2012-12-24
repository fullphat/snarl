VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMusicInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TMusicInstance.cls
'/
'/  Description:    Represents a single instance of a Slideshow notification
'/
'/  © 2010 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Const EM_TITLE_TEXT = 3
Const EM_MARGIN = 10
Const EM_LEFT_MARGIN = 48

Dim mStyle As TStyle
Dim mStarIcon As mfxBitmap

Dim mInfo As notification_info
Dim mIcon As mfxBitmap
Dim mView As mfxView

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

    ' /* notifications are a fixed width and height */

    mView.SizeTo 250, 100
    Set mStarIcon = load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & mStyle.Name & "\star.png")
    If Not is_valid_image(mStarIcon) Then _
        Set mStarIcon = create_bitmap(18, 18)

Dim pi As MImage

    ' /* cache icon */

    Set pi = load_image_obj(NotificationInfo.Icon)
'    Set pi = load_image_obj(g_MakePath(App.Path) & "back3.jpg")

    If is_valid_image(pi) Then _
        Set mIcon = style_MakeSquareImage(pi)

    ' /* pre-configuration */

    mIconOnly = ((mInfo.Title = "") And (mInfo.Text = "")) 'Or (InStr(mInfo.Scheme, "icon only") > 0))
    mIsNumeric = (g_IsPercent(mInfo.Text) And (Val(mStyle.ReadSetting("show_graphically"))))
    Set mTitleRect = new_BRect(0, 0, 0, 0)
    Set mTextRect = new_BRect(0, 0, 0, 0)

    uDraw

    If mIconOnly Then _
        Exit Sub

    If mInfo.Scheme = "none" Then
        ' /* do nothing else */

    ElseIf mInfo.Scheme = "top and bottom" Then
        uDrawTopAndBottom

    Else
        uDrawTopOrBottom

    End If

End Sub

Private Sub uDraw()

    ' /* draw the background */

    With mView
        .Clear
        .EnableSmoothing False

        ' /* background */

        .SetHighColour mStyle.GetSchemeColour(mInfo.Scheme, "background")
        .FillRect .Bounds

        ' /* icon */

Dim c As Long

        If Not (mIcon Is Nothing) Then
            c = MIN(mIcon.Width, 250)
            .DrawScaledImage mIcon, _
                             new_BPoint(Fix((.Width - c) / 2), Fix((.Height - c) / 2)), _
                             new_BPoint(c, c)

        End If

'        .SetHighColour rgba(0, 255, 0)
'        .DrawString "+" & mInfo.Scheme & "+", .Bounds, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

    End With

End Sub

Private Sub uDrawTopAndBottom()
Dim txtFlags As MFX_DRAWSTRING_FLAGS
Dim pr As BRect

    With mView

        txtFlags = MFX_ALIGN_LEFT Or MFX_ALIGN_V_CENTER
'        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
''            .SetLowColour rgba(255, 0, 0, 200)
'            txtFlags = txtFlags Or MFX_SIMPLE_OUTLINE
'
'        End If

        ' /* title */
        .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
        .SetHighColour IIf((mInfo.Flags And S_NOTIFICATION_IS_PRIORITY), rgba(255, 0, 0, 160), rgba(0, 0, 0, 160))
        Set pr = new_BRect(0, 0, .Width - 1, .StringHeight("A") + 6)
        .FillRect pr
        .SetHighColour rgba(255, 255, 255)
        .DrawString uFitText(mInfo.Title, pr), pr.InsetByCopy(4, 0), txtFlags

        ' /* text */
        .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size")), True
        .SetHighColour IIf((mInfo.Flags And S_NOTIFICATION_IS_PRIORITY), rgba(255, 0, 0, 160), rgba(0, 0, 0, 160))
        Set pr = new_BRect(0, 0, .Width - 1, (.StringHeight("A") * 2) + 6)
        pr.OffsetBy 0, .Height - pr.Height
        .FillRect pr

        If mIsNumeric Then
            uDrawStars pr.TopLeft.OffsetByCopy(4, Fix((pr.Height - mStarIcon.Height) / 2))

        Else
            pr.InsetBy 4, 2
            .SetHighColour rgba(255, 255, 255, 220)
            .DrawString uFitText(mInfo.Text, pr), pr, txtFlags

        End If

    End With

End Sub

Private Sub uDrawTopOrBottom()
Dim txtFlags As MFX_DRAWSTRING_FLAGS
Dim cyTitle As Long
Dim cyText As Long
Dim pr As BRect

    With mView

        txtFlags = MFX_ALIGN_LEFT Or MFX_ALIGN_TOP

'        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'            .SetLowColour rgba(255, 0, 0, 200)
'            txtFlags = txtFlags Or MFX_SIMPLE_OUTLINE
'
'        End If

        ' /* calc title */
        .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
        cyTitle = .StringHeight("A")
        Set pr = new_BRect(0, 0, .Width - 1, cyTitle)

        ' /* calc text */
        .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size")), True
        cyText = .StringHeight("A")
        pr.ExpandBy 0, cyText * 2

        pr.ExpandBy 0, 6        ' // padding

        If mInfo.Scheme = "bottom" Then _
            pr.OffsetBy 0, .Height - pr.Height

        .SetHighColour IIf((mInfo.Flags And S_NOTIFICATION_IS_PRIORITY), rgba(255, 0, 0, 160), rgba(0, 0, 0, 160))
        .FillRect pr

        pr.InsetBy 4, 3

        ' /* title */
        .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
        .SetHighColour rgba(255, 255, 255)
        .DrawString uFitText(mInfo.Title, new_BRect(pr.Left, pr.Top, pr.Right, pr.Top + cyTitle + 1)), pr, txtFlags
        pr.OffsetBy 0, .StringHeight("A")

        ' /* text */
        If mIsNumeric Then
            uDrawStars new_BPoint(0, 0)

        Else
            .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size")), True
            .SetHighColour rgba(255, 255, 255, 220)
            .DrawString uFitText(mInfo.Text, new_BRect(pr.Left, pr.Top, pr.Right, pr.Top + cyText * 2)), pr, txtFlags

        End If

    End With

End Sub

Private Sub uDrawStars(ByRef Origin As BPoint)

    If (Origin Is Nothing) Or (mStarIcon Is Nothing) Then _
        Exit Sub

Dim i As Integer
Dim c As Long

    c = g_SafeLong(mInfo.Text)
    c = MAX(MIN(c, 100), 0)
    c = Fix(c \ 20)
'    MsgBox mInfo.Text & " == " & c

    If c < 1 Then _
        Exit Sub

    With mView
        For i = 1 To c
            .DrawScaledImage mStarIcon, Origin, , 220
            Origin.OffsetBy mStarIcon.Width, 0

        Next i

    End With

End Sub

Private Function uFitText(ByVal Text As String, ByRef Bounds As BRect) As String

    If (mView Is Nothing) Or (Text = "") Then _
        Exit Function

Dim pr As BRect
Dim bDots As Boolean

    Do
        mView.MeasureString Text & IIf(bDots, "...", ""), Bounds, pr
        If pr.Height < Bounds.Height Then
            uFitText = Text & IIf(bDots, "...", "")
            Exit Function

        End If

        Text = g_SafeLeftStr(Text, Len(Text) - 1)
        bDots = True
        If Text = "" Then _
            Exit Function

    Loop

End Function

