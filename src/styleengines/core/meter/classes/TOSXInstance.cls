VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TOSXInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    ' /*
    '
    '   TOSXInstance.cls -- Represents a single instance of a OSeXy Style notification
    '
    '   © 2011 full phat products.  Some Rights Reserved.
    '   Licenced as per the Simplified BSD (http://www.opensource.org/licenses/bsd-license.php)
    '
    '   Notes
    '   -----
    '
    '   This style utilises the new (introduced with Snarl R2.4) S_STYLE_V42_CONTENT flag
    '   to allow for more advanced notification content.  Enabling this flag causes Snarl
    '   to send the entire unabridged packed data created by sn41[EZ]Notify() and
    '   sn41[EZ]Update() in notification_info->Text
    '
    ' */

Dim mStyle As TStyle
Dim mInfo As notification_info
Dim mView As mfxView
Dim mIcon As mfxBitmap
Dim mIconSize As Long
Dim mShowing As Boolean

Implements IStyleInstance
Implements IMeterStyle

Private Sub IMeterStyle_SetStyle(Style As TStyle)

    Set mStyle = Style

End Sub

Private Sub IStyleInstance_AdjustPosition(X As Long, y As Long, Alpha As Integer, Done As Boolean)

    X = Fix((g_ScreenWidth() - mView.Width) / 2)
    y = Fix((g_ScreenHeight() - mView.Height) / 2)
    Alpha = IIf(mShowing, 255, 0)
    Done = True

End Sub

Private Function IStyleInstance_GetContent() As melon.MImage
Dim pb As mfxBitmap

    Set pb = mView.ConvertToBitmap()
'    pb.MakeTransparent 120
    Set IStyleInstance_GetContent = pb

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)

    mShowing = Visible

End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    If (mView Is Nothing) Then _
        Set mView = New mfxView

    ' /* notificationinfo->scheme is translated to lowercase by Snarl */

    LSet mInfo = NotificationInfo

    ' /* cache icon */

    mIconSize = 0
    Set mIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))
    If is_valid_image(mIcon) Then _
        mIconSize = 128                 ' // configurable

    ' /* pre-configuration */

Dim ppd As BPackedData
Dim sz As String
Dim n As Long

    Set ppd = New BPackedData
    ppd.SetTo NotificationInfo.Text

    If ppd.Exists("text") Then _
        sz = ppd.ValueOf("text")

    ' /* percent */

    n = -1#                             ' // means we don't need a meter...

    If ppd.Exists("value-percent") Then
        n = g_SafeLong(ppd.ValueOf("value-percent"))
        If (n < 0) Or (n > 100) Then _
            n = -1#

    ElseIf g_IsNumeric(sz) Then
        n = Val(sz)
        If (n < 0) Or (n > 100) Then _
            n = -1#

    End If

    With mView
        .SizeTo 200, 200
        .Clear
        .EnableSmoothing True
        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(230, 0, 0)

        Else
            .SetHighColour g_SafeLong(mStyle.ReadSetting("background"))

        End If
        .FillRoundRect .Bounds

        If mIconSize > 0 Then _
            .DrawScaledImage mIcon, new_BPoint(Fix((.Width - mIconSize) / 2), Fix((.Height - mIconSize) / 2)), new_BPoint(mIconSize, mIconSize)

        ' /* meter */

Dim c As Integer
Dim i As Integer
Dim pr As BRect

        If n > -1 Then
            ' /* draw the meter */
            Select Case mInfo.Scheme
            Case "solid"
                Set pr = .Bounds.InsetByCopy(20, 20)
                pr.Top = pr.Bottom - 8
                .SetHighColour rgba(0, 0, 0, 64)
                .FillRect pr
                .StrokeRect pr
                c = pr.Width * (n / 100)
                pr.Right = pr.Left + c
                .SetHighColour g_SafeLong(mStyle.ReadSetting("foreground"))
                .FillRect pr

            Case Else
                c = Fix(n / 10)
                .EnableSmoothing False
                Set pr = new_BRect(0, 0, 7 - 1, 8 - 1)
                pr.OffsetBy 33, .Bounds.Bottom - 24
                For i = 1 To 10
                    If i > c Then
                        .SetHighColour rgba(0, 0, 0, 48)
    
                    Else
                        .SetHighColour g_SafeLong(mStyle.ReadSetting("foreground"))
    
                    End If
    
                    .FillRect pr
                    .SetHighColour rgba(0, 0, 0, 32)
                    .StrokeRect pr.InsetByCopy(-1, -1)
                    pr.OffsetBy pr.Width + 6, 0
    
                Next i
            End Select

        End If

Dim pbTemp As mfxBitmap
'Dim nAlp As Long

        Set pbTemp = .ConvertToBitmap()
        .Clear
        .EnableSmoothing True
        .DrawScaledImage pbTemp, , , (g_SafeLong(mStyle.ReadSetting("bg_alpha")) / 100) * 255 '// 172

    End With

End Sub


