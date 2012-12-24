VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TClockInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TClockInstance.cls
'/
'/  Description:    Represents a single instance of a Clock style notification
'/
'/  © 2010 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Dim mInfo As notification_info
Dim mIcon As mfxBitmap
Dim mView As mfxView

Const t As Double = 57.29577951

Private Type T_TIME
    Hour As Integer
    Minute As Integer
    Second As Integer
    Date As Integer
    Month As Integer
    Year As Integer

End Type

Dim mTime As T_TIME
Dim mStyle As TStyle
Dim mTintCol As Long

Implements IMeterStyle
Implements IStyleInstance

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
Dim szTime As String

    If (mView Is Nothing) Then
        Set mView = New mfxView
        mView.TextMode = MFX_TEXT_ANTIALIAS

    End If

    LSet mInfo = NotificationInfo

    ' /* now V42 so augment replace title and text with real title and text */

    With New BPackedData
        If .SetTo(NotificationInfo.Text) Then
            mInfo.Title = .ValueOf("title")
            mInfo.Text = .ValueOf("text")
            szTime = .ValueOf("value-date-packed")

            mTintCol = 0
            If .Exists("colour-tint") Then _
                mTintCol = g_SafeLong(.ValueOf("colour-tint"))

        End If
    End With

Dim pi As MImage

    ' /* cache icon */

    Set mIcon = Nothing
    Set pi = load_image_obj(NotificationInfo.Icon)
    If is_valid_image(pi) Then _
        Set mIcon = create_bitmap_from_image(pi)

    ' /* validate text field - it must be in the format "YYYYMMDDHHMMSS" */

    If szTime <> "" Then
        uGetPackedTime szTime

    ElseIf mInfo.Text <> "" Then
        uGetPackedTime mInfo.Text

    End If

    uDraw

End Sub

Private Function uGetPackedTime(ByVal PackedTime As String) As Boolean

    If Len(PackedTime) <> 14 Then _
        Exit Function

'        g_Debug "TClockInstance.UpdateContent(): invalid time string '" & mInfo.Text & "'"
'        mInfo.Text = Format$(Now, "YYYYMMDDHHNNSS")
'        g_Debug "converted to now(): " & mInfo.Text
'
'    End If

    With mTime
        .Year = uVal(Left$(PackedTime, 4))
        .Month = uVal(Mid$(PackedTime, 5, 2))
        .Date = uVal(Mid$(PackedTime, 7, 2))
        .Hour = uVal(Mid$(PackedTime, 9, 2))
        .Minute = uVal(Mid$(PackedTime, 11, 2))
        .Second = uVal(Mid$(PackedTime, 13, 2))

    End With

End Function

Private Sub uDraw()

    If (mView Is Nothing) Then
        g_Debug "TClockInstance.uCalcSize(): view not created", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    Select Case LCase$(mInfo.Scheme)
    Case "analog"
        uDrawAnalog

    Case "digital"
        uDrawDigital

    Case Else
        uDrawReminder

    End Select

End Sub

Private Sub uDrawAnalog()
Dim pi As mfxBitmap
Static i As Integer
Static n As Double
Dim pr As BRect

    With mTime
        If .Hour > 12 Then
            .Hour = .Hour - 12

        ElseIf .Hour = 0 Then
            .Hour = 12

        End If

    End With

    Set pi = load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "clock\back.png")
    If is_valid_image(pi) Then
        With mView
            .SizeTo pi.Width, pi.Height
            .Clear
            .DrawScaledImage pi

        End With

    Else
        With mView
            .SizeTo 96, 96
            .Clear
            .EnableSmoothing True
            .SetHighColour rgba(255, 255, 255)
            .FillEllipse .Bounds
            .SetHighColour rgba(0, 0, 0, 0)
            .SetLowColour rgba(0, 0, 0, 32)
            .FillEllipse .Bounds, MFX_VERT_GRADIENT
            .SetHighColour rgba(0, 0, 0)
            .StrokeEllipse .Bounds, 2

            .SetHighColour rgba(0, 0, 0)
    
            For i = 0 To 59
                n = (i * 6) - 90
                If n < 0 Then _
                    n = n + 360
    
                n = n / t
    
                Set pr = new_BRect(Fix(.Bounds.Width / 2) + (Fix(.Bounds.Width / 2) * 0.97 * Cos(n)), _
                                   Fix(.Bounds.Height / 2) + (Fix(.Bounds.Height / 2) * 0.97 * Sin(n)), _
                                   Fix(.Bounds.Width / 2) + (Fix(.Bounds.Width / 2) * 0.97 * Cos(n)), _
                                   Fix(.Bounds.Height / 2) + (Fix(.Bounds.Height / 2) * 0.97 * Sin(n)))
    
                If (i Mod 5) = 0 Then
                    ' /* draw hour ticks */
                    If MIN(.Width, .Height) > 99 Then
                        pr.InsetBy -2, -2
    
                    Else
                        pr.InsetBy -1, -1
    
                    End If
                    .FillEllipse pr
    
                ElseIf (MIN(.Width, .Height) > 179) Then
                    ' /* draw second ticks */
                    pr.InsetBy -1, -1
                    .FillEllipse pr
    
                End If
    
            Next i

        End With

    End If

Dim cx As Long

    If is_valid_image(mIcon) Then
        cx = MIN(mIcon.Width, 48)
        mView.DrawScaledImage mIcon, new_BPoint(0, mView.Height - cx), new_BPoint(cx, cx)

    End If

Dim hi As Long, mi As Long, si As Long
Dim Hd As Double, Md As Double, Sd As Double
Dim hr As Double, Mr As Double, sr As Double
 
    hi = mTime.Hour
    mi = mTime.Minute
    si = mTime.Second

    If hi >= 12 Then _
        hi = hi - 12

    Hd = hi * 30
    Hd = Hd + mi / 2
    Md = mi * 6
    Sd = si * 6

    Hd = Hd - 90
    Md = Md - 90
    Sd = Sd - 90

    If Hd < 0 Then _
        Hd = Hd + 360

    If Md < 0 Then _
        Md = Md + 360

    If Sd < 0 Then _
        Sd = Sd + 360

    hr = Hd / t
    Mr = Md / t
    sr = Sd / t

    With mView


'MsgBox mStyle.GetSchemeSetting(mInfo.Scheme, "hour_ratio")

        .SetHighColour rgba(0, 0, 0)
        .StrokeLine new_BRect(Fix(.Bounds.Width / 2), _
                              Fix(.Bounds.Height / 2), _
                              Fix(.Bounds.Width / 2) + (Fix(.Bounds.Width / 2) * Val(mStyle.GetSchemeSetting(mInfo.Scheme, "hour_ratio")) * Cos(hr)), _
                              Fix(.Bounds.Height / 2) + (Fix(.Bounds.Height / 2) * Val(mStyle.GetSchemeSetting(mInfo.Scheme, "hour_ratio")) * Sin(hr))), _
                              2

        .StrokeLine new_BRect(Fix(.Bounds.Width / 2), _
                              Fix(.Bounds.Height / 2), _
                              Fix(.Bounds.Width / 2) + (Fix(.Bounds.Width / 2) * Val(mStyle.GetSchemeSetting(mInfo.Scheme, "minute_ratio")) * Cos(Mr)), _
                              Fix(.Bounds.Height / 2) + (Fix(.Bounds.Height / 2) * Val(mStyle.GetSchemeSetting(mInfo.Scheme, "minute_ratio")) * Sin(Mr))), _
                              2

        If mTime.Second > -1 Then
            .SetHighColour rgba(255, 0, 0)
            .StrokeLine new_BRect(Fix(.Bounds.Width / 2), _
                                  Fix(.Bounds.Height / 2), _
                                  Fix(.Bounds.Width / 2) + (Fix(.Bounds.Width / 2) * Val(mStyle.GetSchemeSetting(mInfo.Scheme, "second_ratio")) * Cos(sr)), _
                                  Fix(.Bounds.Height / 2) + (Fix(.Bounds.Height / 2) * Val(mStyle.GetSchemeSetting(mInfo.Scheme, "second_ratio")) * Sin(sr)))

        End If

        ' /* draw centre point if no background */

        If Not is_valid_image(pi) Then
            .SetHighColour rgba(0, 0, 0)
            .FillEllipse .Bounds.InsetByCopy(.Bounds.Width / 2 - 6, .Bounds.Height / 2 - 6)

        End If

        ' /* date */

        .SetFont "Arial", 9, True
        .TextMode = MFX_TEXT_ANTIALIAS
        Set pr = new_BRect(0, 0, 17, 13)
        pr.OffsetBy .Width - (pr.Width + 19), Fix(.Height / 2) - Fix(pr.Height / 2)
'        .SetHighColour rgba(255, 0, 0)
'        .StrokeRect pr

        .SetHighColour rgba(0, 0, 0, 220)
        .DrawString CStr(mTime.Date), pr, MFX_ALIGN_RIGHT Or MFX_ALIGN_V_CENTER

        ' /* title */

        If mInfo.Title <> "" Then
            .SetFont "Arial", 8, True
            .TextMode = MFX_TEXT_ANTIALIAS
            Set pr = new_BRect(8, 8, .Width - 8 - 1, 8 + (.StringHeight("A") * 2) + 1)
            .SetHighColour rgba(255, 255, 255)
            .SetLowColour IIf((mInfo.Flags And S_NOTIFICATION_IS_PRIORITY), rgba(255, 0, 0, 240), rgba(0, 0, 0, 240))
            .DrawString mInfo.Title, pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER Or MFX_SIMPLE_OUTLINE

        End If

    End With

End Sub

Private Sub uDrawDigital()
Dim title_height As Long
Dim time_height As Long
Dim date_height As Long
Dim dateFormat As String
Dim pr As BRect

    With mView

        .SizeTo 120, 54

        ' /* this style only utilises the notification title.  It consists of three parts (one of which is optional)
        '
        '       TITLE
        '      -------
        '       TIME
        '      -------
        '      [Date]
        ' */

        If mTime.Date > -1 Then _
            dateFormat = dateFormat & "d"

        If mTime.Month > -1 Then _
            dateFormat = dateFormat & IIf(dateFormat <> "", " ", "") & "mmmm"

        If mTime.Year > -1 Then _
            dateFormat = dateFormat & IIf(dateFormat <> "", " ", "") & "yyyy"


        ' /* measure the title and text sizes */

        .SetFont "Calibri", 9, True

        If mInfo.Title <> "" Then
            .MeasureString mInfo.Title, .Bounds.InsetByCopy(8, 0), pr
            title_height = pr.Height

        End If

        If dateFormat <> "" Then
            .MeasureString dateFormat, .Bounds.InsetByCopy(8, 0), pr
            date_height = pr.Height

        End If

        .SetFont "Calibri", 20
        .MeasureString "00:00", .Bounds.InsetByCopy(8, 0), pr
        time_height = pr.Height + 2

        ' /* size to correct height */

        .SizeTo 120, title_height + time_height + date_height
        .EnableSmoothing True

        .SetHighColour rgba(32, 32, 32, 220)
        .FillRoundRect .Bounds, 12, 12

        .SetHighColour rgba(0, 0, 0)
        .StrokeRoundRect .Bounds, 12, 12

Dim cx As Long

        If is_valid_image(mIcon) Then
            cx = MIN(mIcon.Width, .Height - 4)
            mView.DrawScaledImage mIcon, new_BPoint(Fix((.Width - cx) / 2), Fix((.Height - cx) / 2)), new_BPoint(cx, cx), 128

        End If

        .SetHighColour rgba(255, 255, 255, 240)
        .SetLowColour IIf((mInfo.Flags And S_NOTIFICATION_IS_PRIORITY), rgba(255, 0, 0, 220), rgba(0, 0, 0, 180))


        ' /* time */

        .SetFont "Calibri", 20
        .DrawString Format$(mTime.Hour, "00") & ":" & Format$(mTime.Minute, "00") & IIf(mTime.Second = -1, "", ":" & Format$(mTime.Second, "00")), _
                    .Bounds, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER Or MFX_SIMPLE_OUTLINE

        ' /* date */

        .SetFont "Calibri", 9, True

        If dateFormat <> "" Then _
            .DrawString Format$(DateSerial(mTime.Year, mTime.Month, mTime.Date), dateFormat), _
                        .Bounds.InsetByCopy(0, 4), MFX_ALIGN_H_CENTER Or MFX_ALIGN_BOTTOM Or MFX_SIMPLE_OUTLINE

        ' /* title */

        If mInfo.Title <> "" Then _
            .DrawString mInfo.Title, .Bounds.InsetByCopy(0, 4), MFX_ALIGN_H_CENTER Or MFX_ALIGN_TOP Or MFX_SIMPLE_OUTLINE

        .SetHighColour rgba(255, 255, 255, 160)
        .SetLowColour rgba(255, 255, 255, 30)
        .FillRoundRect new_BRect(0, 0, .Bounds.Right, 24), 12, 12, MFX_VERT_GRADIENT

    End With

End Sub

Private Function uVal(ByVal str As String) As Integer

    If (LCase$(str) = "xx") Or (LCase$(str) = "xxxx") Then
        uVal = -1

    Else
        uVal = Val(str)

    End If

End Function

Private Sub uDrawReminder()
Dim prBounds As BRect
Dim cy As Long
Const RX = 4

    With mView
        ' /* measure */

        If (mTime.Date = 0) And (mTime.Month = 0) And (mInfo.Title = "") Then
            cy = 50

        ElseIf (mTime.Date = 0) And (mTime.Month = 0) Then
            cy = 70

        Else
            cy = 92

        End If

        .SizeTo 280, cy
        .EnableSmoothing True
        .TextMode = MFX_TEXT_ANTIALIAS
        Set prBounds = .Bounds.InsetByCopy(8, 8)

        ' /* day of month */
        If mTime.Date > 0 Then
            .SetFont "Arial Narrow", 32
            prBounds.Left = prBounds.Left + .StringWidth("00") + 4

        End If

        ' /* draw */
        .SetHighColour rgba(255, 255, 255)
        .FillRoundRect .Bounds, RX, RX

'        If mTintCol <> 0 Then
'            Set pb = g_CreateMarker(rgba(0, 0, 0))
'            .DrawScaledImage pb, new_BPoint(.Width - 24 - 1, 0), new_BPoint(pb.Width + 2, pb.Height + 1), 176
'            .DrawScaledImage g_CreateMarker(mTintCol), new_BPoint(.Width - 24, 0), , 244
'
'        End If

        .SetHighColour rgba(0, 0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 16)
        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT

        ' /* priority */
        If mInfo.Flags And S_NOTIFICATION_IS_PRIORITY Then
            .DrawScaledImage load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "clock\reminder-priority.png"), new_BPoint(Fix((.Width - 128) / 2), Fix((.Height - 128) / 2)), new_BPoint(128, 128), 48
            .SetHighColour rgba(255, 0, 0, 96)

        Else
            .SetHighColour rgba(0, 0, 0, 96)

        End If
        .StrokeRoundRect .Bounds, RX, RX


        ' /* day of month */
        If mTime.Date > 0 Then
            .SetHighColour rgba(255, 64, 64, 220)
            .SetFont "Arial Narrow", 32
            .DrawString Format$(mTime.Date, "00"), .Bounds.InsetByCopy(8, 8).OffsetByCopy(-4, -4)

        End If

        ' /* month and year */

Dim prTitle As BRect

        Set prTitle = prBounds.Duplicate

        If (mTime.Month > 0) Then
            .SetHighColour rgba(0, 0, 0, 250)
            .SetFont "Arial", 10, True
            .DrawString Format$(DateSerial(mTime.Year, mTime.Month, mTime.Date), "dddd mmmm yyyy"), prTitle
            prTitle.Top = prTitle.Top + .StringHeight("A") + 4

        End If

        If mInfo.Title <> "" Then
            .SetHighColour rgba(0, 0, 0, 250)
            .SetFont "Arial", 10
            .DrawString .GetFormattedText(mInfo.Title, prTitle.Width), prTitle

        End If

        ' /* text and icon block */

        Set prBounds = .Bounds.InsetByCopy(8, 8)
        prBounds.Top = prBounds.Bottom - 32
        .SetHighColour rgba(158, 158, 158, 176)
        .StrokeRect prBounds

        If mTintCol <> 0 Then
            .EnableSmoothing False
            .SetHighColour rgba(get_red(mTintCol), get_green(mTintCol), get_blue(mTintCol), 72)
            .FillRect prBounds.InsetByCopy(1, 1)
            .EnableSmoothing True

        End If

Dim pb As MImage

        Set pb = load_image_obj(mInfo.Icon)
        If is_valid_image(pb) Then _
            .DrawScaledImage pb, prBounds.InsetByCopy(8, 8).TopLeft, new_BPoint(16, 16)

        prBounds.Left = prBounds.Left + 32
        .SetHighColour rgba(255, 124, 124)
        .StrokeLine new_BRect(prBounds.Left, prBounds.Top + 1, prBounds.Left, prBounds.Bottom - 1)
        .StrokeLine new_BRect(prBounds.Left + 2, prBounds.Top + 1, prBounds.Left + 2, prBounds.Bottom - 1)
        prBounds.Left = prBounds.Left + 4

        If mInfo.Text <> "" Then
            .SetHighColour rgba(0, 0, 0, 245)
            .SetFont "Arial", 9
            .DrawString .GetFormattedText(mInfo.Text, prBounds.InsetByCopy(4, 4).Width), prBounds.InsetByCopy(4, 4)

        End If

        .SetHighColour rgba(0, 0, 0, 180)
        .SetFont "Arial", 7
        .DrawString Format$(TimeSerial(mTime.Hour, mTime.Minute, 0), "h:mm AMPM"), prBounds.InsetByCopy(4, 4), MFX_ALIGN_LEFT Or MFX_ALIGN_BOTTOM

    End With

End Sub


