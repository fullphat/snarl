VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNotificationCentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TNotificationCentre.cls
    '/
    '/  Description:    Implements the Notification Centre window
    '/
    '/  © 2012 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Const CLASS_NAME = "w>SnarlNotificationCentre"
Const CXMARGIN = 8
Const FONT_SIZE = 9
Const CYTITLE = 20

Dim mhWnd As Long
Dim mView As mfxView
Dim mScrollView As mfxView
Dim mItems As BTagList
Dim mSelected As String
Dim mOffset As Long

Dim mCloseButton As mfxBitmap
Dim mCloseHeader As String

Implements BWndProcSink

Private Sub Class_Initialize()

    Set mItems = new_BTagList()
    Set mCloseButton = load_image_obj(g_MakePath(App.Path) & "etc\icons\nc-header-close.png")

    Set mScrollView = New mfxView
    With mScrollView
        .TextMode = MFX_TEXT_ANTIALIAS
    
    End With

    Set mView = New mfxView
    With mView
        .SizeTo 500, 300

    End With

    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me, , , WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

Dim pt As POINTAPI

    pt.x = Fix((g_ScreenWidth - mView.Width) / 2)
    pt.y = 0

    apply_view_to_window mView, mhWnd, , pt.x, pt.y

End Sub

Private Sub Class_Terminate()

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Sub AddItem(ByRef App As TApp, ByVal Title As String, ByVal Text As String, ByRef Icon As MImage, ByVal Token As Long)
Dim ph As TNCHeader

    If Not mItems.Find(App.Signature, ph) Then
        Set ph = New TNCHeader
        ph.SetTo App
        mItems.Add ph

    End If

    ph.Add Title, Text, Icon, Token

    uArrange
    uRedraw

End Sub

Public Sub Show()

    If Me.IsVisible Then _
        Exit Sub

Dim pn As TNotification
Dim i As Long

    ' /* get missed items */

    With g_NotificationRoster.MissedList
        If .CountItems Then
            For i = .CountItems To 1 Step -1
                Set pn = .TagAt(i)
                AddItem pn.Info.ClassObj.App, pn.Info.Title, pn.Info.Text, load_image_obj(g_TranslateIconPath(pn.Info.IconPath, "")), pn.Info.Token
            
            Next i
        End If
    End With

    ' /* rethink and redraw */

    uArrange
    uRedraw

    apply_view_to_window mView, mhWnd, 0
    g_ShowWindow mhWnd, True, True
    
'Dim i As Integer
    
    For i = 0 To g_PercentToAlpha(g_SafeLong(g_ConfigGet("nc-opacity-percent"))) Step 16
        apply_view_to_window mView, mhWnd, i
        Sleep 1
        DoEvents
        
    Next i

End Sub

Public Sub Hide()

    If Not Me.IsVisible Then _
        Exit Sub

Dim i As Integer

    For i = g_PercentToAlpha(g_SafeLong(g_ConfigGet("nc-opacity-percent"))) To 0 Step -24
        apply_view_to_window mView, mhWnd, i
        Sleep 1
        DoEvents
        
    Next i

    g_ShowWindow mhWnd, False

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static fCloseButton As Boolean
Static phClose As TNCHeader
Static ppClose As TPopUpItem
Dim pn As TNotification
Dim pp As TPopUpItem
Dim ph As TNCHeader

    Select Case uMsg
    Case WM_LBUTTONDOWN
        If uIsInCloseRect(new_BPointFromInt32(lParam)) Then
            fCloseButton = True

        ElseIf uHitTestHeader(new_BPointFromInt32(lParam), ph) Then
            ' /* inside a header */
            If ph.InCloseButton(new_BPointFromInt32(lParam).OffsetByCopy(0, -mOffset - CYTITLE)) Then _
                Set phClose = ph                ' // inside the header's close button

        ElseIf (uHitTestItem(new_BPointFromInt32(lParam), pp)) And (GET_X_LPARAM(lParam) > (mView.Width - CYTITLE)) Then
            Set ppClose = pp

        End If

    Case WM_LBUTTONUP
        If (fCloseButton) And (uIsInCloseRect(new_BPointFromInt32(lParam))) Then
            ' /* the main NC close button was pressed... */
            Me.Hide

        ElseIf NOTNULL(phClose) Then
            ' /* a header close button was pressed... */
            If phClose.InCloseButton(new_BPointFromInt32(lParam).OffsetByCopy(0, -mOffset - CYTITLE)) Then
                phClose.RemoveAll
                mItems.Remove mItems.IndexOf(phClose.App.Signature)
                If mItems.CountItems = 0 Then
                    Me.Hide

                Else
                    uArrange
                    uRedraw

                End If
            End If
            Set phClose = Nothing

        ElseIf NOTNULL(ppClose) Then
            ' /* mark as seen button was pressed... */
            If (ppClose.Frame.Contains(new_BPointFromInt32(lParam).OffsetByCopy(0, -mOffset - CYTITLE))) And (GET_X_LPARAM(lParam) > (mView.Width - CYTITLE)) Then _
                uRemove ppClose.Token

            Set ppClose = Nothing

        End If
        fCloseButton = False

    Case WM_LBUTTONDBLCLK
        If GET_Y_LPARAM(lParam) <= CYTITLE Then
            ' /* must be at top of NC */
            Debug.Print "scroll to top"
            mOffset = 0
            uRedraw

        ElseIf uHitTestHeader(new_BPointFromInt32(lParam), ph) Then
            ph.Expanded = Not ph.Expanded
            uArrange
            uRedraw

        ElseIf uHitTestItem(new_BPointFromInt32(lParam), pp) Then
            If g_NotificationRoster.MissedList.Find(CStr(pp.Token), pn) Then
                pn.Replay
                uRemove pp.Token

            End If

        End If

    Case WM_MOUSEMOVE
        If ISNULL(phClose) Then
            If uHitTestHeader(new_BPointFromInt32(lParam), ph) Then
                If mCloseHeader <> ph.App.Signature Then
                    mCloseHeader = ph.App.Signature
                    uRedraw

                End If

            ElseIf mCloseHeader <> "" Then
                mCloseHeader = ""
                uRedraw

            End If
        End If

    Case WM_MOUSEWHEEL
        If wParam < 0 Then
            ' /* scroll up */
            If mScrollView.Height > (mView.Height - 22) Then
                mOffset = mOffset - 16
                If (mOffset + mScrollView.Height) < (mView.Height - 22) Then _
                    mOffset = (mView.Height - 22) - mScrollView.Height

                uRedraw
            End If

        Else
            ' /* scroll down */
            mOffset = mOffset + 16
            If mOffset > 0 Then _
                mOffset = 0

            uRedraw
            
        End If

    Case WM_CHAR
        If wParam = 27 Then _
            Me.Hide

    Case &H401
        Me.Hide

    End Select

End Function

Private Sub uRedraw()
Dim pp As TPopUpItem
Dim ph As TNCHeader
Dim pr As BRect
Const RX = 8

    With mScrollView
        .Clear
        mItems.Rewind
        Do While mItems.GetNextTag(ph) = B_OK
            ' /* draw header */
            .SetHighColour rgba(0, 0, 0, 64)
            .FillRect ph.Frame
            .StrokeLine new_BRect(ph.Frame.Left, ph.Frame.Top, ph.Frame.Right, ph.Frame.Top)
            .StrokeLine new_BRect(ph.Frame.Left, ph.Frame.Bottom, ph.Frame.Right, ph.Frame.Bottom)
            .DrawScaledImage ph.App.CachedIcon, ph.Frame.TopLeft.OffsetByCopy(2, 3), new_BPoint(16, 16)
            .SetFont frmAbout.Font.Name, FONT_SIZE, True
            .SetHighColour g_SafeLong(g_ConfigGet("nc-col-text"))
            .DrawString ph.App.NameEx, ph.Frame.InsetByCopy(14, 0).OffsetByCopy(9, 0), MFX_ALIGN_LEFT Or MFX_ALIGN_V_CENTER

            ' /* if mouse pointer is over header, draw close gadget */
            If ph.App.Signature = mCloseHeader Then _
                .DrawScaledImage mCloseButton, ph.Frame.TopRight.OffsetByCopy(-19, 3)

            If ph.Expanded Then
                ' /* draw items */
                ph.Notifications.Rewind
                Do While ph.Notifications.GetNextTag(pp) = B_OK
                    Set pr = pp.Frame.Duplicate
                    pr.Right = pr.Right - 24
'                    If NOTNULL(pp.Icon) Then
'                        .DrawScaledImage pp.Icon, pr.TopLeft.OffsetByCopy(0, Fix((pr.Height - 16) / 2)), new_BPoint(16, 16)
'                        pr.Left = pr.Left + 20
'
'                    End If

                    If pp.Label <> "" Then
                        .SetFont frmAbout.Font.Name, FONT_SIZE, True
                        If pp.Data <> "" Then
                            ' /* both */
                            .DrawString .GetFormattedText(g_RemoveLineFeeds(pp.Label), pr.Width), pr.OffsetByCopy(0, 2), MFX_ALIGN_LEFT Or MFX_ALIGN_TOP
                            .SetFont frmAbout.Font.Name, FONT_SIZE
                            .DrawString .GetFormattedText(g_RemoveLineFeeds(pp.Data), pr.Width), pr.OffsetByCopy(0, .StringHeight("A") + 4), MFX_ALIGN_LEFT Or MFX_ALIGN_TOP
                
                        Else
                            .DrawString .GetFormattedText(g_RemoveLineFeeds(pp.Label), pr.Width), pr, MFX_ALIGN_LEFT Or MFX_ALIGN_V_CENTER

                        End If
                
                    ElseIf pp.Data <> "" Then
                        .SetFont frmAbout.Font.Name, FONT_SIZE
                        .DrawString .GetFormattedText(g_RemoveLineFeeds(pp.Data), pr.Width), pr, MFX_ALIGN_LEFT Or MFX_ALIGN_V_CENTER
                                        
                    End If

                    Set pr = new_BRect(0, 0, 12 - 1, 12 - 1)
                    .DrawScaledImage mCloseButton, pr.OffsetByCopy(pp.Frame.Right - pr.Width - 2, pp.Frame.Top + Fix((pp.Frame.Height - pr.Height) / 2)).TopLeft, new_BPoint(pr.Width, pr.Height)

'                    .SetHighColour rgba(255, 255, 255)
'                    .FillEllipse pr.OffsetByCopy(pp.Frame.Right - pr.Width - 2, pp.Frame.Top + Fix((pp.Frame.Height - pr.Height) / 2))
'                    .SetHighColour rgba(0, 0, 0, 180)
'                    .StrokeEllipse pr.OffsetByCopy(pp.Frame.Right - pr.Width - 2, pp.Frame.Top + Fix((pp.Frame.Height - pr.Height) / 2))

                Loop
            End If
        Loop
    End With

    With mView
        .Clear
        .EnableSmoothing True
        .TextMode = MFX_TEXT_ANTIALIAS
        .SetHighColour g_SafeLong(g_ConfigGet("nc-col-background"))
        .FillRoundRect .Bounds, RX, RX
        .SetHighColour rgba(0, 0, 0, 156)
        .StrokeRoundRect .Bounds, RX, RX
        .SetHighColour g_SafeLong(g_ConfigGet("nc-col-text"))
        .SetFont frmAbout.Font.Name, FONT_SIZE, True
        .DrawString "Notification Centre", .Bounds.InsetByCopy(4, 1), MFX_ALIGN_TOP Or MFX_ALIGN_H_CENTER
        .DrawScaledImage mCloseButton, .Bounds.TopRight.OffsetByCopy(-18, 2)

'        .DrawScaledImage mScrollView.ConvertToBitmap, new_BPoint(1, mOffset)
        .DrawImage mScrollView.ConvertToBitmap(), new_BRect(1, CYTITLE, mScrollView.Width, Min(.Height - 22, mScrollView.Height)), _
                                                  new_BRect(0, -mOffset, mScrollView.Width, Min(mScrollView.Height, .Height - 22))

    End With

    apply_view_to_window mView, mhWnd, g_PercentToAlpha(g_SafeLong(g_ConfigGet("nc-opacity-percent")))

End Sub

Private Function uHitTestHeader(ByRef Point As BPoint, ByRef Header As TNCHeader) As Boolean
Dim ph As TNCHeader

    Point.OffsetBy 0, -mOffset - CYTITLE

    With mItems
        .Rewind
        Do While .GetNextTag(ph) = B_OK
            If ph.Frame.Contains(Point) Then
                Set Header = ph
                uHitTestHeader = True
                Exit Do

            End If
        Loop
    End With

End Function

Private Function uHitTestItem(ByRef Point As BPoint, ByRef Item As TPopUpItem) As Boolean
Dim pp As TPopUpItem
Dim ph As TNCHeader

    Point.OffsetBy 0, -mOffset - CYTITLE

    With mItems
        .Rewind
        Do While .GetNextTag(ph) = B_OK
            If ph.Expanded Then
                With ph.Notifications
                    .Rewind
                    Do While .GetNextTag(pp) = B_OK
                        If pp.Frame.Contains(Point) Then
                            Set Item = pp
                            uHitTestItem = True
                            Exit Do

                        End If
                    Loop
                End With
            End If
        Loop
    End With

End Function

Private Sub uArrange()
Dim pp As TPopUpItem
Dim ph As TNCHeader
Dim pr As BRect
Dim cy As Long

'    mView.SetFont g_ConfigGet("nc-font-typeface"), FONT_SIZE, True
    mView.SetFont frmAbout.Font.Name, FONT_SIZE, True

    Set pr = mView.Bounds '.InsetByCopy(1, CXMARGIN)
    mItems.Rewind
    Do While mItems.GetNextTag(ph) = B_OK
        pr.Bottom = pr.Top + 22 - 1
        Set ph.Frame = pr.Duplicate
        pr.OffsetBy 0, pr.Height
        If ph.Expanded Then
            ph.Notifications.Rewind
            Do While ph.Notifications.GetNextTag(pp) = B_OK
                cy = 0
                If pp.Label <> "" Then _
                    cy = cy + mView.StringHeight("A") + 4

                If pp.Data <> "" Then _
                    cy = cy + mView.StringHeight("A") + 4

                cy = Max(cy, 26)
                pr.Bottom = pr.Top + cy - 1
                Set pp.Frame = pr.InsetByCopy(4, 0)
                pr.OffsetBy 0, pr.Height

            Loop
        End If
    Loop

    mScrollView.SizeTo mView.Width - 2, pr.Top

End Sub

Public Function IsVisible() As Boolean

    IsVisible = (IsWindowVisible(mhWnd) <> 0)

End Function

Private Sub uRemove(ByVal Token As Long)
Dim ph As TNCHeader
Dim i As Long

    ' /* remove from roster */
    g_NotificationRoster.MissedList.Remove g_NotificationRoster.MissedList.IndexOf(CStr(Token))

    ' /* remove from NC cache */
    With mItems
        .Rewind
        Do While .GetNextTag(ph) = B_OK
            If ph.Expanded Then
                i = uIndexOfToken(ph, Token)
                If i <> 0 Then
                    ' /* is in this header, so remove it */
                    ph.Notifications.Remove i

                    ' /* if that makes the count zero, remove the header too */
                    If ph.Notifications.CountItems = 0 Then _
                        .Remove .IndexOf(ph.App.Signature)

                    Set ph = Nothing

                    ' /* if we now have no headers, hide the NC */
                    If .CountItems = 0 Then
                        Me.Hide

                    Else
                        uArrange
                        uRedraw
                        
                    End If
                    Exit Do

                End If
            End If
        Loop
    End With

End Sub

' /* uIndexOfToken(): returns the index of a particular notification within a particular header */

Private Function uIndexOfToken(ByRef Header As TNCHeader, ByVal Token As Long) As Long
Dim pp As TPopUpItem
Dim i As Long

    With Header.Notifications
        If .CountItems Then
            For i = 1 To .CountItems
                Set pp = .TagAt(i)
                If pp.Token = Token Then
                    uIndexOfToken = i
                    Exit For
                
                End If
            Next i
        End If
    End With

End Function

Private Function uIsInCloseRect(ByRef Point As BPoint) As Boolean

    uIsInCloseRect = (new_BRect(mView.Width - 20, 0, mView.Width - 2, CYTITLE - 1).Contains(Point))

End Function

Public Sub Repaint()

    uRedraw

End Sub


