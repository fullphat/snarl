VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAppsPopUpWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TAppsPopUpWindow.cls
    '/
    '/  Description:    Implements the floating daemon-picker window
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Event Closed()
Public Event Selected(ByVal Signature As String)

Const CLASS_NAME = "w>SnarlDaemons"
Dim mhWnd As Long

Dim mView As mfxView
Dim mItems As BTagList
Dim mItemHeight As Long
Dim mIconSize As Long

'Const CXICON = 32
Const CXMARGIN = 4
'Const CYITEM = 36
Dim mSelected As String

Implements BWndProcSink

Public Sub Create(ByVal ItemHeight As Long)

    mItemHeight = ItemHeight
    mIconSize = mItemHeight - 2
    Set mItems = new_BTagList()

End Sub

Public Sub AddItem(ByVal Label As String, ByVal Data As String, Optional ByRef Icon As MImage)
Dim ppi As TPopUpItem

    Set ppi = New TPopUpItem
    ppi.SetTo Label, Data, Icon
    mItems.Add ppi

End Sub

Public Sub Show(Optional ByRef Position As BRect)

    If mItems.CountItems = 0 Then
        RaiseEvent Closed
        Exit Sub

    End If

Dim dw As Long

    dw = GetMessagePos()

    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me, , , WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

    Set mView = New mfxView
    With mView
        .SetFont "Tahoma", 9 ', True
        .SetFont "Arial", 8 ', True
        .TextMode = MFX_TEXT_ANTIALIAS

    End With

Dim pp As TPopUpItem
Dim cx As Long
Dim n As Long

    ' /* figure out required width */

    cx = 80
    With mItems
        .Rewind
        Do While .GetNextTag(pp) = B_OK
            n = mView.StringWidth(pp.Label) + 24
            If NOTNULL(pp.Icon) Then _
                n = n + mIconSize + 4

            cx = MAX(cx, n)

        Loop

    End With


'    If ISNULL(Actions) Then
'        ' /* daemons */
'        cx = 50
'        With g_AppRoster
'            If .CountApps Then
'                For i = 1 To .CountApps
'                    If .AppAt(i).IncludeInMenu Then
'                        Set pp = New TPopUpItem
'                        Set pp.Item = .AppAt(i)
'                        pp.Name = .AppAt(i).Name
'                        pp.Value = .AppAt(i).Signature
'                        Set pp.Icon = .AppAt(i).CachedIcon
'                        mItems.Add pp
'                        cx = MAX(cx, mView.StringWidth(.AppAt(i).Name))
'
'                    End If
'                Next i
'            End If
'        End With
'
'    Else
'        cx = 80
'        With Actions
'            .Rewind
'            Do While .GetNextTag(pn) = B_OK
'                i = i + 1
'                Set pp = New TPopUpItem
'                Set pp.Item = pn
'                pp.Name = pn.Name
'                pp.Value = pn.Value '// CStr(i)
'                mItems.Add pp
'                cx = MAX(cx, mView.StringWidth(pn.Name) + 24)
'
'            Loop
'        End With
'
'    End If

'    If mItems.CountItems = 0 Then _
        Exit Sub

    ' /* set up the view */

Dim pr As BRect

    mView.SizeTo cx + (CXMARGIN * 2), (mItems.CountItems * mItemHeight) + ((mItems.CountItems - 1) * 2) + (CXMARGIN * 2)
    Set pr = mView.Bounds
    pr.InsetBy CXMARGIN, CXMARGIN
    pr.Bottom = pr.Top + (mItemHeight - 1)


'    If ISNULL(Actions) Then
'        ' /* apps */
'        mView.SizeTo cx + CXICON + 4 + (CXMARGIN * 2) + 64, (mItems.CountItems * CYITEM) + ((mItems.CountItems - 1) * 2) + (CXMARGIN * 2)
'        Set pr = mView.Bounds
'        pr.InsetBy CXMARGIN, CXMARGIN
'        pr.Bottom = pr.Top + (CYITEM - 1)
'
'    Else
'        mView.SizeTo cx + (CXMARGIN * 2), (mItems.CountItems * 24) + ((mItems.CountItems - 1) * 2) + (CXMARGIN * 2)
'        Set pr = mView.Bounds
'        pr.InsetBy CXMARGIN, CXMARGIN
'        pr.Bottom = pr.Top + (24 - 1)
'
'    End If

    With mItems
        .Rewind
        Do While .GetNextTag(pp) = B_OK
            Set pp.Frame = pr.Duplicate
            pr.OffsetBy 0, pr.Height + 2

        Loop

    End With

    uRedraw

Dim pt As POINTAPI

    If ISNULL(Position) Then
        pt.x = GET_X_LPARAM(dw) - Fix(mView.Width / 2)
        pt.y = GET_Y_LPARAM(dw) - Fix(mView.Height / 2)
        If pt.x + mView.Width + 16 > g_ScreenWidth Then _
            pt.x = g_ScreenWidth() - (mView.Width + 16)

        If pt.y + mView.Height + 16 > g_ScreenHeight() Then _
            pt.y = g_ScreenHeight() - (mView.Height + 16)

    Else
        pt.x = Position.Left
        pt.y = Position.Bottom

    End If

    apply_view_to_window mView, mhWnd, , pt.x, pt.y
    g_ShowWindow mhWnd, True, True
    SetCapture mhWnd

End Sub

Public Sub Hide()

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static szPressed As String
Dim pp As TPopUpItem
Dim sz As String

    Select Case uMsg
    Case WM_CAPTURECHANGED
        PostMessage hWnd, &H401, 0, ByVal 0&

    Case WM_MOUSEMOVE
        If uHitTest(new_BPointFromInt32(lParam), pp) Then
            sz = pp.Guid

        Else
            sz = ""

        End If

        If sz <> mSelected Then
            mSelected = sz
            uRedraw
            apply_view_to_window mView, hWnd

        End If

    Case WM_LBUTTONDOWN
        If uHitTest(new_BPointFromInt32(lParam), pp) Then _
            szPressed = pp.Guid

    Case WM_LBUTTONUP
        If uHitTest(new_BPointFromInt32(lParam), pp) Then
            If szPressed = pp.Guid Then _
                RaiseEvent Selected(pp.Data)

        End If

        PostMessage hWnd, &H401, 0, ByVal 0&
        szPressed = ""

    Case &H401
        RaiseEvent Closed
        Me.Hide

    End Select

End Function

Private Sub uRedraw()
Dim pp As TPopUpItem
Dim ox As Long
Const RX = 6            ' // 14

    With mView
        .Clear
        .EnableSmoothing True

        .SetHighColour rgba(242, 242, 242)
        .FillRoundRect .Bounds, RX, RX

        .SetHighColour rgba(0, 0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 64)
        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT

        .SetHighColour rgba(0, 0, 0, 156)
        .StrokeRoundRect .Bounds, RX, RX

    End With

    With mItems
        .Rewind
        Do While .GetNextTag(pp) = B_OK
            If pp.Guid = mSelected Then
                ' /* selected item */
                mView.SetHighColour rgba(0, 0, 0, 24)
                mView.FillRoundRect pp.Frame.InsetByCopy(-1, -1), RX, RX ', MFX_VERT_GRADIENT
                mView.SetHighColour rgba(0, 0, 0, 64)
                mView.StrokeRoundRect pp.Frame.InsetByCopy(-1, -1), RX, RX

            End If

            ox = 4
            If NOTNULL(pp.Icon) Then _
                ox = ox + mItemHeight - 2

            mView.SetHighColour rgba(0, 0, 0, 250)
            mView.DrawString pp.Label, pp.Frame.OffsetByCopy(ox, 0), MFX_ALIGN_LEFT Or MFX_ALIGN_V_CENTER

            If NOTNULL(pp.Icon) Then _
                mView.DrawScaledImage pp.Icon, pp.Frame.TopLeft.OffsetByCopy(0, Fix((mItemHeight - mIconSize) / 2)), new_BPoint(mIconSize, mIconSize)

        Loop

    End With

End Sub

Private Function uHitTest(ByRef Point As BPoint, ByRef Item As TPopUpItem) As Boolean
Dim pp As TPopUpItem

    With mItems
        .Rewind
        Do While .GetNextTag(pp) = B_OK
            If pp.Frame.Contains(Point) Then
                Set Item = pp
                uHitTest = True
                Exit Function

            End If
        Loop
    End With

End Function
