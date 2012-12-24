VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRequester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRequester.cls
    '/
    '/  Description:    Represents an on-screen requester instance
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mhWnd As Long
Dim mView As mfxView

Private Type T_BUTTON
    Label As String
    Frame As BRect

End Type

Dim mButton() As T_BUTTON
Dim mButtons As Long
Dim mPressed As Long

Dim mReplyTo As Long
Dim mReplyMsg As Long

Dim mTitle As String
Dim mDuration As Integer

Dim mText As String
Dim mTextRect As BRect
Dim mSubText As String
Dim mSubTextRect As BRect
Dim mIcon As MImage
Dim mAppIcon As MImage

Dim mToken As Long
Dim mEnabled As Boolean

Dim mTargetAlpha As Integer
Dim mAlpha As Integer
Dim mAlphaStep As Integer

Dim mActualView As mfxView

Implements BWndProcSink

Public Function Go(ByVal Token As Long, ByRef Data As BPackedData) As Long

    If (Data Is Nothing) Then
        Debug.Print "TRequester.Go(): no data"
        Exit Function

    End If

Dim sv As String

    ' /* following are required */

    mReplyTo = Val(Data.ValueOf("reply-to"))
    mReplyMsg = Val(Data.ValueOf("reply"))
    sv = Data.ValueOf("actions")
    If (IsWindow(mReplyTo) = 0) Or (mReplyMsg = 0) Or (sv = "") Then
        Debug.Print "TRequester.Go(): bad args"
        Exit Function

    End If

Dim sAction() As String
Dim i As Integer

    sAction = Split(sv, ",")

    For i = 0 To UBound(sAction)
        uAddButton sAction(i)

    Next i

    If mButtons = 0 Then
        Debug.Print "TRequester.Go(): no buttons"
        Exit Function

    End If

    EZRegisterClass "w>requester"
    mhWnd = EZ4AddWindow("w>requester", Me, , WS_POPUPWINDOW, WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
    If IsWindow(mhWnd) = 0 Then
        Debug.Print "TRequester.Go(): bad window"
        Exit Function

    End If

    g_GetProcessFriendlyName Token, mTitle, True

Dim szPath As String
Dim pi As BIcon

    ' /* application icon */

    szPath = g_GetExeFullPath(Token)
    If szPath <> "" Then
        With new_BIconContent(szPath)
            If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pi) Then _
                Set mAppIcon = pi.Render()

        End With
    End If

    ' /* build initial view */

Const EM_GAP = 4

    Set mView = New mfxView
    With mView
        .SizeTo 320, 160

        ' /* size vertically */

Dim cy As Long
Dim pr As BRect

        cy = 4 + 24

        ' /* notification icon - can be one of !warning, !question, !info, !critical or a custom one */

        Set mIcon = load_image_obj(Data.ValueOf("icon"))
        If Not is_valid_image(mIcon) Then _
            Set mIcon = Nothing


        .TextMode = MFX_TEXT_CLEARTYPE
        .SetFont "Droid Sans", 10, True

        ' /* text */

        mText = Replace$(Data.ValueOf("text"), "\n", vbCrLf)
        If mText <> "" Then
            .SetFont "Droid Sans", 10
            Set pr = .Bounds.InsetByCopy(4, 0)
            If Not (mIcon Is Nothing) Then _
                pr.Left = pr.Left + 48 + EM_GAP

            .MeasureString mText, pr, mTextRect
            mTextRect.ExpandBy 0, 2
            mTextRect.OffsetBy 4, cy
            cy = cy + mTextRect.Height + EM_GAP

        End If

        ' /* subtext */

        mSubText = Replace$(Data.ValueOf("subtext"), "\n", vbCrLf)
        If mSubText <> "" Then
            .SetFont "Droid Sans", 9
            Set pr = .Bounds.InsetByCopy(4, 0)
            If Not (mIcon Is Nothing) Then _
                pr.Left = pr.Left + 48 + EM_GAP

            .MeasureString mSubText, pr, mSubTextRect
            mTextRect.ExpandBy 0, 2
            mSubTextRect.OffsetBy 4, cy
            cy = cy + mSubTextRect.Height + EM_GAP

        End If

        cy = cy + 28 + 16 + 4        ' // button height + margin + edge

        .SizeTo 320, cy

    End With

    ' /* defaults */

    mEnabled = True
    mDuration = 10

    ' /* draw content */

    uDraw

    ' /* show */

    g_ShowWindow mhWnd, True, False
    SetTimer mhWnd, 1, 1000, 0

    uFadeTo 240

    gRequestId = gRequestId + 1
    mToken = gRequestId
    Go = gRequestId

End Function

Private Sub uDraw()
Dim pr As BRect
Dim i As Long

Const BUTTONRX = 6

    With mView
        .EnableSmoothing True
'        .SetHighColour rgba(0, 0, 0)
        .SetHighColour rgba(245, 245, 245)
        .FillRoundRect .Bounds

'        .SetHighColour rgba(255, 255, 255, 64)
'        .SetLowColour rgba(255, 255, 255, 0)
        .SetHighColour rgba(0, 0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 64)
        .FillRoundRect .Bounds, , , MFX_VERT_GRADIENT

        .SetHighColour rgba(0, 0, 0, 190)
        .StrokeRoundRect .Bounds

        .TextMode = MFX_TEXT_CLEARTYPE
        .SetHighColour rgba(255, 255, 255)
        .SetHighColour rgba(0, 0, 0)
        .SetFont "Droid Sans", 10, True

        ' /* title */

        Set pr = .Bounds.InsetByCopy(4, 4)
        pr.Bottom = pr.Top + .StringHeight("A")
        .DrawString mTitle, pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

        ' /* icon */

        .DrawScaledImage mIcon, .Bounds.TopLeft.OffsetByCopy(4 + 4, 4 + 16), new_BPoint(48, 48)
        .DrawScaledImage mAppIcon, .Bounds.TopLeft.OffsetByCopy(4 + 4 + 24, 4 + 16 + 24), new_BPoint(24, 24)

        ' /* text */

        If Not (mTextRect Is Nothing) Then
            .SetHighColour rgba(255, 0, 0)
'            .StrokeRect mTextRect
            .SetHighColour rgba(255, 255, 255)
            .SetHighColour rgba(0, 0, 0)
            .SetFont "Droid Sans", 10
            .DrawString mText, mTextRect

        End If

        ' /* subtext */

        If Not (mSubTextRect Is Nothing) Then
            .SetHighColour rgba(255, 255, 0)
'            .StrokeRect mSubTextRect
            .SetHighColour rgba(255, 255, 255)
            .SetHighColour rgba(0, 0, 0)
            .SetFont "Droid Sans", 9
            .DrawString mSubText, mSubTextRect

        End If

        ' /* calc largest button */



        ' /* buttons */

        Set pr = new_BRect(0, 0, 79, 27)
        i = (pr.Width * mButtons) + (16 * (mButtons - 1))
        i = Fix((.Width - i) / 2)

        pr.OffsetBy i, .Height - (pr.Height + 8)
        .SetFont "Droid Sans", 10

        For i = 1 To mButtons
            Set mButton(i).Frame = pr.OffsetByCopy(10, 10)

            If i = mPressed Then
                ' /* draw pressed */
                .SetHighColour rgba(245, 245, 245)
                .FillRoundRect pr, BUTTONRX, BUTTONRX

                .SetHighColour rgba(0, 0, 0, 80)
                .SetLowColour rgba(0, 0, 0, 40)
                .FillRoundRect pr, BUTTONRX, BUTTONRX, MFX_VERT_GRADIENT

            Else
                ' /* draw normal */

                ' /* dropshadow */

                .SetHighColour rgba(0, 0, 0, 32)
                .FillRoundRect pr.OffsetByCopy(2, 2), BUTTONRX, BUTTONRX
                
                .SetHighColour rgba(245, 245, 245)
                .FillRoundRect pr, BUTTONRX, BUTTONRX
                
                .SetHighColour rgba(0, 0, 0, 0)
                .SetLowColour rgba(0, 0, 0, 40)
                .FillRoundRect pr, BUTTONRX, BUTTONRX, MFX_VERT_GRADIENT

            End If

            ' /* border */
            .SetHighColour rgba(0, 0, 0, 143)
            .StrokeRoundRect pr, BUTTONRX, BUTTONRX

            .SetHighColour rgba(0, 0, 0)
            .DrawString mButton(i).Label, pr.Duplicate, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

            pr.OffsetBy pr.Width + 16, 0

        Next i

Const CIRCSIZE = 16

        If mDuration < 6 Then
            .SetHighColour rgba(255, 255, 255)
            .SetHighColour rgba(0, 0, 0)
            Set pr = new_BRect(0, 0, CIRCSIZE - 1, CIRCSIZE - 1)
            pr.OffsetBy .Width - (pr.Width + 4), .Height - (pr.Height + 4)
            .DrawString CStr(mDuration), pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

'        i = (MIN(mDuration, 5) / 5) * pr.Height
'        pr.InsetBy i, i
'        .FillEllipse pr


        End If

        ' /* ghosting if not enabled */

        If Not mEnabled Then
            .SetHighColour rgba(255, 255, 255, 102)
            .FillRoundRect .Bounds

        End If

        ' /* fix stupid bug in GDI+/gfxlib to do with centered text */
        .DrawString "!", .Bounds.OffsetByCopy(-32, -32)

    End With

Dim pb As mfxBitmap

    Set pb = create_dropshadow(mView, , , 190)
'IIf(mEnabled, 220, 64)
    Set mActualView = New mfxView
    With mActualView
        .SizeTo pb.Width, pb.Height
        .DrawScaledImage pb

    End With

    uApply

End Sub

Private Sub uApply()

    apply_view_to_window mActualView, mhWnd, _
                         mAlpha, _
                         Fix((g_ScreenWidth - mActualView.Width) / 2), Fix((g_ScreenHeight - mActualView.Height) / 2)

'IIf(mEnabled, 240, 190)

End Sub

Public Sub Quit()

    EZ4RemoveWindow mhWnd
    EZUnregisterClass "w>requester"
    g_PopRequest2

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static fPaused As Boolean
Static iPressed As Long
Dim i As Long

    Select Case uMsg
    Case WM_LBUTTONDOWN
        iPressed = uHitTest(new_BPointFromInt32(lParam))
        Debug.Print iPressed
        If iPressed > 0 Then
            mPressed = iPressed
            uDraw

        End If


    Case WM_LBUTTONUP
        i = uHitTest(new_BPointFromInt32(lParam))
        If i = iPressed Then
            Select Case i
            Case Is > 0
                PostMessage mReplyTo, mReplyMsg, i, ByVal mToken
                Quit
                Exit Function

            End Select

        End If

        mPressed = 0
        iPressed = 0
        uDraw


    Case WM_MOUSEMOVE
        If iPressed Then
            i = uHitTest(new_BPointFromInt32(lParam))
            If (i <> iPressed) And (mPressed <> 0) Then
                Debug.Print "cold"
                mPressed = 0
                uDraw

            ElseIf (i = iPressed) And (mPressed = 0) Then
                Debug.Print "hot"
                mPressed = iPressed
                uDraw

            End If

        End If


    Case WM_SETCURSOR
        If Not fPaused Then
            fPaused = True
            SetTimer hWnd, 2, 100, 0

        End If


    Case WM_TIMER
        Select Case wParam
        Case 1
            If Not fPaused Then
                ' /* duration */
                mDuration = mDuration - 1
                If mDuration = 0 Then
                    uDraw
                    Quit

                ElseIf mDuration < 6 Then
                    uDraw
        
                End If
            End If

        Case 2
            If WindowFromPoint(GET_X_LPARAM(GetMessagePos()), GET_Y_LPARAM(GetMessagePos())) <> hWnd Then
                fPaused = False
                KillTimer hWnd, wParam

            End If

        Case 3
            ' /* alpha fade */

            If mTargetAlpha > mAlpha Then
                ' /* fade in */
                mAlpha = mAlpha + mAlphaStep

'                Debug.Print "fade in: " & mAlpha & " / " & mTargetAlpha

                ' /* if we've over-run, snap to target alpha */
                If mAlpha > mTargetAlpha Then _
                    mAlpha = mTargetAlpha

                uApply

            ElseIf mTargetAlpha < mAlpha Then
                ' /* fade out */
                mAlpha = mAlpha - mAlphaStep

                ' /* if we've under-run, snap to target alpha */
                If mAlpha < mTargetAlpha Then _
                    mAlpha = mTargetAlpha

                uApply

            Else
                ' /* stop */
                KillTimer hWnd, wParam

            End If

        End Select

    Case &H40F
        fPaused = (wParam = 0)

'    Case Else
'        Debug.Print "TRequester: " & g_HexStr(wParam) & " " & g_HexStr(lParam)

    End Select

End Function

Private Sub uAddButton(ByVal Label As String)

    mButtons = mButtons + 1
    ReDim Preserve mButton(mButtons)
    mButton(mButtons).Label = Label

End Sub

Private Function uHitTest(ByRef Position As BPoint) As Long

    ' /* zero means client, -1 means close button, >0 indicates button */

Dim i As Long

    For i = 1 To mButtons
        If mButton(i).Frame.Contains(Position) Then
            uHitTest = i
            Exit Function

        End If

    Next i

    uHitTest = 0

End Function

Public Sub SetEnabled(ByVal Enabled As Boolean)

    SendMessage mhWnd, &H40F, IIf(Enabled, -1, 0), ByVal 0&
    EnableWindow mhWnd, IIf(Enabled, -1, 0)
    mEnabled = Enabled
    uDraw

    uFadeTo IIf(Enabled, 240, 190)

End Sub

Private Sub uFadeTo(ByVal Target As Integer)

    If Target = mAlpha Then _
        Exit Sub

'    Debug.Print "FADETO: " & CStr(Abs(Target - mAlpha))
'    Debug.Print "FADETO: " & CStr(Abs(Target - mAlpha) / 30)

    mAlphaStep = ROUNDUP(Abs(Target - mAlpha) / 30)
    Debug.Print "FADETO: step = " & mAlphaStep

    mTargetAlpha = Target
    SetTimer mhWnd, 3, 10, 0

End Sub
