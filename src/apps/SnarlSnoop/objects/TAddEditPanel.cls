VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAddEditPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mPage As BPrefsPage
Dim mhWndPanel As Long

Dim mItem As CConfSection
'Dim mOriginalItem As CConfSection

Public Event Update(ByRef Item As CConfSection)
'Public Event AddNew(ByRef Item As CConfSection)

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Select Case Control.GetName()
    Case "task_title"
        mItem.Update "title", Value

    Case "task_text"
        mItem.Update "text", Value

    Case "task_icon"
        mItem.Update "icon", Value

    Case "task_ack"
        mItem.Update "cmd", Value

    Case "task_name"
        mItem.SetName Value

    Case Else
        Debug.Print Control.GetName

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pc As BControl

    Select Case Control.GetName()
    Case "fb2>done"

        If Not (uCheckValue("rec_minute", 0, 59)) Then
            MsgBox "Minute value appears to be invalid...", vbExclamation Or vbOKOnly, App.Title
            Exit Sub

        End If

        If Not (uCheckValue("rec_hour", 0, 23)) Then
            MsgBox "Minute value appears to be invalid...", vbExclamation Or vbOKOnly, App.Title
            Exit Sub

        End If

        If Not (uCheckValue("rec_day", 1, 31)) Then
            MsgBox "Minute value appears to be invalid...", vbExclamation Or vbOKOnly, App.Title
            Exit Sub

        End If

        If Not (uCheckValue("rec_month", 1, 12)) Then
            MsgBox "Minute value appears to be invalid...", vbExclamation Or vbOKOnly, App.Title
            Exit Sub

        End If

        If Not (uCheckValue("rec_weekday", 0, 6)) Then
            MsgBox "Minute value appears to be invalid...", vbExclamation Or vbOKOnly, App.Title
            Exit Sub

        End If

        ' /* build the recurrence value */

        mItem.Update "recurrence", uValueOf("rec_minute") & " " & _
                                   uValueOf("rec_hour") & " " & _
                                   uValueOf("rec_day") & " " & _
                                   uValueOf("rec_month") & " " & _
                                   uValueOf("rec_weekday")

        RaiseEvent Update(mItem)

        mPanel.Quit

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 100
    mPage.SetWidth 220

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    EnableWindow mhWndPanel, -1
    g_ShowWindow mhWndPanel, True, True

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Friend Sub Go(ByVal hWndPanel As Long, ByRef Item As CConfSection, ByVal AddNew As Boolean)
Dim pPage As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg
Dim sz() As String

    Set mItem = Item.Duplicate
'    Set mOriginalItem = Item

    mhWndPanel = hWndPanel
    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetWindow 1
        .SetTitle IIf(AddNew, "Add New", "Edit") & " Task..."
        .SetWidth 440

        Set pPage = new_BPrefsPage("", , Me)
        With pPage
            .Add new_BPrefsControl("banner", "", "Details")
'            .Add new_BPrefsControl("label", "task_name", Item.Name, "Name:")
            .Add new_BPrefsControl("fancyedit", "task_title", Item.GetValueWithDefault("title", ""), "Title:")
            .Add new_BPrefsControl("fancyedit", "task_text", Item.GetValueWithDefault("text", ""), "Text:")
            .Add new_BPrefsControl("label", "", "Note that if only a title is provided, it will be displayed in the notification's text element")
            .Add new_BPrefsControl("fancyfilepicker", "task_icon", Item.GetValueWithDefault("icon", ""), "Icon:")
            .Add new_BPrefsControl("label", "", "If you want a command to be executed when the notification is clicked, enter it below:", "Command:")
            .Add new_BPrefsControl("fancyedit", "task_ack", Item.GetValueWithDefault("cmd", ""), "")

                            '# .---------------- minute (0 - 59)
                            '# |  .------------- hour (0 - 23)
                            '# |  |  .---------- day of month (1 - 31)
                            '# |  |  |  .------- month (1 - 12) OR jan,feb,mar,apr ...
                            '# |  |  |  |  .---- day of week (0 - 6) (Sunday=0 or 7)  OR sun,mon,tue,wed,thu,fri,sat
                            '# |  |  |  |  |
                            '  *  *  *  *  *  command to be executed

            .Add new_BPrefsControl("banner", "", "Recurrence")
            sz = Split(mItem.GetValueWithDefault("recurrence"), " ")
            .Add new_BPrefsControl("label", "", "Use * in any of the following to indicate a wild card value", "")
            .Add new_BPrefsControl("fancyedit", "rec_minute", sz(0), "Minute:")
            .Add new_BPrefsControl("fancyedit", "rec_hour", sz(1), "Hour:")
            .Add new_BPrefsControl("label", "", "Use the 24 hour clock here (i.e. 0-23)")
            .Add new_BPrefsControl("fancyedit", "rec_day", sz(2), "Day:")
            .Add new_BPrefsControl("fancyedit", "rec_month", sz(3), "Month:")
            .Add new_BPrefsControl("fancyedit", "rec_weekday", sz(4), "Weekday:")
            .Add new_BPrefsControl("label", "", "Sunday=0, Monday=1, Saturday=6, etc.")

            .Add new_BPrefsSeparator
            Set pc = new_BPrefsControl("fancybutton2", "fb2>done", "Ok")
            pc.SizeTo 80, 0             ' // a zero height here means "let control decide"
            .Add pc

        End With
        .AddPage pPage

        .Go
        g_SetWindowIconToAppResourceIcon .hwnd

    End With

Dim dw As Long

    dw = GetWindowLong(mPanel.hwnd, GWL_STYLE)
    dw = dw And (Not WS_MINIMIZEBOX)
    SetWindowLong mPanel.hwnd, GWL_STYLE, dw

Dim rcConfig As RECT
Dim rc As RECT

    SetWindowLong mPanel.hwnd, GWL_HWNDPARENT, hWndPanel
    EnableWindow hWndPanel, 0

    GetWindowRect hWndPanel, rc
    GetWindowRect mPanel.hwnd, rcConfig
    g_RectNormalise rcConfig

    g_MoveWindow mPanel.hwnd, _
                 rc.Left + Fix(((rc.Right - rc.Left) - rcConfig.Right) / 2), _
                 Fix((g_ScreenHeight - (rcConfig.Bottom - rcConfig.Top)) / 2)

    g_ShowWindow mPanel.hwnd, True, True

End Sub

Private Function uToWildcard(ByVal Value As Integer) As String

    uToWildcard = IIf(Value = -1, "*", CStr(Value))

End Function

Private Function uValidate(ByVal Value As String, ByVal MIN As Integer, ByVal MAX As Integer) As Boolean
Dim i As Integer

    On Error Resume Next

    uValidate = True
    If Value = "*" Then _
        Exit Function

    uValidate = False
    If Not g_IsNumeric(Value) Then _
        Exit Function

    Err.Clear
    i = CInt(Val(Value))
    If Err.Number <> 0 Then _
        Exit Function

    If (Value < MIN) Or (Value > MAX) Then _
        Exit Function

    uValidate = True

End Function

Private Function uCheckValue(ByVal Name As String, ByVal MIN As Integer, ByVal MAX As Integer) As Boolean

    If (mPanel Is Nothing) Then _
        Exit Function

Dim pc As BControl

    If mPanel.Find(Name, pc) Then _
        uCheckValue = uValidate(pc.GetText, MIN, MAX)

End Function

Private Function uValueOf(ByVal Name As String) As String
Dim pc As BControl

    If mPanel.Find(Name, pc) Then _
        uValueOf = pc.GetText

End Function
