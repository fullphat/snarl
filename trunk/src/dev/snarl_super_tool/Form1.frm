VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "SnarlSuperTool"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H005ED27D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00263C2E&
      Height          =   270
      Left            =   60
      TabIndex        =   1
      Top             =   2580
      Width           =   3075
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00263C2E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005ED27D&
      Height          =   2295
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4515
   End
   Begin VB.Label Label1 
      BackColor       =   &H005ED27D&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00263C2E&
      Height          =   285
      Left            =   1740
      TabIndex        =   2
      Top             =   3180
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim sz As String

    Select Case uMsg
    Case &H400
        sz = "** Signal " & uError(LoWord(wParam)) & " from notification token $" & g_HexStr(lParam) & " **"
        uOut sz

    End Select

End Function

Private Sub Form_Load()

    AddSubClass Me.hWnd, Me

    uOut vbCrLf & "Welcome to SnarlSuperTool " & App.Major & "." & App.Minor & " (Build " & App.Revision & ")"
    uOut App.LegalCopyright
    uOut ""

'    If uCheckSnarl() Then _
'        uOut "Snarl V" & snDoRequest("version")
'
'    uOut ""
    uPrompt

    g_SimpleSetWindowIcon Me.hWnd

    Me.Show
    Text2.SetFocus

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    Text2.Move 12, Me.ScaleHeight - Text2.Height - 0, Me.ScaleWidth - 12
    Text1.Move 0, 0, Me.ScaleWidth - 0, Text2.Top '- 1

    Label1.Move 0, Text2.Top - 1

End Sub

'Private Function uCheckSnarl() As Boolean
'
'    uCheckSnarl = snIsSnarlRunning()
'    If Not uCheckSnarl Then _
'        uOut "Snarl is not running"
'
'End Function

Private Sub uOut(ByVal Text As String, Optional ByVal AddCRLF As Boolean = True)

    With Text1
        .Text = .Text & Text & IIf(AddCRLF, vbCrLf, "")
        .SelLength = 0
        .SelStart = Len(.Text)

    End With

End Sub

Private Sub uPrompt()

    With Text2
        .Text = ""
        .SelLength = 0
        .SelStart = Len(.Text)

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    RemoveSubClass Me.hWnd

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then _
        Text2.SetFocus

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

'    Select Case KeyCode
'    Case vbKeyBack
'        Debug.Print Len(Text2.Text)
'        If Len(Text2.Text) = 2 Then _
'            KeyCode = 0
'
'    End Select

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim sz As String
Dim hr As Long

    Select Case KeyAscii
'    Case vbKeyBack
'        If Len(Text2.Text) = 2 Then _
'            KeyAscii = 0

    Case 13
        KeyAscii = 0
        sz = Text2.Text         '//(Text2.Text, Len(Text2.Text) - 0)
        If (sz = "help") Or (sz = "?") Then
            uOut "--SST Help--" & vbCrLf & "Enter a Snarl request at the prompt and press return.  You can use the %HWND% marker to enable signals to this application window." & vbCrLf & "Some examples:" & vbCrLf & "  register?app-sig=foo/bar&title=Some App%HWND%" & vbCrLf & "  notify?app-sig=foo/bar&title=Hello, world!" & vbCrLf
            uPrompt

        ElseIf sz <> "" Then
            uOut g_Quote(sz) & " -> ", False
            hr = snDoRequest(Replace$(sz, "%HWND%", "&reply-to=" & CStr(Me.hWnd) & "&reply-with=" & CStr(&H400)))
            If hr >= 0 Then
                uOut "OK: " & CStr(hr)  '& vbCrLf

            Else
                hr = Abs(hr)
                uOut "Error " & CStr(hr) & ": " & uError(hr)  '& vbCrLf
            
            End If

            uPrompt

        End If

    End Select

End Sub

Private Function uError(ByVal Error As Long) As String

    Select Case Error
    Case SNARL_SUCCESS:                     uError = "SNARL_SUCCESS"
    Case SNARL_CALLBACK_R_CLICK:            uError = "SNARL_CALLBACK_R_CLICK"
    Case SNARL_CALLBACK_TIMED_OUT:          uError = "SNARL_CALLBACK_TIMED_OUT"
    Case SNARL_CALLBACK_INVOKED:            uError = "SNARL_CALLBACK_INVOKED"
    Case SNARL_CALLBACK_MENU_SELECTED:      uError = "SNARL_CALLBACK_MENU_SELECTED"
    Case SNARL_CALLBACK_M_CLICK:            uError = "SNARL_CALLBACK_M_CLICK"
    Case SNARL_CALLBACK_CLOSED:             uError = "SNARL_CALLBACK_CLOSED"
    Case SNARL_ERROR_FAILED:                uError = "SNARL_ERROR_FAILED"
    Case SNARL_ERROR_UNKNOWN_COMMAND:       uError = "SNARL_ERROR_UNKNOWN_COMMAND"
    Case SNARL_ERROR_TIMED_OUT:             uError = "SNARL_ERROR_TIMED_OUT"
    Case SNARL_ERROR_BAD_SOCKET:            uError = "SNARL_ERROR_BAD_SOCKET"
    Case SNARL_ERROR_BAD_PACKET:            uError = "SNARL_ERROR_BAD_PACKET"
    Case SNARL_ERROR_INVALID_ARG:           uError = "SNARL_ERROR_INVALID_ARG"
    Case SNARL_ERROR_ARG_MISSING:           uError = "SNARL_ERROR_ARG_MISSING"
    Case SNARL_ERROR_SYSTEM:                uError = "SNARL_ERROR_SYSTEM"
    Case SNARL_ERROR_ACCESS_DENIED:         uError = "SNARL_ERROR_ACCESS_DENIED"
    Case SNARL_ERROR_UNSUPPORTED_VERSION:   uError = "SNARL_ERROR_UNSUPPORTED_VERSION"
    Case SNARL_ERROR_NO_ACTIONS_PROVIDED:   uError = "SNARL_ERROR_NO_ACTIONS_PROVIDED"
    Case SNARL_ERROR_UNSUPPORTED_ENCRYPTION:    uError = "SNARL_ERROR_UNSUPPORTED_ENCRYPTION"
    Case SNARL_ERROR_UNSUPPORTED_HASHING:   uError = "SNARL_ERROR_UNSUPPORTED_HASHING"
    Case SNARL_ERROR_NOT_RUNNING:           uError = "SNARL_ERROR_NOT_RUNNING"
    Case SNARL_ERROR_NOT_REGISTERED:        uError = "SNARL_ERROR_NOT_REGISTERED"
    Case SNARL_ERROR_ALREADY_REGISTERED:    uError = "SNARL_ERROR_ALREADY_REGISTERED"
    Case SNARL_ERROR_CLASS_ALREADY_EXISTS:  uError = "SNARL_ERROR_CLASS_ALREADY_EXISTS"
    Case SNARL_ERROR_CLASS_BLOCKED:         uError = "SNARL_ERROR_CLASS_BLOCKED"
    Case SNARL_ERROR_CLASS_NOT_FOUND:       uError = "SNARL_ERROR_CLASS_NOT_FOUND"
    Case SNARL_ERROR_NOTIFICATION_NOT_FOUND:    uError = "SNARL_ERROR_NOTIFICATION_NOT_FOUND"
    Case SNARL_ERROR_FLOODING:              uError = "SNARL_ERROR_FLOODING"
    Case SNARL_ERROR_DO_NOT_DISTURB:        uError = "SNARL_ERROR_DO_NOT_DISTURB"
    Case SNARL_ERROR_COULD_NOT_DISPLAY:     uError = "SNARL_ERROR_COULD_NOT_DISPLAY"
    Case SNARL_ERROR_AUTH_FAILURE:          uError = "SNARL_ERROR_AUTH_FAILURE"
    Case SNARL_ERROR_DISCARDED:             uError = "SNARL_ERROR_DISCARDED"
    Case SNARL_ERROR_NOT_SUBSCRIBED:        uError = "SNARL_ERROR_NOT_SUBSCRIBED"
    Case SNARL_WAS_MERGED:                  uError = "SNARL_WAS_MERGED"

    Case SNARL_NOTIFY_GONE:                 uError = "SNARL_NOTIFY_GONE"
    Case 302:                               uError = "SNARL_NOTIFY_CLICK"
    Case SNARL_NOTIFY_EXPIRED:              uError = "SNARL_NOTIFY_EXPIRED"
    Case SNARL_NOTIFY_INVOKED:              uError = "SNARL_NOTIFY_INVOKED"
    Case SNARL_NOTIFY_MENU:                 uError = "SNARL_NOTIFY_MENU"
    Case 306:                               uError = "SNARL_NOTIFY_EX_CLICK"
    Case SNARL_NOTIFY_CLOSED:               uError = "SNARL_NOTIFY_CLOSED"
    Case SNARL_NOTIFY_ACTION:               uError = "SNARL_NOTIFY_ACTION"
    Case SNARL_NOTIFY_APP_DO_ABOUT:         uError = "SNARL_NOTIFY_APP_DO_ABOUT"
    Case SNARL_NOTIFY_APP_DO_PREFS:         uError = "SNARL_NOTIFY_APP_DO_PREFS"

    Case Else:                              uError = "Undefined error"

    End Select

End Function

