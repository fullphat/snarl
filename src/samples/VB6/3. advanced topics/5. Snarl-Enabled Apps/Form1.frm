VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snarl-Enabled App"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   237
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Use static hint text"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Register using callback"
      Height          =   555
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Register using config-tool"
      Height          =   555
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " "
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2220
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "This will launch notepad in place of the configuration tool"
      Height          =   435
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    If uMsg = snAppMsg() Then
        Debug.Print "callback: " & CStr(wParam)
        Select Case wParam
        Case SNARLAPP_DO_ABOUT
            Label2.Caption = "SHOW ABOUT"
            MsgBox "There's nothing to see here...", vbOKOnly Or vbInformation, App.Title

        Case SNARLAPP_DO_PREFS
            Label2.Caption = "SHOW PREFS"
            Me.WindowState = FormWindowStateConstants.vbNormal

        Case SNARLAPP_ACTIVATED
            Label2.Caption = "ACTIVATED"
            Me.WindowState = FormWindowStateConstants.vbNormal

        Case SNARLAPP_QUIT_REQUESTED
            Label2.Caption = "QUIT"
            Unload Me

        End Select

    End If

End Function

Private Sub Command1_Click()

    snarl_unregister App.ProductName, "abc"
    snDoRequest "register?app-sig=" & App.ProductName & _
                "&app-title=" & App.Title & _
                "&icon=file://" & App.Path & "\icon.png" & _
                "&password=abc" & _
                "&reply-to=" & CStr(Me.hwnd) & _
                "&app-daemon=1" & _
                IIf(Check1.Value = vbChecked, "&hint=Version 10.267.1 Gamma Zero\n\n©2011 Acme Productions, Inc.\nAll Rights Reserved", "")

End Sub

Private Sub Command2_Click()

    snarl_unregister App.ProductName, "abc"
    snDoRequest "register?app-sig=" & App.ProductName & _
                "&app-title=" & App.Title & _
                "&icon=file://" & App.Path & "\icon.png" & _
                "&hint=Version 10.267.1 Gamma Zero\n\n©2011 Acme Productions, Inc.\nAll Rights Reserved" & _
                "&password=abc" & _
                "&config-tool=notepad.exe"

End Sub

Private Sub Form_Load()

    window_subclass Me.hwnd, Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    snarl_unregister App.ProductName, "abc"
    window_subclass Me.hwnd, Nothing

End Sub
