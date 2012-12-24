VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pulser"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
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
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3060
      TabIndex        =   2
      Top             =   1200
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pause"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3840
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":000C
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)

Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    If uMsg = snAppMsg() Then
        Debug.Print "callback: " & CStr(wParam)
        Select Case wParam
        Case SNARLAPP_DO_ABOUT
            ' /* no need to handle this: we provided a hint when we registered */

        Case SNARLAPP_DO_PREFS
            ' /* show our prefs */
            Me.Show

        Case SNARLAPP_ACTIVATED
            ' /* show our prefs */
            Me.Show

        Case SNARLAPP_QUIT_REQUESTED
            PostQuitMessage 0

        End Select

    End If

End Function

Private Sub Command1_Click()

    Timer1.Enabled = Not Timer1.Enabled
    Command1.Caption = IIf(Timer1.Enabled, "Pause", "Resume")
 
End Sub

Private Sub Command2_Click()

    PostQuitMessage 0

End Sub

Private Sub Form_Load()

    window_subclass Me.hWnd, Me
    Timer1_Timer

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Me.Hide
        Cancel = -1

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    snarl_unregister App.ProductName
    window_subclass Me.hWnd, Nothing

End Sub

Private Sub Timer1_Timer()
Dim hr As Long

    If snDoRequest("register?app-sig=" & App.ProductName & _
                   "&app-title=" & App.Title & _
                   "&icon=file://" & App.Path & "\icon.png" & _
                   "&reply-to=" & CStr(Me.hWnd) & _
                   "&app-daemon=1" & _
                   "&hint=Version " & App.Major & "." & App.Minor & "\n\n" & App.LegalCopyright & "\nAll Rights Reserved") > 0 Then

        snDoRequest "notify?app-sig=" & App.ProductName & "&title=Pulse&icon=file://" & App.Path & "\icon.png"

    End If

End Sub
