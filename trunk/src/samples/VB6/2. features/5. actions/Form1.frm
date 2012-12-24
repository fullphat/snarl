VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snarl Actions Sample"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
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
   ScaleHeight     =   3855
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Actions"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   3300
      TabIndex        =   5
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Action"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   3300
      TabIndex        =   4
      Top             =   900
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1500
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1500
      TabIndex        =   2
      Top             =   660
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "@1234 @-768 @32767 @999"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   3540
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "calc.exe"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   3300
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "http://www.google.co.uk"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   3060
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":000C
      Height          =   795
      Left            =   60
      TabIndex        =   9
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Command:"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mMsg As Long

Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    If uMsg = &H407 Then
        If LoWord(wParam) = SNARL_NOTIFY_ACTION Then
            Label3.Caption = ""

            Select Case HiWord(wParam)
            Case 333
                Label3.BackColor = vbRed

            Case 666
                Label3.BackColor = vbGreen

            Case 999
                Label3.BackColor = vbBlue

            Case 0
                Label3.BackColor = vb3DFace

            Case Else
                Label3.Caption = vbCrLf & "Dynamic callback '" & CStr(HiWord(wParam)) & "'"
                Label3.BackColor = vb3DFace

            End Select

        End If

    End If

End Function

Private Sub Command1_Click(Index As Integer)
Dim hr As Long

    If Index = 0 Then
        hr = snDoRequest("addaction?token=" & CStr(mMsg) & "&label=" & Text1(0).Text & "&cmd=" & Text1(1).Text)
        If hr <> 0 Then
            Label3.Caption = vbCrLf & "Error adding action (" & CStr(hr) & ")"
            
        Else
            Label3.Caption = ""

        End If

    ElseIf Index = 1 Then
        hr = snDoRequest("clearactions?token=" & CStr(mMsg))

    End If

    Debug.Print hr

End Sub

Private Sub Command4_Click()
Dim hr As Long

    ' /* in order to have actions against a notification, we must have a valid reply-to window and reply message */

    hr = snarl_register(App.ProductName, App.Title, App.Path & "\icon.png", , Me.hWnd, &H407)
    If hr <= 0 Then
        MsgBox "Error registering with Snarl (" & CStr(hr) & ")", vbExclamation Or vbOKOnly, App.Title

    Else

'        mMsg = snDoRequest("notify?app-sig=" & App.ProductName & _
                           "&title=Notifications with actions" & _
                           "&text=The gear icon in the bottom right corner indicates the notification supports actions; " & _
                           "you can access these by clicking the gear gadget in the top-left corner." & _
                           "&timeout=0" & _
                           "&action=Red,@333&action=Green,@666&action=Blue,@999&action=Reset,@000")

        mMsg = snDoRequest("notify?app-sig=" & App.ProductName & _
                           "&title=Launch detection" & _
                           "&text=Indian Ocean station 4 has detected ICBM launch" & _
                           "&timeout=0&icon=!system-warning" & _
                           "&action=Protest,@333&action=Ignore,@666&action=Panic,@999&action=Invoke Armageddon,@000")

        Command1(1).Enabled = True

    End If

End Sub

Private Sub Form_Load()

    window_subclass Me.hWnd, Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    snarl_unregister App.ProductName
    window_subclass Me.hWnd, Nothing

End Sub

Private Sub Text1_Change(Index As Integer)

    Command1(0).Enabled = ((Text1(0).Text <> "") And (Text1(1).Text <> ""))

End Sub


