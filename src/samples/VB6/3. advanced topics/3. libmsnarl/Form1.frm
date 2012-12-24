VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
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
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Allow merging"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   2160
      Value           =   1  'Checked
      Width           =   4395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1740
      TabIndex        =   5
      Top             =   3180
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":000C
      Top             =   1080
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Notification Title"
      Top             =   360
      Width           =   4395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   3180
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mGotSnarl As Boolean
Dim mMsg As Long

Dim WithEvents myApp As SnarlApp
Attribute myApp.VB_VarHelpID = -1

Private Sub Command1_Click()

'    If mMsg Then _
        sn41EZUpdate mMsg, , Text2.Text

End Sub

Private Sub Command4_Click()

    mMsg = myApp.EZNotify("", Text1.Text, Text2.Text, 0, "shell32.dll,-17", , , , , IIf(Check1.Value = vbChecked, NOTIFICATION_ALLOWS_MERGE, 0))

End Sub

Private Sub Form_Load()

    Set myApp = New SnarlApp

    If is_snarl_running() Then
        uRegister

    Else
        Me.Caption = "Snarl not running, waiting..."
        uQuit

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Private Sub uRegister()

    If myApp.SetTo(App.ProductName, App.Title, "shell32.dll,-17") = B_OK Then
        Me.Caption = "Registered with Snarl V" & CStr(snarl_version()) & " (" & Hex$(myApp.Token) & ")"
        Command1.Enabled = True
        Command4.Enabled = True

    End If

End Sub

Private Sub uQuit()

    Command1.Enabled = False
    Command4.Enabled = False

End Sub

Private Sub myApp_SnarlLaunched()

    uRegister

End Sub

Private Sub myApp_SnarlQuit()

    Me.Caption = "Snarl has quit, waiting..."
    uQuit

End Sub

