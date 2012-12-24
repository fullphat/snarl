VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SnarlMail Log"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Inbox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2940
      TabIndex        =   4
      Top             =   3900
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Prefs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   3900
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Meeting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1500
      TabIndex        =   2
      Top             =   3900
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Email"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   3900
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    SendMessage Val(Me.Tag), WM_TEST, 0, ByVal 0&

End Sub

Private Sub Command2_Click()

    SendMessage Val(Me.Tag), WM_TEST, 1, ByVal 0&

End Sub

Private Sub Command3_Click()

    SendMessage Val(Me.Tag), snAppMsg(), SNARLAPP_DO_PREFS, ByVal 0&

End Sub

Private Sub Command4_Click()

    SendMessage Val(Me.Tag), WM_TEST, 2, ByVal 0&

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then _
        PostQuitMessage 0

End Sub

Public Sub Add(ByVal Text As String)

    With Form1.List1
        .AddItem Text
        .ListIndex = .ListCount - 1
        g_Debug "_add(): " & Text

    End With

End Sub
