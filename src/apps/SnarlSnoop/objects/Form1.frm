VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SnarlSnooper Log"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
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
   ScaleHeight     =   3090
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Prefs"
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3180
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mClassId As String

Private Sub Command1_Click()

    SendMessage Val(Me.Tag), WM_TEST, 0, ByVal 0&

End Sub

Private Sub Command2_Click()

    SendMessage Val(Me.Tag), WM_TEST, 2, ByVal 0&

End Sub

Private Sub Command3_Click()

    SendMessage Val(Me.Tag), sn41AppMsg(), SNARL41_APP_PREFS, ByVal 0&

End Sub

Private Sub Command4_Click()

    mClassId = CStr(Rnd * 65535)
    sn41AddClass Val(List1.Tag), mClassId, CStr(mClassId)

End Sub

Private Sub Command5_Click()

    sn41RemClass Val(List1.Tag), mClassId

End Sub

Private Sub Command6_Click()

    sn41RemAllClasses Val(List1.Tag)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then _
        PostQuitMessage 0

End Sub

Public Sub Add(ByVal Text As String)

    With Form1.List1
        .AddItem Text
        .ListIndex = .ListCount - 1

    End With

End Sub
