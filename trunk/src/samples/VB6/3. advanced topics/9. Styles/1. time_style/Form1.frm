VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snarl time format sample"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   780
      TabIndex        =   4
      Top             =   1980
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   795
      Left            =   3780
      TabIndex        =   3
      Top             =   1500
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   780
      TabIndex        =   2
      Top             =   1500
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Style:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Text:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   4515
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":008E
      Height          =   615
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

Private Sub Command1_Click()
Dim hr As Long

    hr = snarl_register(App.ProductName, App.Title, App.Path & "\icon.png")
    If hr < 0 Then
        MsgBox "Error registering with Snarl (" & CStr(Abs(hr)) & ")", vbExclamation Or vbOKOnly, App.Title

    Else
        snDoRequest "notify?app-sig=" & App.ProductName & _
                    "&title=Lorem Ipsum" & _
                    "&text=" & Text1.Text & _
                    "&style=" & Text2.Text

    End If

End Sub

Private Sub Form_Load()

    Text1.Text = Format$(Now(), "yyyymmddhhnnss")
    Text2.Text = "Clock/Analog"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    snarl_unregister App.ProductName

End Sub
