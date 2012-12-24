VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge Sample"
   ClientHeight    =   3795
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
   ScaleHeight     =   3795
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
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
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":001E
      Height          =   1035
      Left            =   60
      TabIndex        =   5
      Top             =   2700
      Width           =   4395
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

Private Sub Command4_Click()

    If snDoRequest("reg?app-sig=" & App.ProductName & "&title=" & App.Title & "&icon=" & App.Path & "\icon.png") > 0 Then

        snDoRequest "notify?app-sig=" & App.ProductName & _
                    "&uid=" & Text1.Text & _
                    "&title=" & Text1.Text & _
                    "&text=" & Text2.Text & _
                    "&icon=" & App.Path & "\icon.png" & _
                    "&merge-uid=" & Text1.Text

    Else
        MsgBox "Error registering with Snarl.  Check Snarl is running.", vbExclamation Or vbOKOnly, "Snarl Sample"

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    snDoRequest "unreg?app-sig=" & App.ProductName

End Sub

