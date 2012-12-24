VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   4365
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
   ScaleHeight     =   4365
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Sticky"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3780
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
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
      Top             =   3780
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "All new applications should ideally utilise the new V42 API and only fall back to earlier versions of the API if necessary."
      Height          =   495
      Left            =   60
      TabIndex        =   7
      Top             =   3180
      Width           =   4395
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":001E
      Height          =   855
      Left            =   60
      TabIndex        =   6
      Top             =   2220
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

Dim mToken As Long
Dim mMsg As Long

Private Sub Command1_Click()

    If mMsg Then _
        snUpdateMessage mMsg, Text1.Text, Text2.Text

End Sub

Private Sub Command4_Click()

    mMsg = snShowMessageEx("Test Class", Text1.Text, Text2.Text, IIf(Check1.Value = vbChecked, 0, -1), App.Path & "\icon.png", Me.hWnd, 0)

End Sub

Private Sub Form_Load()
Dim hr As Long

    If snGetSnarlWindow() = 0 Then
        MsgBox "Snarl isn't running - launch Snarl, then run this demo.", vbExclamation Or vbOKOnly, App.Title
        Unload Me

    Else
        mToken = snRegisterApp(App.Title, App.Path & "\icon.png", App.Path & "\icon.png", Me.hWnd, 0)
        Debug.Print mToken

        Me.Caption = App.Title

'        If hr = 0 Then
'            Me.Caption = "Registered with Snarl V" & snGetVersionEx()
'            snRegisterAlert App.Title, "Test Class"
'
'        Else
'            Me.Caption = "Error registering with Snarl: " & Hex$(hr)
'
'        End If

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim hr As Long

    hr = snUnregisterApp()
    Debug.Print hr

End Sub


