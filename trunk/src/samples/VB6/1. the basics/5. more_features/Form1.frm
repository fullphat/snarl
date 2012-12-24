VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   5940
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
   ScaleHeight     =   5940
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check4 
      Caption         =   "Text"
      Height          =   255
      Left            =   1740
      TabIndex        =   15
      Top             =   4500
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Title"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   4500
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Sticky"
      Height          =   255
      Left            =   1740
      TabIndex        =   13
      Top             =   3120
      Width           =   1395
   End
   Begin VB.TextBox Text3 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   60
      PasswordChar    =   "*"
      TabIndex        =   11
      Text            =   "password"
      Top             =   300
      Width           =   4395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Register"
      Height          =   495
      Left            =   60
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unregister"
      Height          =   495
      Left            =   1740
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":000C
      Left            =   60
      List            =   "Form1.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include icon"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   60
      TabIndex        =   5
      Top             =   4860
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   795
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":004B
      Top             =   2220
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Lorem ipsum dolor sit amet"
      Top             =   1560
      Width           =   4395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   3900
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   1155
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   5520
      Width           =   4395
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1980
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1320
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'
End Sub

Private Sub Command2_Click()
Dim hr As Long

    Label3.Caption = ""
    hr = snarl_unregister(App.ProductName, Text3.Text)
    If hr < 0 Then _
        Label3.Caption = "Error unregistering with Snarl (" & CStr(Abs(hr)) & ")"

End Sub

Private Sub Command3_Click()

    Label3.Caption = ""

Dim hr As Long

    hr = snarl_register(App.ProductName, App.Title, App.Path & "\app.png", Text3.Text)
    If hr < 0 Then _
        Label3.Caption = "Error registering with Snarl (" & CStr(Abs(hr)) & ")"

End Sub

Private Sub Command4_Click()
Dim szArgs As String
Dim pri As Long
Dim hr As Long

    Select Case Combo1.ListIndex
    Case 0
        pri = -1

    Case 1
        pri = 0

    Case 2
        pri = 1

    End Select

    szArgs = "app-sig=" & App.ProductName & "&title=" & Text1.Text & "&text=" & Text2.Text & _
             "&priority=" & CStr(pri) & "&password=" & Text3.Text

    If Check1.Value = vbChecked Then _
        szArgs = szArgs & "&icon=" & "http://www.iconarchive.com/icons/iconka/santa/48/santa-5-icon.png" '  App.Path & "\icon.png"

    If Check2.Value = vbChecked Then _
        szArgs = szArgs & "&timeout=0"

    hr = snDoRequest("notify?" & szArgs)
    If hr < 0 Then _
        Label3.Caption = "snDoRequest() failed (" & CStr(Abs(hr)) & ")"


End Sub

Private Sub Form_Load()

    Me.Caption = App.Title
    Combo1.ListIndex = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Command2_Click

End Sub
