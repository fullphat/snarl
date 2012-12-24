VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
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
   ScaleHeight     =   3090
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   7515
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":000C
      Top             =   360
      Width           =   7515
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Request"
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

'    If mMsg Then _
        sn41EZUpdate mMsg, , Text2.Text

End Sub

Private Sub Command4_Click()
Dim hr As Long

'    hr = snDoRequest("register?app-sig=" & App.ProductName & "&title=" & App.Title & "&icon=" & App.Path & "\icon.png")
'    If hr < 0 Then
'        uPrint "Error registering with Snarl: " & Abs(hr)
'
'    Else
        hr = snDoRequest(Text2.Text)
        uPrint "Response from Snarl: " & CStr(hr)

'    End If

End Sub

Private Sub Form_Load()

    Text2.Text = ""
    Text1.Text = ""

    g_SetWindowIconToAppResourceIcon2 Me.hWnd
    Me.Caption = App.Title
    Me.Show
    
    With Text2
        .SelStart = 0
        .SelLength = Len(.Text)

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim hr As Long
'
'    hr = snDoRequest("unregister?app-sig=" & App.ProductName)
'    If hr < 0 Then
'        Debug.Print "FAILED: " & Abs(hr)
'
'    Else
'        Debug.Print "OK"
'
'    End If

End Sub

Private Sub uPrint(ByVal Text As String)

    With Text1
        .Text = .Text & Text & vbCrLf
        .SelLength = 0
        .SelStart = Len(.Text) - 2

    End With

End Sub
