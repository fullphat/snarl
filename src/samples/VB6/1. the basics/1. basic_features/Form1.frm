VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   1035
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
   ScaleHeight     =   1035
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Test!"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   4395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click()
Dim hr As Long

    hr = snarl_register(App.ProductName, App.Title, App.Path & "\icon.png")
    If hr < 0 Then
        Label3.Caption = "Error registering with Snarl (" & Abs(hr) & ")"

    Else
        Label3.Caption = "Registered with Snarl V" & CStr(snDoRequest("version")) & " (" & Hex$(hr) & ")"
        snarl_add_class App.ProductName, "1", "Enabled", True
        snarl_add_class App.ProductName, "2", "Disabled", False
        snarl_add_class App.ProductName, "3", "Sticky", True, , , , , 0
        snarl_add_class App.ProductName, "4", "Google", True, , , , , , , "http://www.google.com/"

        snarl_ez_notify App.ProductName, "1", "Hello, world!", "You should see this notification"
        snarl_ez_notify App.ProductName, "2", "Hello, world!", "You should see not this notification by default"
        snarl_ez_notify App.ProductName, "3", "Hello, world!", "This notification is sticky by default"
        snarl_ez_notify App.ProductName, "4", "Hello, world!", "Click me to go to Google!"

    End If

End Sub

Private Sub Form_Load()

    Me.Caption = App.Title

End Sub

Private Sub Form_Unload(Cancel As Integer)

    snarl_unregister App.ProductName

End Sub
