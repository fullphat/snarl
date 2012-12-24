VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   4050
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
   ScaleHeight     =   4050
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":000C
      Left            =   60
      List            =   "Form1.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include icon"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   1740
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":004B
      Top             =   1080
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Lorem ipsum dolor sit amet"
      Top             =   360
      Width           =   4395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   3660
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
        sn41EZUpdate mMsg, , Text2.Text

End Sub

Private Sub Command4_Click()
Dim pri As Long

    Select Case Combo1.ListIndex
    Case 0
        pri = -1

    Case 1
        pri = 0

    Case 2
        pri = 1

    End Select

    If mToken Then _
        mMsg = sn41EZNotify(mToken, "", Text1.Text, Text2.Text, -1, _
                            IIf(Check1.Value = vbChecked, App.Path & "\icon.png", ""), pri)

    Label3.Caption = "token: " & CStr(mMsg) & " lasterror: " & CStr(sn41GetLastError())

End Sub

Private Sub Form_Load()
Dim hr As Long

    If Not sn41IsSnarlRunning() Then
        MsgBox "Snarl isn't running - launch Snarl, then run this demo.", vbExclamation Or vbOKOnly, App.Title
        Unload Me

    Else
        hr = sn41RegisterApp(App.ProductName, App.Title, App.Path & "\icon.png")
        If hr = 0 Then
            Me.Caption = "Error registering with Snarl: " & sn41GetLastError()

        Else
            Me.Caption = "Registered with Snarl V" & CStr(sn41GetVersion()) & " (" & Hex$(hr) & ")"
            mToken = hr

        End If

        Combo1.ListIndex = 1

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim hr As Long

    hr = sn41UnregisterApp(mToken)
    If hr = 0 Then
        Debug.Print "FAILED: " & sn41GetLastError()

    Else
        Debug.Print "OK: " & hr

    End If

End Sub
