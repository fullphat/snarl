VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
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
   ScaleHeight     =   3570
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Pick Icon"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   3000
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pick Image"
      Height          =   495
      Left            =   60
      TabIndex        =   7
      Top             =   3000
      Width           =   1635
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":628A
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Notification Title"
      Top             =   300
      Width           =   5415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   495
      Left            =   4020
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Icon (leave blank to use Form icon)"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   720
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

Private Declare Function SHChangeIconDialog Lib "shell32" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As String, ByVal Reserved As Long, lpIconIndex As Long) As Long
'Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Dim mAppToken As Long

Private Sub Command1_Click()

    With New CFileDialog
        .DialogType = E_DIALOG_OPEN
        .Title = "Pick image file..."

        If Text3.Text = "" Then
            .InitialPath = App.Path
            
        Else
            .InitialPath = Text3.Text
            
        End If

        If .Go(False, E_FILE_DIALOG_CENTRE_SCREEN) Then
            Text3.Text = .SelectedFile


        End If

    End With

End Sub

Private Sub Command2_Click()
Dim sz As String
Dim dw As Long

    sz = String$(260, 0)
    If SHChangeIconDialog(Me.hWnd, sz, 260, dw) <> 0 Then _
        Text3.Text = g_TrimStr(StrConv(sz, vbFromUnicode)) & ",-" & CStr(dw + 1)

    Text3.Text = Replace$(LCase$(Text3.Text), "%systemroot%", uSysDir())

End Sub

Private Sub Command4_Click()
Dim sz As String
Dim hr As Long

    hr = snarl_register(App.ProductName, App.Title, App.Path & "\icon.png")
    If hr < 0 Then
        MsgBox "Error registering with Snarl (" & CStr(Abs(hr)) & ")", vbExclamation Or vbOKOnly, App.Title

    Else
        If Text3.Text = "" Then
            sz = "%" & Me.Icon.Handle

        Else
            sz = Text3.Text

        End If

        snarl_ez_notify App.ProductName, "", Text1.Text, Text2.Text, sz

    End If

End Sub

Private Sub Form_Load()

    Me.Caption = App.Title

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Debug.Print snarl_unregister(App.ProductName)

End Sub

Private Function uSysDir() As String
Dim sz As String
Dim i As Long

    sz = String$(1024, 0)
    i = GetWindowsDirectory(sz, Len(sz))
    If i > 0 Then _
        uSysDir = Left$(sz, i)

End Function
