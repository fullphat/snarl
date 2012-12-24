VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fake Printer Monitor"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
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
   ScaleHeight     =   141
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1860
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "123456"
      Top             =   960
      Width           =   1875
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4620
      Top             =   840
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":000C
      Left            =   1860
      List            =   "Form1.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   780
      Top             =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Now"
      Default         =   -1  'True
      Height          =   435
      Left            =   1860
      TabIndex        =   1
      Top             =   1500
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   1020
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Check frequency:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Server:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mIconData As String

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Private Sub Combo1_Click()

    Select Case Combo1.ListIndex
    Case 0
        Timer2.Enabled = False

    Case 1
        Timer2.Interval = (CInt(Rnd * 12) + 10) * 1000
        Debug.Print Timer2.Interval
        Timer2.Enabled = True

    Case 2
        Timer2.Interval = (CInt(Rnd * 12) + 20) * 1000
        Debug.Print Timer2.Interval
        Timer2.Enabled = True

    End Select

End Sub

Private Sub Command1_Click()

    uRunCheck

End Sub

Private Sub Form_Load()

    uEncodeIcon mIconData

    Combo1.ListIndex = 0
    Me.Caption = App.Title

End Sub

Private Sub theSocket_OnConnect()
Dim sz As String
Dim i As Integer

    Timer1.Enabled = False

    sz = "SNP/3.0" & vbCrLf & _
         "register?app-sig=" & App.ProductName & "&title=" & App.Title & "&icon-phat64=" & mIconData & vbCrLf & _
         "notify?app-sig=" & App.ProductName & "&title="

    Randomize Timer
    i = (Rnd * 3)

    Debug.Print i

    Select Case i
    Case 0
        sz = sz & "Paper low&icon=!system-warning"

    Case 1
        sz = sz & "Black ink low&icon=!system-info"

    Case 2
        sz = sz & "Colour ink low&icon=!system-info"

    Case 3
        sz = sz & "Paper jam&icon=!system-critical"

    End Select

    sz = sz & "&text=PhatJet 3000X on " & g_GetComputerName() & vbCrLf & _
         "END" & vbCrLf

    theSocket.SendData sz

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim sz As String

    theSocket.GetData sz
    Debug.Print sz

    theSocket.CloseSocket
    uReset

End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False

    If theSocket.State <> sckConnected Then
        MsgBox "Couldn't contact remote server", vbExclamation Or vbOKOnly, App.Title
        uReset

    End If

End Sub

Private Sub uRunCheck()

    If (theSocket Is Nothing) Then
        Command1.Enabled = False
        Set theSocket = New CSocket
        theSocket.Connect Text1.Text, 9887
        Timer1.Enabled = True

    End If

End Sub

Private Sub uReset()

    Set theSocket = Nothing
    Command1.Enabled = True

End Sub

Private Sub Timer2_Timer()

    uRunCheck

End Sub

Private Function uEncodeIcon(ByRef Base64 As String) As Boolean
Dim sz As String
Dim i As Integer

    On Error Resume Next

    i = FreeFile()

    Err.Clear
    Open g_MakePath(App.Path) & "icon.png" For Binary Access Read Lock Write As #i
    If Err.Number = 0 Then
        sz = String$(LOF(i), Chr$(0))
        Get #i, , sz
        Close #i

        sz = Encode64orig(sz)                   ' // encode as standard Base64
        If sz <> "" Then
            Base64 = Replace$(sz, vbCrLf, "#")  ' // replace CRLFs
            Base64 = Replace$(Base64, "=", "%")
            uEncodeIcon = True

        End If

    Else
        g_Debug "uEncodeIcon(): " & Err.Description, LEMON_LEVEL_CRITICAL
    
    End If

End Function

