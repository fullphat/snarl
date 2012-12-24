VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
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
   ScaleHeight     =   735
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Test"
      Default         =   -1  'True
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1455
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

Private Sub Command4_Click()

    If mToken Then _
        mMsg = sn41EZNotify(mToken, "", "Test Notification", "This should be queued..." & vbCrLf & CStr(Now), 0, App.Path & "\icon.png", -1)

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
