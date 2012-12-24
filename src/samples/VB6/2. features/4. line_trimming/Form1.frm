VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Line Trimming"
   ClientHeight    =   1515
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
   ScaleHeight     =   1515
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Text            =   "your message here"
      Top             =   360
      Width           =   4395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Chat Message"
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
                    "&uid=12345" & _
                    "&title=Some Body" & _
                    "&text=" & Text2.Text & _
                    "&icon=" & App.Path & "\person.png" & _
                    "&merge-uid=12345"

    Else
        MsgBox "Error registering with Snarl.  Check Snarl is running.", vbExclamation Or vbOKOnly, "Snarl Sample"

    End If

'    If sn41IsVisible(mMsgToken) <> 0 Then
'        ' /* current notification is still visible so simply update to that one
'        '    special R2.4 feature: escape character '\+' represents the current
'        '    notification title or text so the following appends a CRLF and our
'        '    line of text to the current notification text */
'        sn41EZUpdate mMsgToken, , "\+\n" & Text2.Text
'
'    Else
'        ' /* create a new one */
'        mMsgToken = sn41EZNotify(mAppToken, "", "Some Body", Text2.Text, 0, App.Path & "\person.png")
'
'    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    snDoRequest "unreg?app-sig=" & App.ProductName

End Sub

