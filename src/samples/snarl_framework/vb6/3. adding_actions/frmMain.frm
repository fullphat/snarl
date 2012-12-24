VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding Actions"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   4515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   4515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents myApp As SnarlApp
Attribute myApp.VB_VarHelpID = -1

Private Sub Command1_Click()

    ' /* create notification */

Dim pn As Notification

    Set pn = New Notification
    With pn
        .Title = "Take action!"
        .Text = "It's your choice - pick from the menu..."
        .UID = "123456"
        .Class = "simpleclass"

    End With

Dim pa As Actions

    Set pa = New Actions
    With pa
        .Add "Yes", "choice:yes"
        .Add "No", "selection/no"
        .Add "Maybe", "action=maybe"

    End With
    
    pn.Actions = pa

    ' /* go */

    myApp.Show pn

End Sub

Private Sub Form_Load()

    ' /* create the app object */

    Set myApp = New SnarlApp

    With myApp
        .Signature = "test/sfw_adding_actions"
        .Title = "Adding Actions"

Dim pClasses As Classes

        ' /* define our notification classes */

        Set pClasses = New Classes
        With pClasses
            .Add "simpleclass", "A simple class"

        End With

        ' /* attach our classes to the app object */

        myApp.Classes = pClasses

    End With

    uUpdateStatus

End Sub

Private Sub Form_Unload(Cancel As Integer)

    myApp.TidyUp

End Sub

Private Sub uUpdateStatus()

    Label1.Caption = "Snarl is " & IIf(myApp.IsSnarlRunning(), "", "not ") & "running"

End Sub

Private Sub myApp_NotificationActionSelected(ByVal UID As String, ByVal Identifier As String)

    List1.AddItem Now() & ": you picked '" & Identifier & "'"
    List1.ListIndex = List1.ListCount - 1

    ' /* if "No" is selected, hide the notification */

    If Identifier = "selection/no" Then _
        myApp.Hide UID

End Sub

Private Sub myApp_NotificationClosed(ByVal UID As String)

    List1.AddItem Now() & ": notification closed"
    List1.ListIndex = List1.ListCount - 1

End Sub

Private Sub myApp_NotificationExpired(ByVal UID As String)

    List1.AddItem Now() & ": notification timed out"
    List1.ListIndex = List1.ListCount - 1

End Sub

Private Sub myApp_NotificationInvoked(ByVal UID As String)

    List1.AddItem Now() & ": notification clicked"
    List1.ListIndex = List1.ListCount - 1

End Sub
