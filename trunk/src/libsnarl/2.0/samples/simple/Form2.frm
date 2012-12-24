VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snarl Framework Demo"
   ClientHeight    =   3090
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
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   1980
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   1035
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form2.frx":1042
      Top             =   840
      Width           =   4515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "Your notification title here"
      Top             =   480
      Width           =   4515
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Win32 API"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   2700
      Width           =   4515
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents myApp As SnarlApp
Attribute myApp.VB_VarHelpID = -1

Private Sub Command1_Click()

    ' /* create actions */

Dim pa As Actions

    Set pa = New Actions
    pa.Add "Action 1", "act1"
    pa.Add "Forward", "fwd"
    pa.Add "Quit", "x"

    ' /* create notification */

Dim pn As Notification

    Set pn = New Notification
    With pn
        .Title = Text1.Text
        .Text = Text2.Text
        .UID = "123456"
        .Actions = pa
        .Class = "1"

        .Add "value-percent", "100"
        .Add "foo-bar", "!"
        .Add "value-percent", "50", True
        .Remove "foo-bar"

    End With

    If Check1.Value = vbUnchecked Then _
        myApp.RemoteComputer = "127.0.0.1"

    ' /* go */

    Debug.Print myApp.Show(pn)

    Debug.Print myApp.IsVisible(pn.UID) & " - " & myApp.IsVisible("?")

End Sub

Private Sub Form_Load()

    ' /* create app */

    Set myApp = New SnarlApp

    Debug.Print myApp.IsSnarlInstalled()

    With myApp
        .Signature = "test/libmsnarl2xx023"
        .Title = "snarl.library 2 test"
'        .Icon = .MakePath(App.Path) & "icon.png"
'        .Hint = "Acme Products Present"
        .IsDaemon = True

    End With

Dim pClasses As Classes

    Set pClasses = New Classes
    pClasses.Add "1", "My class" ', , "Default title", "Default text"
    myApp.Classes = pClasses

    uUpdateStatus

End Sub

Private Sub Form_Unload(Cancel As Integer)

    myApp.TidyUp

End Sub

Private Sub myApp_Activated()

    Me.Show

End Sub

Private Sub myApp_NotificationActionSelected(ByVal UID As String, ByVal Identifier As String)

    Debug.Print "## action '" & Identifier & "' from " & UID & " selected ##"

    If Identifier = "x" Then _
        myApp.Hide UID

End Sub

Private Sub myApp_NotificationClosed(ByVal UID As String)

    Debug.Print "## CLOSED: " & UID

End Sub

Private Sub myApp_NotificationInvoked(ByVal UID As String)

    Debug.Print "## INVOKED: " & UID

End Sub

Private Sub myApp_Quit()

    Unload Me

End Sub

Private Sub myApp_ShowAbout()

    MsgBox "About..."

End Sub

Private Sub myApp_ShowConfig()

    MsgBox "Config..."
    
End Sub

Private Sub myApp_SnarlLaunched()

    uUpdateStatus

End Sub

Private Sub myApp_SnarlQuit()

    uUpdateStatus

End Sub

Private Sub myApp_SnarlStarted()

    uUpdateStatus

End Sub

Private Sub myApp_SnarlStopped()

    uUpdateStatus

End Sub

Private Sub uUpdateStatus()

    Label1.Caption = "Snarl is " & IIf(myApp.IsSnarlRunning(), "", "not ") & "running"

End Sub

Private Sub myApp_UserAway()

    Me.Caption = "User went away"

End Sub

Private Sub myApp_UserReturned()

    Me.Caption = "User came back"

End Sub
