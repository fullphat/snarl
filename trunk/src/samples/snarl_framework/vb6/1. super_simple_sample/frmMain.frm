VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Simple Sample"
   ClientHeight    =   1215
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
   ScaleHeight     =   1215
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
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

Dim myApp As SnarlApp
Attribute myApp.VB_VarHelpID = -1

Private Sub Command1_Click()
Dim pa As Actions

    ' /* create notification */

Dim pn As Notification

    Set pa = New Actions
    pa.Add "Open", "@1"
    pa.Add "Run disk cleanup...", "@2"
    

    Set pn = New Notification
    With pn
        .Title = "Low disk space"
        .Text = "MOVIES (g:\) is nearly out of space\nCapacity: 320.0GB, free: 15.1GB"
        .UID = "123456"
        .Class = "simpleclass"
        .Actions = pa
        .DefaultCallback = "!open"

    End With

    ' /* go */

    myApp.Show pn

End Sub

Private Sub Form_Load()

    ' /* create the app object */

    Set myApp = New SnarlApp

    With myApp
        .Signature = "test/sfw_super_simple_sample"
        .Title = "DiskMonitor"
        .Hint = "A very simple (but super) sample of the Snarl Framework"
        .Icon = "!disk-low_space"

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

