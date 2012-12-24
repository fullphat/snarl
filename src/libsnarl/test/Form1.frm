VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "libmsnarl test"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   555
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents mySnarl As Snarl
Attribute mySnarl.VB_VarHelpID = -1
Dim WithEvents myApp As SnarlApp
Attribute myApp.VB_VarHelpID = -1

Dim mNID As Long

Private Sub Command1_Click()

'    mNID = myApp.EZNotify("class1", "Hello, world", "This is the registered class", , , , , "Item 1#?1|Item 2#?2||Item 3#?3")
    List1.AddItem CStr(mNID)

End Sub

Private Sub Command2_Click()

'    myApp.SetTitle mNID, CStr(Now) & " " & CStr(Now) & " " & CStr(Now)

End Sub

Private Sub Form_Load()

    Set mySnarl = get_snarl()

    With List1
        .AddItem "snarl running: " & is_snarl_running()
        .AddItem "version: " & snarl_version()
        .AddItem "path: " & get_etc_path()

'        If is_snarl_running Then
'            .AddItem "test notification: " & CStr(mySnarl.SimpleNotify("", "Hello, World!", g_MakePath(App.Path) & "icon.png"))
'
'            Set myApp = New SnarlApp
'            myApp.SetTo "app/vnd.acme-test", "libmsnarl test", g_MakePath(App.Path) & "icon.png"
'            myApp.AddClass "class1", "Test Class #1"
'
'            .AddItem "test app: " & CStr(myApp.Token)
'
'        End If

    End With

End Sub

Private Sub myApp_NotificationAcknowleged(ByVal Notification As Long)

    List1.AddItem "ACK: " & Notification

End Sub

Private Sub myApp_NotificationDismissed(ByVal Notification As Long)

    List1.AddItem "DISMISS: " & Notification

End Sub

Private Sub myApp_NotificationMenuSelected(ByVal Notification As Long, ByVal ItemIndex As Long)

    List1.AddItem "MENU: " & Notification & " item " & ItemIndex

End Sub

Private Sub myApp_SnarlStarted()

    List1.AddItem "## SNARL STARTED ##"

End Sub

Private Sub myApp_SnarlStopped()

    List1.AddItem "## SNARL STOPPED ##"

End Sub

Private Sub mySnarl_SnarlLaunched()

    List1.AddItem "## SNARL LAUNCHED ##"

End Sub

Private Sub mySnarl_SnarlQuit()

    List1.AddItem "## SNARL QUIT ##"

End Sub
