VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "libmsnarl test"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
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
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   555
      Left            =   3720
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   6780
      TabIndex        =   3
      Top             =   60
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   555
      Left            =   1860
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

'Private Type T_DEST
'    Name As String              ' // or ip address
'    Connection As SnarlApp
'
'End Type

Dim mDest() As TConnection
Dim mCount As Long

Dim mNID As Long

Private Sub Command1_Click()

    If mCount > 0 Then _
        mDest(1).Notify

'    List1.AddItem CStr(mNID)

End Sub

Private Sub Command2_Click()

'    myApp.SetTitle mNID, CStr(Now) & " " & CStr(Now) & " " & CStr(Now)

    Debug.Print myApp.IsConnected

End Sub

Private Sub Command3_Click()
Dim sz As String

    sz = InputBox("Enter the IP address or DNS name of a server to send alerts to:", "Add Remote Server")

    If sz <> "" Then
        uAdd sz


    End If

End Sub

Private Sub Form_Load()

    Set mySnarl = get_snarl()

    With List1
        .AddItem "snarl running: " & is_snarl_running()
        .AddItem "version: " & snarl_version()
        .AddItem "path: " & get_etc_path()

        If is_snarl_running Then
'            .AddItem "test notification: " & CStr(mySnarl.SimpleNotify("", "Hello, World!", g_MakePath(App.Path) & "icon.png"))

'            Set myApp = New_SnarlApp("127.0.0.1")
'            myApp.SetTo "app/vnd.acme-test", "libmsnarl test", g_MakePath(App.Path) & "icon.png"
'            myApp.AddClass "class1", "Test Class #1"
'
'            .AddItem "test app: " & CStr(myApp.Token)

        End If

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

'    If Not (myApp Is Nothing) Then _
        myApp.Unset

Static i As Long

    If mCount Then
        For i = mCount To 1 Step -1
            mDest(i).Quit

        Next i

    End If

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

Private Sub uAdd(ByVal Destination As String)
Dim pc As TConnection

    Set pc = New TConnection
    If pc.Init(Destination) Then
        List2.AddItem Destination

        mCount = mCount + 1
        ReDim Preserve mDest(mCount)
        Set mDest(mCount) = pc

    Else
        MsgBox Destination & " is not responding", vbInformation

    End If

End Sub
