VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SystemSpy Preferences"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   502
      TabIndex        =   13
      Top             =   0
      Width           =   7530
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   2
         Height          =   570
         Left            =   2760
         Top             =   60
         Width           =   570
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmSettings.frx":000C
         Tag             =   "Windows"
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   720
         Picture         =   "frmSettings.frx":0C4E
         Tag             =   "Processes"
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   1320
         Picture         =   "frmSettings.frx":1890
         Tag             =   "Applications"
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   1920
         Picture         =   "frmSettings.frx":24D2
         Tag             =   "Folders"
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H007D9B7D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Index           =   2
      Left            =   180
      ScaleHeight     =   4395
      ScaleWidth      =   7095
      TabIndex        =   3
      Top             =   900
      Width           =   7095
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   30
         Top             =   1680
         Width           =   480
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   29
         Top             =   1680
         Width           =   480
      End
      Begin VB.ListBox app_include_list 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   7095
      End
      Begin VB.OptionButton app_mode 
         Caption         =   "Notify for only these applications:"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   27
         Top             =   60
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton app_mode 
         Caption         =   "Notify for all except these applications:"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   26
         Top             =   2220
         Width           =   3135
      End
      Begin VB.ListBox app_exclude_list 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   0
         TabIndex        =   25
         Top             =   2520
         Width           =   7095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   24
         Top             =   3840
         Width           =   480
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   23
         Top             =   3840
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H009A7950&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Index           =   0
      Left            =   180
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   1
      Top             =   900
      Width           =   7095
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1980
         Top             =   1020
      End
      Begin VB.CommandButton rem_window_spy 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   10
         Top             =   2760
         Width           =   480
      End
      Begin VB.CommandButton add_window_spy 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   9
         Top             =   2760
         Width           =   480
      End
      Begin VB.ListBox window_spy_list 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   0
         TabIndex        =   8
         Top             =   300
         Width           =   7095
      End
      Begin VB.Label Label1 
         Caption         =   "Notify for windows which match any of the following rules:"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   60
         Width           =   4275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4755
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   7335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00AAAAAA&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Index           =   3
      Left            =   180
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   4
      Top             =   900
      Width           =   7095
      Begin VB.ListBox folder_spy_list 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   0
         TabIndex        =   7
         Top             =   300
         Width           =   7095
      End
      Begin VB.CommandButton add_folder_spy 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   6
         Top             =   3780
         Width           =   480
      End
      Begin VB.CommandButton rem_folder_spy 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   5
         Top             =   3780
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "#"
         Height          =   615
         Left            =   0
         TabIndex        =   22
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Monitor changes in the following folders:"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   60
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7B88D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Index           =   1
      Left            =   180
      ScaleHeight     =   4395
      ScaleWidth      =   7095
      TabIndex        =   2
      Top             =   900
      Width           =   7095
      Begin VB.CommandButton add_exclude_process 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   21
         Top             =   3840
         Width           =   480
      End
      Begin VB.CommandButton rem_exclude_process 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   20
         Top             =   3840
         Width           =   480
      End
      Begin VB.ListBox process_exclude_list 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   0
         TabIndex        =   19
         Top             =   2520
         Width           =   7095
      End
      Begin VB.OptionButton process_mode 
         Caption         =   "Notify for all except these processes:"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   18
         Top             =   2220
         Width           =   3015
      End
      Begin VB.OptionButton process_mode 
         Caption         =   "Notify for only these processes:"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.ListBox process_include_list 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   7095
      End
      Begin VB.CommandButton add_include_process 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   15
         Top             =   1680
         Width           =   480
      End
      Begin VB.CommandButton rem_include_process 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   14
         Top             =   1680
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Go()

    zUpdateFolderWatchList
    zUpdateIncludeProcessList
    zUpdateExcludeProcessList
    zUpdateIncludeAppList
    zUpdateExcludeAppList

    process_mode(Form1.theProcessSpy.Mode).Value = True
    app_mode(Form1.theAppSpy.Mode).Value = True

    Me.Show

End Sub

Public Sub zUpdateFolderWatchList()
Dim pw As TFolderWatch

    folder_spy_list.Clear

    With Form1.theFolderSpy.List
        .Rewind
        Do While .GetNextTag(pw) = B_OK
            folder_spy_list.AddItem g_FormattedMidStr(pw.Path, 79)

        Loop

    End With

End Sub

Private Sub add_exclude_process_Click()
Dim sz As String

    sz = InputBox("Enter process name", "Add New Process")
    If sz <> "" Then
        If Form1.theProcessSpy.AddNewExcludeProcess(sz) Then
            Me.zUpdateExcludeProcessList

        End If
    End If

End Sub

Private Sub add_folder_spy_Click()
Dim sz As String

    sz = g_PickFolder(Me.hWnd, "Select folder to watch")
    If sz <> "" Then
        Form1.theFolderSpy.Add sz
        zUpdateFolderWatchList

    End If

End Sub

Private Sub add_include_process_Click()
Dim sz As String

    sz = InputBox("Enter process name", "Add New Process")
    If sz <> "" Then
        If Form1.theProcessSpy.AddNewIncludeProcess(sz) Then
            zUpdateIncludeProcessList

        End If
    End If

End Sub

Private Sub app_mode_Click(Index As Integer)

    app_include_list.Enabled = (Index = 0)
    app_exclude_list.Enabled = (Index = 1)

    Form1.theAppSpy.SetMode Index

End Sub

Private Sub Form_Load()
Dim pImg As Image

    For Each pImg In Image1
        pImg.ToolTipText = pImg.Tag

    Next

Dim pPic As PictureBox

    For Each pPic In Picture1
        pPic.BackColor = Me.BackColor

    Next

    Image1_Click 0

End Sub

Private Sub Image1_Click(Index As Integer)

    Picture1(Index).ZOrder 0
    Frame1.Caption = Image1(Index).Tag
    Shape2.Move Image1(Index).Left - 3, Image1(Index).Top - 3
        
End Sub

Private Sub process_mode_Click(Index As Integer)

    process_include_list.Enabled = (Index = 0)
    process_exclude_list.Enabled = (Index = 1)

    Form1.theProcessSpy.SetMode Index

End Sub

Private Sub rem_folder_spy_Click()

    If folder_spy_list.ListIndex > -1 Then
        Form1.theFolderSpy.Remove folder_spy_list.ListIndex + 1
        zUpdateFolderWatchList

    End If

End Sub

Public Sub UpdateWindowWatchList()
Dim pr As TRule

    window_spy_list.Clear

    With Form1.theWindowSpy.Rules
        .Rewind
        Do While .GetNextTag(pr) = B_OK
            window_spy_list.AddItem pr.Detail

        Loop

    End With

End Sub

Private Sub Timer1_Timer()
Static n As Integer
Static d As Integer

    If n <= 48 Then
        d = 16

    ElseIf n >= 255 Then
        d = -16

    End If

    n = n + d
    Shape2.BorderColor = RGB(n, 0, n)

End Sub

Public Sub zUpdateIncludeProcessList()

    process_include_list.Clear

Dim pe As CConfEntry

    With Form1.theProcessSpy.IncludeList
        .Rewind
        Do While .NextEntry(pe)
            process_include_list.AddItem pe.Name

        Loop

    End With

End Sub

Public Sub zUpdateExcludeProcessList()

    process_exclude_list.Clear

Dim pe As CConfEntry

    With Form1.theProcessSpy.ExcludeList
        .Rewind
        Do While .NextEntry(pe)
            process_exclude_list.AddItem pe.Name

        Loop

    End With

End Sub

Public Sub zUpdateIncludeAppList()

    app_include_list.Clear

Dim pe As CConfEntry

    With Form1.theAppSpy.IncludeList
        .Rewind
        Do While .NextEntry(pe)
            app_include_list.AddItem pe.Name

        Loop

    End With

End Sub

Public Sub zUpdateExcludeAppList()

    app_exclude_list.Clear

Dim pe As CConfEntry

    With Form1.theAppSpy.ExcludeList
        .Rewind
        Do While .NextEntry(pe)
            app_exclude_list.AddItem pe.Name

        Loop

    End With

End Sub



