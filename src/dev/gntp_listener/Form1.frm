VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GNTP Listener"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   8115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mIndent As Integer
Dim mConn() As TGNTPConnection
Dim mCount As Long

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Private Sub Form_Load()

    Set theSocket = New CSocket
    If Not theSocket.Bind(23053) Then
        MsgBox App.Title & " failed: socket already in use.  Ensure Snarl is not already listening for incoming" & vbCrLf & _
               "GNTP notifications and click Retry, or click Cancel to quit.", vbCritical Or vbOKOnly, "Can't start..."

        Unload Me
        Exit Sub

    End If

    theSocket.Listen

Dim sz As String

    sz = g_GetSystemFolderStr(CSIDL_DESKTOPDIRECTORY)
    If sz = "" Then
        Me.Output "couldn't locate desktop folder"

    Else
        l3OpenLog g_MakePath(sz) & "gntplistener.log"

    End If

    g_SetWindowIconToAppResourceIcon2 Me.hWnd

    Me.Output vbCrLf & App.ProductName & " " & App.Major & "." & App.Minor & " (" & App.Revision & ") initialised"
    Me.Output "Logging to '" & l3LogPath & "'" & vbCrLf

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

    theSocket.CloseSocket
    l3CloseLog

End Sub

Private Sub theSocket_OnConnectionRequest(ByVal requestID As Long)

    Me.Output vbCrLf & "[ Incoming Request ]" & vbCrLf

    mCount = mCount + 1
    ReDim Preserve mConn(mCount)
    Set mConn(mCount) = New TGNTPConnection
    mConn(mCount).Accept requestID

End Sub

Public Sub Output(ByVal Text As String)

    With Text1
        .Text = .Text & Space$(Abs(mIndent)) & Text & vbCrLf
        .SelLength = 0
        .SelStart = Len(.Text)

    End With

    g_Debug Text

End Sub

Public Sub Indent()

    mIndent = mIndent + 2
    Me.Output "{"
    
End Sub

Public Sub Outdent()

    Me.Output "}"
    mIndent = mIndent - 2

End Sub
