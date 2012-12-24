VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "audiomon_destination_test"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
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
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      LargeChange     =   10
      Left            =   1200
      Max             =   100
      TabIndex        =   2
      Top             =   2100
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mute"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2100
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1875
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim mhWnd As Long

Private Const CLASS_NAME = "w>audiomon"

Implements BWndProcSink

Private Sub uAddText(ByVal Text As String)

    With Text1
        .Text = .Text & Text & vbCrLf
        .SelLength = 0
        .SelStart = Len(.Text) - 1

    End With

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    Select Case uMsg
    Case &H440
        uAddText "0x0440: " & g_HexStr(wParam) & " " & g_HexStr(lParam)
        
        Select Case wParam
        Case 0
            uAddText "MUTE: " & g_HexStr(lParam)

        Case 1
            uAddText "VOL:" & g_HexStr(lParam) & " (" & CStr(lParam) & ")"

        End Select

    End Select

End Function

Private Sub Command1_Click()
Dim i As Long

    i = Rnd * 100

    SendMessage mhWnd, &H440, 1, ByVal i

End Sub

Private Sub Check1_Click()

    If Check1.Value = vbChecked Then
        SendMessage

End Sub

Private Sub Form_Load()

    If Not EZRegisterClass(CLASS_NAME) Then
        MsgBox "error creating window class", vbCritical, ""
        Unload Me
        Exit Sub

    End If

    mhWnd = EZAddWindow(CLASS_NAME, Me, CLASS_NAME)
    If mhWnd = 0 Then
        MsgBox "error creating listening window", vbCritical, ""
        EZUnregisterClass CLASS_NAME
        Unload Me
        Exit Sub

    End If

    uAddText "Window created: handle is " & g_HexStr(mhWnd)
    uAddText "Waiting..."

End Sub

Private Sub Form_Unload(Cancel As Integer)

    EZRemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub



