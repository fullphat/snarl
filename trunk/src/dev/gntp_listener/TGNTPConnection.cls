VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TGNTPConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Public Sub Accept(ByVal Request As Long)

    Set theSocket = New CSocket
    theSocket.Accept Request

End Sub

Private Sub Class_Terminate()

    If Not (theSocket Is Nothing) Then _
        theSocket.CloseSocket

End Sub

Private Sub theSocket_OnClose()

    Debug.Print "== closed =="

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim ppd As BPackedData
Dim bKeepOpen As Boolean
Dim szReply As String
Dim sz As String

    theSocket.PeekData sz
    Debug.Print "data arrival: " & Asc(g_SafeMidStr(sz, Len(sz) - 3, 1)) & " " & Asc(g_SafeMidStr(sz, Len(sz) - 2, 1)) & " " & Asc(g_SafeMidStr(sz, Len(sz) - 1, 1)) & " " & Asc(g_SafeMidStr(sz, Len(sz) - 0, 1))

    If g_SafeRightStr(sz, 1) = Chr$(0) Then
        ' /* is it a flash permission request? */
        If LCase$(g_TrimStr(sz)) = "<policy-file-request/>" Then _
            theSocket.SendData "<?xml version=""1.0""?><!DOCTYPE cross-domain-policy SYSTEM ""/xml/dtds/cross-domain-policy.dtd""><cross-domain-policy><site-control permitted-cross-domain-policies=""master-only""/><allow-access-from domain=""*"" to-ports=""*"" /></cross-domain-policy>" & Chr$(0)

    ElseIf g_SafeRightStr(sz, 4) = vbCrLf & vbCrLf Then
        theSocket.GetData sz
        Form1.Output "==Start of raw packet=="
        Form1.Output Replace$(sz, vbCrLf, "¶" & vbCrLf)
        Form1.Output "==End of raw packet==" & vbCrLf

        gntp_Process sz, theSocket, szReply, bKeepOpen

        Form1.Output ""
        Form1.Output "==Start of GNTP reply=="
        Form1.Output Replace$(szReply, vbCrLf, "¶" & vbCrLf)
        Form1.Output "==End of GNTP reply==" & vbCrLf

        theSocket.SendData szReply

        If Not bKeepOpen Then _
            theSocket.CloseSocket

    End If

End Sub

