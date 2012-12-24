VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSNP3Req"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mGuid As String
Dim mApp As SnarlApp
Dim mContent As Notification

Dim mStuffToSend As String

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Public Sub SendData(ByVal Data As String, ByVal RemoteIP As String)

    If (Data = "") Or (RemoteIP = "") Then _
        Exit Sub

    Debug.Print "TSNP3Req.SendData(): connecting to '" & RemoteIP & "'..."

    mStuffToSend = Data
    Set theSocket = New CSocket
    theSocket.Connect RemoteIP, 9887

End Sub

Public Sub send(ByRef AppInfo As SnarlApp, ByRef Notification As Notification, ByVal RemoteIP As String)

    If (AppInfo Is Nothing) Or (Notification Is Nothing) Or (RemoteIP = "") Then _
        Exit Sub

    Debug.Print "app: " & AppInfo.Signature & " notification: " & Notification.Title & " dest: " & RemoteIP

    Set mApp = AppInfo
    Set mContent = Notification

    Set theSocket = New CSocket
    theSocket.Connect RemoteIP, 9887

End Sub

Private Sub Class_Initialize()

    mGuid = g_CreateGUID(True)

End Sub

''Public Function IsConnected() As Boolean
''
''    If Not (theSocket Is Nothing) Then
''        IsConnected = (theSocket.State = sckConnected)
''
''    Else
''
''
''    End If
''
''End Function

Private Sub theSocket_OnClose()

    ' /* could mean that the remote Snarl instance has quit... */

    Debug.Print "TSNP3Req: socket closed"

End Sub

Private Sub theSocket_OnConnect()

    Debug.Print "TSNP3Req: connected to " & theSocket.RemoteHostIP & " / " & theSocket.RemoteHost

    ' /* build the SNP3 request */

Dim sz As String

    If mStuffToSend <> "" Then
        sz = "SNP/3.0" & vbCrLf & mStuffToSend & vbCrLf & "END" & vbCrLf

    Else
        sz = "SNP/3.0" & vbCrLf
        sz = sz & mApp.bAsRequest(False)
        sz = sz & mContent.bAsRequest(mApp.Signature, mApp.bPassword) & vbCrLf
        sz = sz & "END" & vbCrLf

    End If

    Debug.Print "sending '" & Replace$(sz, vbCrLf, "+") & "'..."
    theSocket.SendData sz

End Sub

'Private Function uEncodeFile(ByVal Path As String) As String
'Dim i As Integer
'
'    On Error Resume Next
'
'    i = FreeFile()
'
'    Err.Clear
'    Open Path For Binary Access Read Lock Write As #i
'    If Err.Number <> 0 Then
'        Debug.Print "uEncodeFile(): couldn't open '" & Path & "'"
'        Exit Function
'
'    End If
'
'    If LOF(i) = 0 Then
'        Debug.Print "uEncodeFile(): '" & Path & "' is empty"
'        Close #i
'        Exit Function
'
'    End If
'
'Dim sz As String
'
'    sz = String$(LOF(i), Chr$(0))
'    Get #i, , sz
'    uEncodeFile = Replace$(Encode64(sz), "=", "%")          ' // must replace end marker with something other than '='
'    Close #i
'
'End Function

Public Function Guid() As String

    Guid = mGuid

End Function

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim sz As String

    theSocket.GetData sz
    Debug.Print theSocket.RemoteHost & ": " & sz
    theSocket.CloseSocket

End Sub

Private Sub theSocket_OnSendComplete()

    Debug.Print theSocket.RemoteHost & ": send complete"

End Sub
