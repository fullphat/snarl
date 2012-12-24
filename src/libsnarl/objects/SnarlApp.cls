VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SnarlApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const CLASS_NAME = "libsnarl_app_callback"

Private Type T_INFO
    ' /* set during SetTo() */
    RemoteHost As String
    Signature As String
    Title As String
    Icon As String
    Password As String
    Classes As NotificationClasses

    ' /* set during Register() */
    zToken As Long

End Type

Dim mCached As T_INFO

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Dim mhWnd As Long                       ' // if registered using Win api
Dim WithEvents theSocket As CSocket     ' // if registered using SNP
Attribute theSocket.VB_VarHelpID = -1

Public Event Invoked(ByVal Notification As Long)
Public Event Expired(ByVal Notification As Long)
Public Event ActionSelected(ByVal Notification As Long, ByVal Identifier As String)
Public Event MenuSelected(ByVal Notification As Long, ByVal Index As Long)

Public Event SnarlLaunched()
Public Event SnarlQuit()
Public Event SnarlStarted()
Public Event SnarlStopped()

Dim mWaitReply As Boolean
Dim mResponse() As String               ' // 0-based SNP response received

Implements BWndProcSink

Private Sub Class_Initialize()
'//
End Sub

Private Sub Class_Terminate()

    Debug.Print "SnarlApp.Terminate"

'    Unregister

    If Not (theSocket Is Nothing) Then
        Debug.Print "closing socket..."
        mWaitReply = False
        theSocket.CloseSocket
        Set theSocket = Nothing

    ElseIf mhWnd <> 0 Then
        EZRemoveWindow mhWnd
        EZUnregisterClass CLASS_NAME

    End If

    Debug.Print "SnarlApp.Terminate complete"

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    If uMsg = &H41F Then
        Select Case LoWord(wParam)
        Case SNARL_CALLBACK_MENU_SELECTED
            RaiseEvent MenuSelected(lParam, HiWord(wParam))

        Case SNARL_CALLBACK_M_CLICK
            Debug.Print "middle button"

        Case SNARL_CALLBACK_INVOKED
            RaiseEvent Invoked(lParam)

        Case SNARL_CALLBACK_R_CLICK
            Debug.Print "right click"

        Case SNARL_CALLBACK_TIMED_OUT
            RaiseEvent Expired(lParam)

        Case SNARL_NOTIFY_ACTION
            RaiseEvent ActionSelected(lParam, CStr(HiWord(wParam)))

        Case Else
            Debug.Print g_HexStr(wParam)

        End Select

    ElseIf uMsg = snBroadcastMsg() Then
        Select Case wParam
        Case SNARL_BROADCAST_LAUNCHED
            RaiseEvent SnarlLaunched

'            If (mSig <> "") And (mTitle <> "") Then _
                mToken = sn41RegisterApp(mSig, mTitle, mIcon, mhWnd, &H41F)

        Case SNARL_BROADCAST_QUIT
            RaiseEvent SnarlQuit

        Case SNARL_BROADCAST_STARTED
            RaiseEvent SnarlStarted

        Case SNARL_BROADCAST_STOPPED
            RaiseEvent SnarlStopped

        End Select

    End If

End Function

Public Function SetTo(ByVal RemoteHost As String, ByVal Signature As String, ByVal Title As String, ByVal Icon As String, ByRef Classes As NotificationClasses, Optional ByVal Password As String) As S_STATUS_CODE

    If (Not (theSocket Is Nothing)) Or (mhWnd <> 0) Then
        Debug.Print "SnarlApp.SetTo(): either socket or window created"
        SetTo = S_ERROR_ACCESS_DENIED

    Else

        With mCached
            .RemoteHost = RemoteHost
            .Signature = Signature
            .Title = Title
            Set .Classes = Classes
            .Password = Password
            .Icon = Icon

        End With

        SetTo = S_SUCCESS

    End If

End Function

Public Function Register() As S_STATUS_CODE

    If mCached.RemoteHost <> "" Then
        Register = uRegisterRemote()

    Else
        Register = uRegisterLocal()

    End If

End Function

Private Function uRegisterLocal() As S_STATUS_CODE

    ' /* return */

    If (mCached.Signature = "") Or (mCached.Title = "") Or (mCached.Classes Is Nothing) Then
        Debug.Print "SnarlApp.uRegisterLocal(): invalid args"
        uRegisterLocal = S_ERROR_ARG_MISSING
        Exit Function

    End If

    If mCached.Classes.Count = 0 Then
        Debug.Print "SnarlApp.uRegisterLocal(): invalid args"
        uRegisterLocal = S_ERROR_ARG_MISSING
        Exit Function

    End If

    If IsWindow(mhWnd) = 0 Then
        ' /* create our callback window */
        EZRegisterClass CLASS_NAME
        mhWnd = EZAddWindow(CLASS_NAME, Me)

    End If

    ' /* did window create ok? */

    If mhWnd = 0 Then
        Debug.Print "SnarlApp.uRegisterLocal(): couldn't create callback window"
        uRegisterLocal = S_ERROR_FAILED
        Exit Function

    End If

Dim sz As String

    sz = "register?app-sig=" & mCached.Signature & "&title=" & mCached.Title & "&reply-to=" & CStr(mhWnd) & "&reply=" & CStr(&H41F)

    If mCached.Icon <> "" Then _
        sz = sz & "&icon=" & mCached.Icon

    If mCached.Password <> "" Then _
        sz = sz & "&password=" & mCached.Password

    mCached.zToken = snDoRequest(sz)
    If mCached.zToken <= 0 Then
        ' /* failed */
        uRegisterLocal = Abs(mCached.zToken)
        Debug.Print "SnarlApp.uRegisterLocal(): failed: " & CStr(uRegisterLocal)
        EZRemoveWindow mhWnd
        EZUnregisterClass CLASS_NAME
        Exit Function

    End If

    ' /* success */

Dim i As Long

    With mCached.Classes
        Debug.Print "adding class(es)..."
        For i = 1 To .Count
            snDoRequest .AsRequest(i, mCached.zToken, mCached.Password)

        Next i

    End With

    Debug.Print "SnarlApp.uRegisterLocal(): token=" & CStr(mCached.zToken)

    uRegisterLocal = S_SUCCESS

End Function

Private Function uRegisterRemote() As S_STATUS_CODE

    ' /* return true if registered (or re-registered) successfully */

    If (mCached.Signature = "") Or (mCached.Title = "") Or (mCached.Classes Is Nothing) Then
        Debug.Print "SnarlApp.uRegisterRemote(): invalid args"
        uRegisterRemote = S_ERROR_ARG_MISSING
        Exit Function

    End If

    ' /* already registered using Win API? */

'    If IsWindow(mhWnd) <> 0 Then
'        Debug.Print "SnarlApp.RegisterRemote(): can't re-use Win32 connection"
'        Exit Function
'
'    End If

    If Me.IsConnected Then
        theSocket.CloseSocket

    ElseIf (theSocket Is Nothing) Then
        Set theSocket = New CSocket

    End If

    theSocket.Connect mCached.RemoteHost, 9887

Dim t As Long

    ' /* wait for a connection or time-out, whichever comes first */

    t = GetTickCount()
    Do While theSocket.State <> sckConnected
        Sleep 1
        DoEvents
        If Abs(GetTickCount() - t) > 3000 Then
            Debug.Print "SnarlApp.uRegisterRemote(): timed-out connecting to " & mCached.RemoteHost
            uRegisterRemote = S_ERROR_TIMED_OUT
            Exit Function

        End If

    Loop

    Debug.Print "SnarlApp.uRegisterRemote(): registering with Snarl..."

Dim i As Long

Dim sz As String

    sz = "reg?app-sig=" & mCached.Signature & "&title=" & mCached.Title & IIf(mCached.Icon <> "", "&icon=" & mCached.Icon, "")

    If mCached.Password <> "" Then _
        sz = sz & "&password=" & mCached.Password

Dim pToken As String

    ' /* send the request */

    uRegisterRemote = uSendSNP(sz, pToken)

    If uRegisterRemote = S_SUCCESS Then
        mCached.zToken = Val(pToken)
        Debug.Print "ok: token=" & mCached.zToken

        With mCached.Classes
            Debug.Print "adding class(es)..."
            For i = 1 To .Count
                uSendSNP .AsRequest(i, mCached.zToken, mCached.Password)

            Next i

        End With
    End If

End Function

Public Sub Unregister()
Dim sz As String

    Debug.Print "SnarlApp.Unregister"

    If mCached.Signature <> "" Then

        sz = "unreg?app-sig=" & CStr(mCached.Signature)
        If mCached.Password <> "" Then _
            sz = sz & "&password=" & mCached.Password

        If Not (theSocket Is Nothing) Then
            ' /* via SNP */
            Debug.Print "unregistering via SNP..."
            uSendSNP sz

            ' /* close the socket? */

        Else
            Debug.Print "unregistering via Win32..."
            snDoRequest sz

            EZRemoveWindow mhWnd
            EZUnregisterClass CLASS_NAME

        End If

        mCached.zToken = 0

    Else
        Debug.Print "not initialised"

    End If

    Debug.Print "SnarlApp.Unregister done"

End Sub

'Public Function Token() As Long
'
'    Token = mCached.zToken
'
'End Function

Public Function Notify(ByVal Class As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Icon As String, Optional ByVal Duration As Long = -1, Optional ByVal DefaultCallback As String, Optional ByRef Actions As NotificationActions) As S_STATUS_CODE

    If mCached.Signature = "" Then
        Debug.Print "not initialised"
        Exit Function

    End If

'    If (Class <> "") And (Not mCached.Classes.bHasClass(Class)) Then
'        Debug.Print "Notify: class '" & Class & "' not found"
'        Exit Function
'
'    End If

Dim sz As String

    sz = "notify?app-sig=" & mCached.Signature & "&id=" & Class

    If Title <> "" Then _
        sz = sz & "&title=" & Title

    If Text <> "" Then _
        sz = sz & "&text=" & Text

    If Icon <> "" Then _
        sz = sz & "&icon=" & Icon

    If Duration <> -1 Then _
        sz = sz & "&timeout=" & CStr(Duration)

    If DefaultCallback <> "" Then _
        sz = sz & "&callback=" & DefaultCallback

    If mCached.Password <> "" Then _
        sz = sz & "&password=" & mCached.Password

Dim pToken As String
Dim hr As Long

    If Not (theSocket Is Nothing) Then
        ' /* via SNP */
        If uSendSNP(sz, pToken) = 0 Then _
            hr = g_SafeLong(pToken)

    Else
        hr = snDoRequest(sz)

    End If

    ' /* if it succeeded and we have actions, add them now */

Dim i As Long

    If Not (Actions Is Nothing) Then
        With Actions
            Debug.Print "adding action(s)..."
            For i = 1 To .Count
                sz = .AsRequest(i, mCached.zToken, mCached.Password)

                If Not (theSocket Is Nothing) Then
                    uSendSNP sz
                    
                Else
                    snDoRequest sz

                End If

            Next i

        End With

    End If

End Function

Private Function uSendSNP(ByVal Data As String, Optional ByRef Result As String) As S_STATUS_CODE

    ' /* returns Snarl status code and (optionally) result field, if there is one */

    If (theSocket Is Nothing) Then
        Debug.Print "uSendSNP(): socket not created"
        uSendSNP = SNARL_ERROR_BAD_SOCKET
        Exit Function

    End If

    If Not theSocket.State = sckConnected Then
        Debug.Print "uSendSNP(): socket not connected"
        uSendSNP = SNARL_ERROR_BAD_SOCKET
        Exit Function

    End If

    Debug.Print "sending '" & Data & "'..."
    mWaitReply = True
    theSocket.SendData "snp://" & Data & Chr$(13)

    ' /* sync wait loop */

Dim t As Long

    t = GetTickCount()
    Do While mWaitReply
        DoEvents
        Sleep 1
        If Abs(GetTickCount() - t) > 3000 Then
            Debug.Print "SnarlApp.uSendSNP(): timed out"
            uSendSNP = SNARL_ERROR_TIMED_OUT
            Exit Function

        End If

    Loop

    Debug.Print "uSendSNP(): reply received"

    If UBound(mResponse) > 3 Then _
        Result = mResponse(4)

    uSendSNP = mResponse(2)

End Function

Public Function IsConnected() As Boolean

    If Not (theSocket Is Nothing) Then
        IsConnected = (theSocket.State = sckConnected)

    Else


    End If

End Function

Private Sub theSocket_OnClose()

    ' /* could mean that the remote Snarl instance has quit... */

    Debug.Print "SnarlApp: socket closed"

'    RaiseEvent SnarlQuit

End Sub

Private Sub theSocket_OnConnect()

    Debug.Print "SnarlApp: connected to " & theSocket.RemoteHostIP & " / " & theSocket.RemoteHost
'    uRemoteReg

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim sz As String

    theSocket.GetData sz

    ' /* need to decode here - if it's a callback we don't break the wait loop */

    mResponse = Split(sz, "/")

    Debug.Print "#DataArrival: data=" & sz & " ubound=" & UBound(mResponse)

    If UBound(mResponse) < 3 Then
        Debug.Print "#DataArrival: invalid response ignored"
        Exit Sub

    End If

    ' /* check status code - if a 3xx then raise appropriate event here */

    Select Case mResponse(2)
    Case SNARL_NOTIFY_ACTION
        RaiseEvent ActionSelected(g_SafeLong(mResponse(3)), mResponse(4))
        Exit Sub

    Case SNARL_NOTIFY_EXPIRED
        RaiseEvent Expired(g_SafeLong(mResponse(3)))
        Exit Sub

    Case SNARL_NOTIFY_INVOKED
        RaiseEvent Invoked(g_SafeLong(mResponse(3)))
        Exit Sub

    Case Is > 300
        Debug.Print "#DataArrival: invalid async callback " & mResponse(2)
        Exit Sub

    End Select

    ' /* are we waiting for a reply? if so, break the loop */

    If mWaitReply Then _
        mWaitReply = False

End Sub

Private Sub theSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    Debug.Print "SnarlApp: '" & Description & "'"

End Sub

Private Sub theSocket_OnSendComplete()

    Debug.Print "send complete"

End Sub

Private Function uEncodeFile(ByVal Path As String) As String
Dim i As Integer

    On Error Resume Next

    i = FreeFile()

    Err.Clear
    Open Path For Binary Access Read Lock Write As #i
    If Err.Number <> 0 Then
        Debug.Print "uEncodeFile(): couldn't open '" & Path & "'"
        Exit Function

    End If

    If LOF(i) = 0 Then
        Debug.Print "uEncodeFile(): '" & Path & "' is empty"
        Close #i
        Exit Function

    End If
    
Dim sz As String

    sz = String$(LOF(i), Chr$(0))
    Get #i, , sz
    uEncodeFile = Replace$(Encode64(sz), "=", "%")          ' // must replace end marker with something other than '='
    Close #i

End Function

Public Function DoRequest(ByVal Request As String, Optional ByVal ReplyTimeout As Long = 1000) As Long

    If mCached.Signature = "" Then
        Debug.Print "SnarlApp.DoRequest(): not initialised"
        Exit Function

    End If

    If Not (theSocket Is Nothing) Then
        DoRequest = uSendSNP(Request)

    Else
        DoRequest = snDoRequest(Request, ReplyTimeout)

    End If

End Function






