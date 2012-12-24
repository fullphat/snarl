VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TForwarder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TForwarder.cls
    '/
    '/  Description:    Represents an instance of a single notification forwarded to a single destination
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mSubscriber As TSubscriber
Dim mContent As String
Dim mDest As String
Dim mGuid As String

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mGuid

End Function

Private Function BTagItem_Value() As String
End Function

Public Function Init(ByVal Guid As String, ByRef Subscriber As TSubscriber, ByVal Content As String, Optional ByVal IsGNTP As Boolean = False) As Boolean

    mGuid = Guid
    mContent = Content
    mDest = Subscriber.Destination

    Set theSocket = New CSocket
    If theSocket.Connect(Subscriber.Destination, IIf(IsGNTP, GNTP_DEFAULT_PORT, SNP_DEFAULT_PORT)) Then
        g_Debug "TForwarder.Init(): connecting to '" & Subscriber.Destination & "'..."
        Set mSubscriber = Subscriber
        Set theTimer = new_BTimer(5000, True, True)             ' // should probably be configurable via snarl41.config...
        Init = True

    Else
        g_Debug "TForwarder.Init(): invalid destination address '" & Subscriber.Destination & "'", LEMON_LEVEL_CRITICAL
        Set theSocket = Nothing

    End If

End Function

Private Sub theSocket_OnConnect()

    g_Debug "TForwarder.OnConnect(): connected to '" & mDest & "'", LEMON_LEVEL_INFO
    Set theTimer = Nothing
    theSocket.SendData mContent

End Sub

Private Sub theSocket_OnSendComplete()

    g_Debug "TForwarder.OnSendComplete(): content sent to '" & mDest & "'", LEMON_LEVEL_INFO
    theSocket.CloseSocket
    Set theSocket = Nothing
    mSubscriber.RemoveForward mGuid

End Sub

Private Sub theTimer_Pulse()

    g_Debug "TForwarder.Pulse(): timed out connecting to '" & mDest & "'", LEMON_LEVEL_WARNING
    theSocket.CloseSocket
    Set theSocket = Nothing
    mSubscriber.RemoveForward mGuid

End Sub
