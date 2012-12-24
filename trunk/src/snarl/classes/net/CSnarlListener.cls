VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSnarlListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           CSnarlListener.cls
    '/
    '/  Description:    Represents a listening socket bound to TCP port 9887 (SNP), 9889 (JSON) or
    '/                  23053 (GNTP)
    '/
    '/  © 2010 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mConn() As TRemoteConnection
Dim mCount As Long

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Public Sub Go(ByVal PortNumber As Long)

    Set theSocket = New CSocket
    theSocket.Bind PortNumber
    theSocket.Listen

    g_Debug "CSnarlListener.Go(): bound to port " & theSocket.LocalPort

End Sub

Public Sub Quit()

    If Not (theSocket Is Nothing) Then
        g_Debug "CSnarlListener.Quit(): unbinding from port " & theSocket.LocalPort & "..."
        theSocket.CloseSocket
        Set theSocket = Nothing

    Else
        g_Debug "CSnarlListener.Quit(): socket object is invalid", LEMON_LEVEL_WARNING

    End If

    ReDim mConn(0)
    mCount = 0

End Sub

Private Sub theSocket_OnConnectionRequest(ByVal requestID As Long)

    g_Debug "CSnarlListener.OnConnectionRequest(): requestID=0x" & g_HexStr(requestID), LEMON_LEVEL_PROC

    mCount = mCount + 1
    ReDim Preserve mConn(mCount)
    Set mConn(mCount) = New TRemoteConnection
    mConn(mCount).Accept requestID, theSocket.LocalPort

End Sub
