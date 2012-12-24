VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TOutlookLogInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TWebForwardInstance.cls
    '/
    '/  Description:    Represents an active instance of a web forward
    '/
    '/  Created:        6-Jul-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage
End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Dim pOutlook As Outlook.Application

    If Not uGetOutlook(pOutlook) Then
        g_Debug "TOutlookLogInstance.UpdateContent(): couldn't get Outlook", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim pInbox As Outlook.Folder

    Set pInbox = pOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    If (pInbox Is Nothing) Then
        g_Debug "TOutlookLogInstance.UpdateContent(): couldn't get Inbox folder", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo NotificationInfo.Text

Dim szData As String
Dim szn As String
Dim szv As String

    With ppd
        .Rewind
        Do While .GetNextItem(szn, szv)
            If g_SafeLeftStr(LCase$(szn), 5) = "data-" Then _
                szData = szData & szn & ": " & szv & vbCrLf

        Loop

    End With

Dim pMail As MailItem
Dim i As Long

    Set pMail = pOutlook.CreateItem(olMailItem)

    With pMail
        .SentOnBehalfOfName = "[Snarl] " & NotificationInfo.Title
        .Subject = ppd.ValueOf("title")
        .Body = ppd.ValueOf("text") & IIf(szData <> "", vbCrLf & szData, "")

        i = Val(ppd.ValueOf("priority"))
        If i > 0 Then
            .Importance = olImportanceHigh

        ElseIf i < 0 Then
            .Importance = olImportanceLow

        End If

        .Save
        .Move pInbox

    End With

End Sub

Private Function uGetOutlook(ByRef pOutlook As Outlook.Application) As Boolean

    On Error Resume Next

    Err.Clear
    Set pOutlook = GetObject(, "outlook.application")
    uGetOutlook = (Err.Number = 0) And (Not (pOutlook Is Nothing))

End Function

