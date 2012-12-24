VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWebForwardInstance"
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

Dim WithEvents theRequest As CHTTPRequest2
Attribute theRequest.VB_VarHelpID = -1
Dim mList As BTagList
Dim mUID As Long

Implements IStyleInstance
Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mUID)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = CStr(mUID)

End Function

Public Sub Init(ByRef List As BTagList)

    Set mList = List

End Sub

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage
End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Dim pf As CConfFile2

    With mList
        .Rewind
        Do While .GetNextTag(pf) = B_OK
            If LCase$(pf.GetValueWithDefault("name", "")) = NotificationInfo.Scheme Then
                uDoWebForward pf, NotificationInfo
                Exit Sub

            End If

        Loop

    End With

    g_Debug "TWebForwardInstance.UpdateContent(): '" & NotificationInfo.Scheme & "' not found", LEMON_LEVEL_CRITICAL

End Sub

Private Sub uDoWebForward(ByRef Config As CConfFile2, ByRef NotificationInfo As notification_info)

    If (Config Is Nothing) Then
        g_Debug "TWebForwardInstance.uDoWebForward(): NULL Config", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    If (Config.GetValueWithDefault("url") = "") Then
        g_Debug "TWebForwardInstance.uDoWebForward(): " & NotificationInfo.Scheme & ": missing base url and/or template", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim ppd As BPackedData

    ' /* translate the V42 packed content into a BPackedData object */

    Set ppd = New BPackedData
    ppd.SetTo NotificationInfo.Text
    ppd.Add "app", NotificationInfo.Title

Dim szURL As String

    szURL = Config.GetValueWithDefault("template")
    If szURL <> "" Then
        ' /* translate content */
        szURL = uReplaceMarked(szURL, "#", "#", ppd)

    End If

    ' /* prefix with base URL */

    szURL = Config.GetValueWithDefault("url") & szURL
    Debug.Print "WF: URL=" & szURL

    ' /* create and assign ourselves a new id */

    gID = gID + 1
    mUID = gID

    ' /* build and send the request */

    Set theRequest = New CHTTPRequest2
    With theRequest
        .SetRequestType IIf(Config.ValueOf("use_post") = "1", "POST", "GET")

        ' /* headers */

        If Config.ValueOf("suppress_std_headers") <> "1" Then
            ' /* standard */
            .AddRequestHeader "From", "gandalf@system.local"
            .AddRequestHeader "User-Agent", "Snarl-WebForward/1.0"

        End If

Dim szn As String
Dim szv As String

        If Config.ValueOf("headers") <> "" Then
            ' /* custom */
            With New BPackedData
                If .SetTo(Config.ValueOf("headers"), "\n", ": ") Then
                    .Rewind
                    Do While .GetNextItem(szn, szv)
                        theRequest.AddRequestHeader szn, szv

                    Loop
                End If
            End With
        End If

Dim szBody As String

        ' /* body (only if POST) */

        If Config.ValueOf("use_post") = "1" Then
            szBody = uReplaceMarked(Config.ValueOf("content"), "#", "#", ppd)
            Debug.Print "WF: Content=" & szBody
            .SetBody szBody

        End If

        .GetContent szURL

    End With

    g_AddThis Me

End Sub

Private Function uStripCRLF(ByVal Text As String) As String

    Text = Replace$(Text, vbCrLf, "¶")
    Text = Replace$(Text, vbCr, "¶")
    Text = Replace$(Text, vbLf, "¶")

    uStripCRLF = Text

End Function

Public Sub Go(ByRef Scheme As CConfFile2, ByRef NotificationInfo As notification_info)

    uDoWebForward Scheme, NotificationInfo

End Sub

Private Sub theRequest_Completed(ByVal WasSuccessful As Boolean)

    g_Debug "TWebForwardInstance.theRequest_Completed(): request " & CStr(mUID) & " completed (success=" & CStr(WasSuccessful) & ")"

    If WasSuccessful Then _
        g_Debug "TWebForwardInstance.theRequest_Completed(): response was '" & uStripCRLF(theRequest.ResponseBody) & "'"

    g_RemoveThis Me

End Sub

Private Function uReplaceMarked(ByVal Text As String, ByVal StartMarker As String, ByVal EndMarker As String, ByRef Data As BPackedData, Optional ByVal DefaultData As String) As String

    If (Text = "") Or (StartMarker = "") Or (EndMarker = "") Or (Data Is Nothing) Then _
        Exit Function

Dim i As Long
Dim j As Long
Dim sz As String
Dim szt As String
Dim sz64 As String

    j = 1
    i = InStr(Text, StartMarker)
    If i = 0 Then
        ' /* no start marker: just return text */
        uReplaceMarked = Text
        Exit Function

    End If

    Do While i
        Debug.Print "uReplaceMarked: start marker at " & CStr(i)
        If i > j Then _
            sz = sz & g_SafeMidStr(Text, j, i - j)      ' // add text up to marker

        j = i + Len(StartMarker)
        i = InStr(j, Text, EndMarker)
        If i = 0 Then
            Debug.Print "uReplaceMarked: error - no end marker"
            Exit Function

        Else
            Debug.Print "uReplaceMarked: end marker at " & CStr(i)

        End If

        ' /* get marker */
        szt = g_SafeMidStr(Text, j, i - j)

        Select Case szt
        Case "icon-base64"
            ' /* special case: insert icon data encoded as Base64 */
            If EncodeFile(Data.ValueOf("icon"), sz64) Then _
                sz = sz & sz64

        Case Else
            If Data.Exists(szt) Then
                sz = sz & Data.ValueOf(szt)

            Else
                sz = sz & DefaultData

            End If

        End Select

        j = i + Len(EndMarker)
        i = InStr(j, Text, StartMarker)

    Loop

    uReplaceMarked = sz

End Function





