VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeA" (ByVal pszURL As String, ByVal pszEscaped As String, pcchEscaped As Long, ByVal dwFlags As Long) As Long

''Treat entire URL param as one URL segment
'Private Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
'Private Const URL_ESCAPE_PERCENT         As Long = &H1000
'Private Const URL_UNESCAPE_INPLACE       As Long = &H100000
'
''escape #'s in paths
'Private Const URL_INTERNAL_PATH          As Long = &H800000
'Private Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
'Private Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
'Private Const URL_DONT_SIMPLIFY          As Long = &H8000000

Dim WithEvents theRequest As CHTTPRequest

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

    g_Debug "StyleInstance.UpdateContent()", LEMON_LEVEL_PROC_ENTER

    If gConfig.UserKey = "" Then
        g_Debug "UserKey not set", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If

    If (gConfig.OnlyShowPriorityNotifications) And ((NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) = 0) Then
        g_Debug "User has disabled forwarding non-priority notifications", LEMON_LEVEL_INFO Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If

Dim szText As String
Dim pp As BPackedData
Dim i As Long

    Set pp = New BPackedData
    pp.SetTo NotificationInfo.Text

    ' /* apply RedactSensitive setting */

    If (pp.Exists("sensitivity")) And (gConfig.RedactSensitive) Then
        i = g_SafeLong(pp.ValueOf("sensitivity"))
        If i > 0 Then
            g_Debug "redacting..."
            szText = "<REDACTED>"

        Else
            szText = uTidy(pp.ValueOf("text"))

        End If

    Else
        szText = uTidy(pp.ValueOf("text"))

    End If

Dim sz As String

    sz = "https://prowl.weks.net/publicapi/add?" & _
         "apikey=" & gConfig.UserKey & _
         "&application=" & g_URLEncode(uGetAppName(NotificationInfo.Title)) & _
         "&event=" & g_URLEncode(pp.ValueOf("title")) & _
         "&description=" & g_URLEncode(szText) & _
         "&priority=" & IIf(NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY, "1", "0")

'    g_Debug "--"
'    g_Debug sz
'    g_Debug "--"

'Dim szEscaped As String
'Dim cb As Long
'
'    szEscaped = Space$(4096)
'    cb = Len(szEscaped)
'
'    If UrlEscape(sz, szEscaped, cb, URL_DONT_SIMPLIFY) <> 0 Then
'        g_Debug "StyleInstance.UpdateContent(): url escaping failed", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    Else
'        szEscaped = g_SafeLeftStr(szEscaped, cb)
'
'    End If

'    MsgBox sz

    g_Debug "creating request..."
    Set theRequest = New CHTTPRequest

    If gConfig.UseProxyServer Then
        g_Debug "setting proxy info..."
        theRequest.SetProxyInfo gConfig.ProxyServer, gConfig.ProxyServerPort, , gConfig.ProxyUsername, gConfig.ProxyPassword

    End If

    g_Debug "sending request to Prowl service..."
    theRequest.GetContent sz, gConfig.Timeout
    
    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Function uGetAppName(ByVal Sender As String) As String

    uGetAppName = gConfig.AppText
    uGetAppName = Replace$(uGetAppName, "%APP%", Sender)
    uGetAppName = Replace$(uGetAppName, "%COMPUTER%", g_GetComputerName())

End Function

Private Function uTidy(ByVal Text As String) As String

    If gConfig.ReplaceCRLFs Then
        Text = Replace$(Text, Chr$(13), "")
        Text = Replace$(Text, Chr$(10), " ")

    End If

    uTidy = Text

End Function

Private Sub theRequest_Completed()

    g_Debug "theRequest.Completed(): " & theRequest.Content()

End Sub

Private Sub theRequest_TimedOut()

    g_Debug "theRequest.TimedOut()", LEMON_LEVEL_WARNING

End Sub
