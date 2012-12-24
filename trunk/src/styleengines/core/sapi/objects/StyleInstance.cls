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

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Any) As Long

Dim WithEvents theVoice As SpVoice
Attribute theVoice.VB_VarHelpID = -1
Dim mStream As Long

Implements IStyleInstance
Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mStream)

End Function

Private Function BTagItem_Value() As String
End Function

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage
End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Dim pConfig As CConfFile
Dim pVoice As SpVoice

    g_Debug "StyleInstance.bInit()", LEMON_LEVEL_PROC_ENTER

    g_Debug "looking for '" & NotificationInfo.Scheme & "'..."
    If uGetVoice(NotificationInfo.Scheme, pVoice) Then
        g_Debug "found"
        Set theVoice = pVoice

        If uLoadConfig(NotificationInfo.Scheme, pConfig) Then

            With theVoice

                If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then _
                    .Priority = SVPAlert

                .Rate = g_SafeLong(pConfig.ValueOf("rate"))

                .Speak NotificationInfo.Title, SVSFlagsAsync
                mStream = .Speak(NotificationInfo.Text, SVSFlagsAsync)
                g_Debug "stream is #" & CStr(mStream), LEMON_LEVEL_INFO

            End With

        Else
            ' /* can't happen; uLoadConfig() always returns TRUE */
            g_Debug "there was a problem loading the config for this scheme", LEMON_LEVEL_CRITICAL

        End If

    Else
        g_Debug "not found", LEMON_LEVEL_CRITICAL
        ' /* use default voice? */

    End If
    
    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Function uGetVoice(ByVal Name As String, ByRef aVoice As SpVoice) As Boolean

    On Error Resume Next

    Err.Clear
    Set aVoice = New SpVoice
    uGetVoice = (Err.Number = 0)

    If Not uGetVoice Then _
        Exit Function

Dim pSot As ISpeechObjectTokens
Dim i As Long

    Set pSot = aVoice.GetVoices()
    For i = 0 To (pSot.Count - 1)
        If LCase$(pSot.Item(i).GetDescription) = Name Then
            Debug.Print "uGetVoice: found voice '" & Name & "'"
            aVoice.Voice = pSot.Item(i)

        End If
    Next i

End Function

Private Sub theVoice_EndStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)

    g_Debug "#EndStream: #" & CStr(StreamNumber)

End Sub

Private Function uLoadConfig(ByVal VoiceName As String, ByRef Config As CConfFile) As Boolean
Dim sz As String

    g_Debug "StyleInstance.uLoadConfig()", LEMON_LEVEL_PROC_ENTER

    ' /* load the config - doesn't matter if it's missing; we fill in any
    '    gaps with our defaults and still return TRUE */

    Set Config = New CConfFile

    If style_GetSnarlStylesPath2(False, sz) Then
        ' /* got path to user style settings - create our /speech/ folder if it doesn't
        '    already exist */

        sz = g_MakePath(sz) & "speech"
        If Not g_IsFolder(sz) Then
            g_Debug "creating folder '" & sz & "'..."
            CreateDirectory sz, ByVal 0&

        End If

        sz = g_MakePath(sz) & "voicename.conf"
        g_Debug "loading config '" & sz & "'..."
        Config.SetTo sz, False

    Else
        g_Debug "user styles folder missing", LEMON_LEVEL_WARNING

    End If

    ' /* set defaults/missing entries here */

    With Config
        .AddIfMissing "rate", "0"

    End With

    uLoadConfig = True

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function
