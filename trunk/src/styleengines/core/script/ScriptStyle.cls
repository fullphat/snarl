VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theRequest As CHTTPRequest3
Attribute theRequest.VB_VarHelpID = -1
'Dim WithEvents thePulse As BTimer

Dim mContent As Notification
'Dim mPulseFreq As Long

Dim mInfo As ScriptStyleInfo
Dim mOwner As TScriptStyle
Dim mUserPath As String
Dim mPath As String

Private Sub Class_Initialize()

    Set mInfo = New ScriptStyleInfo
'    Set mContent = New Notification
'    mPulseFreq = 250

End Sub

Private Sub Class_Terminate()
'//
End Sub

Public Property Get Info() As ScriptStyleInfo

    Set Info = mInfo

End Property

Friend Sub bSetPath(ByVal Path As String, ByVal UserPath As String)

    mPath = g_MakePath(Path)
    mUserPath = g_MakePath(UserPath)

End Sub

Friend Sub bSetOwner(ByRef Owner As TScriptStyle)

    Set mOwner = Owner

    If ISNULL(Owner) Then
        ' /* tidy up */
        Me.CancelAsyncWebRequest
'        Set thePulse = Nothing

    End If

End Sub

'Friend Sub bStartPulse()
'
'    Set thePulse = new_BTimer(mPulseFreq)
'
'End Sub

Public Property Get Path() As String

    Path = mPath

End Property

Public Property Get UserPath() As String

    UserPath = mUserPath

End Property

'Public Sub SetPulseRate(ByVal Milliseconds As Long)
'
'    If Milliseconds < 100 Then _
'        Exit Sub
'
'    mPulseFreq = Milliseconds
'
'    If NOTNULL(thePulse) Then _
'        Set thePulse = new_BTimer(Milliseconds)
'
'End Sub
'
'Public Function PulseRate() As Long
'
'    PulseRate = mPulseFreq
'
'End Function

Public Function DoAsyncWebRequest(ByVal URL As String, Optional ByVal TimeoutSeconds As Long = 5, Optional ByVal Username As String, Optional ByVal Password As String) As Boolean

    If ISNULL(theRequest) Then
        Set theRequest = New CHTTPRequest3
        DoAsyncWebRequest = theRequest.GetContent(URL, TimeoutSeconds, Username, Password)
        If Not DoAsyncWebRequest Then _
            Set theRequest = Nothing

    Else
        Debug.Print "busy"

    End If

End Function

Public Sub CancelAsyncWebRequest()

    If ISNULL(theRequest) Then _
        Exit Sub
    
    theRequest.Abort
    Set theRequest = Nothing

End Sub

Private Sub theRequest_Completed(ByVal WasSuccessful As Boolean)

    g_Debug "ScriptStyle.theRequest.Completed()", LEMON_LEVEL_PROC_ENTER
    g_Debug "success=" & CStr(WasSuccessful)
    g_Debug "responsebody=" & theRequest.ResponseBody

    If NOTNULL(mOwner) Then
        If mOwner.HasProc("style_AsyncWebRequestDone") Then _
            mOwner.CallProc "style_AsyncWebRequestDone", IIf(WasSuccessful, "1", "0"), theRequest.ResponseBody

    Else
        g_Debug "no owner set", LEMON_LEVEL_CRITICAL

    End If

    Set theRequest = Nothing
    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Friend Sub bUpdateContent(ByRef Info As notification_info, ByVal RealIcon As String)

    g_Debug "ScriptStyle.bUpdateContent()", LEMON_LEVEL_PROC_ENTER

    g_Debug "app: " & Info.Title
    g_Debug "scheme: " & Info.Scheme
    g_Debug "content: " & Info.Text

    Set mContent = New Notification

    With mContent
        .bSetTo Info.Text

        .bAdd "app-name", Info.Title
        .bAdd "scheme", Info.Scheme

        ' /* need to replace the icon passed in the content with the decoded one */

        .bRemove .IndexOf("icon")
        If RealIcon <> "" Then _
            .bAdd "icon", RealIcon

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function Notification() As Notification

    Set Notification = mContent

End Function
