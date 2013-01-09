VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWebListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theRequest As BWebRequest
Attribute theRequest.VB_VarHelpID = -1

Dim mGuid As String
Dim mName As String     ' // user-friendly
Dim mURL As String

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mGuid

End Function

Private Function BTagItem_Value() As String

End Function

Private Sub Class_Initialize()
'///
End Sub

Private Sub Class_Terminate()

    theRequest.Abort
    Set theRequest = Nothing

End Sub

Public Function SetFrom(ByRef Config As ConfigSection) As Boolean

'    If g_SafeLong(Config.GetValueWithDefault("type", "")) <> SN_ST_COMET Then _
'        Exit Function
'
'    mGuid = Config.Name
'    mName = Config.GetValueWithDefault("name", "")
'    mURL = Config.GetValueWithDefault("url", "")
'
'    ' /* TO DO: check URL is valid rather than just not empty */
'
'    SetFrom = (mName <> "") And (mURL <> "") And (mGuid <> "")
'
'    If SetFrom Then _
'        uStartPolling

End Function

Public Function CreateNew(ByVal Name As String, ByVal URL As String) As Boolean

    If (Name = "") Or (URL = "") Then _
        Exit Function

    mGuid = g_CreateGUID(True)
    mName = Name
    mURL = URL
    uStartPolling
    CreateNew = True

End Function

Public Function URL() As String

    URL = mURL

End Function

Public Function Name() As String

    Name = mName

End Function

Public Function Guid() As String

    Guid = mGuid

End Function

Private Sub uStartPolling()

    g_Debug "TWebListener.uStartPolling(): listening to '" & mURL & "'..."
    Set theRequest = New BWebRequest
    theRequest.GetContent mURL, 3600            ' // timeout of an hour...

End Sub

Private Sub theRequest_Completed(ByVal WasSuccessful As Boolean)

    g_Debug "TWebListener.Completed(): '" & mURL & "': " & WasSuccessful

    If WasSuccessful Then
        g_Debug "--"
        g_Debug theRequest.Content
        g_Debug "--"
    
    End If

    uStartPolling

End Sub
