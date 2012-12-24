VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyleEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum SN_STYLEENGINE_STATE
    SN_SE_FAILED = -1            ' // mError should contain reason
    SN_SE_NOT_LOADED = 0
    SN_SE_LOADED = 1

End Enum

Dim mClassId As String
Dim mObject As IStyleEngine
Dim mState As SN_STYLEENGINE_STATE
Dim mFlags As Long

Dim mError As String

Dim mSupportsEngineConfig As Boolean

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mClassId

End Function

Private Function BTagItem_Value() As String
End Function

Public Sub Init(ByVal ClassId As String, ByVal Flags As Long)

    mClassId = ClassId
    mState = SN_SE_NOT_LOADED
    mFlags = Flags

End Sub

Public Function Load() As Boolean
Dim pse As IStyleEngine

    On Error Resume Next

    If mState = SN_SE_LOADED Then
        g_Debug "TStyleEngine.Load(): '" & mClassId & "' is already loaded", LEMON_LEVEL_WARNING
        Exit Function

    End If

    err.Clear

    If mClassId = "<internal>.styleengine" Then
        ' /* special case: create our own internal style engine */
        Set pse = New TInternalStyle

    Else
        Set pse = CreateObject(mClassId)

    End If

    If err.Number <> 0 Then
        ' /* fatal error */
        g_Debug "TStyleEngine.Load(): " & err.Description & "' creating '" & mClassId & "'", LEMON_LEVEL_CRITICAL
        mError = "Not a valid style engine"
        mState = SN_SE_FAILED
        Exit Function

    End If

    ' /* created as IStyleEngine */

Dim hr As M_RESULT

    hr = pse.Initialize()

    If hr <> M_OK Then
        g_Debug "TStyleRoster.uAddStyleEngine(): style engine didn't initialize (" & g_HexStr(hr) & ")", LEMON_LEVEL_CRITICAL
        mError = pse.LastError()
        mState = SN_SE_FAILED
        Exit Function

    End If

    ' /* R2.5 Beta 2: ask the style engine for information about itself
    '    by calling IStyleEngine::StyleAt(0) */

Dim psi As style_info

    pse.StyleAt 0, psi
    If (psi.Flags And S_STYLE_IS_CONFIGURABLE) Then _
        mSupportsEngineConfig = True


    If pse.CountStyles < 1 Then
        g_Debug "TStyleRoster.uAddStyleEngine(): style engine has no styles", LEMON_LEVEL_WARNING
        mError = "No styles found"
        mState = SN_SE_FAILED
        Exit Function

    End If

    ' /* success */

    Set mObject = pse
    mState = SN_SE_LOADED
    Load = True

End Function

Public Sub Unload()

    If mState = SN_SE_LOADED Then
        If Not (mObject Is Nothing) Then
            mObject.TidyUp
            Set mObject = Nothing

        End If

    End If

    mState = SN_SE_NOT_LOADED

End Sub

Public Sub TidyUp()


End Sub

Public Function State() As SN_STYLEENGINE_STATE

    State = mState

End Function

Public Function Error() As String

    Error = mError

End Function

Public Function Engine() As IStyleEngine

    Set Engine = mObject

End Function

Public Function InfoString() As String

    If (mObject Is Nothing) Then
        InfoString = mClassId

    Else
        With mObject
            InfoString = g_RemoveExtension(.Name) & " V" & CStr(.Version) & "." & CStr(.Revision) & IIf(.Date <> "", " (" & .Date & ")", "")

        End With

    End If

End Function

Public Function StyleCountString() As String

    If (mObject Is Nothing) Then
        StyleCountString = "No styles"

    Else
        StyleCountString = CStr(mObject.CountStyles) & " style" & IIf(mObject.CountStyles = 1, "", "s")

    End If

End Function

Public Sub DoDetailsNotification()

    If (mObject Is Nothing) Then
        g_Debug "TStyleEngine.DoDetailsNotification(): style engine not set for '" & mClassId & "'", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim ps As style_info
Dim szs As String
Dim sz As String
Dim i As Long

    With mObject

        If (.Version <> 0) And (.Revision <> 0) Then _
            sz = "Version: " & CStr(.Version) & "." & CStr(.Revision)

        If .Date <> "" Then _
            sz = sz & " (" & .Date & ")" & vbCrLf

        If .Description <> "" Then _
            sz = sz & .Description

        ' /* list of styles it provides */

'        If .CountStyles > 0 Then
'            sz = sz & vbCrLf & vbCrLf & "Styles: "
'
'            For i = 1 To .CountStyles
'                .StyleAt i, ps
'                szs = szs & ps.Name & ", "
'
'            Next i
'
'            sz = sz & g_SafeLeftStr(szs, Len(szs) - 2)
'
'        End If

        g_NotificationRoster.Hide gStyleEngineDetailsToken, "", "", ""
        gStyleEngineDetailsToken = g_PrivateNotify(SNARL_CLASS_GENERAL, g_RemoveExtension(.Name), sz, , g_MakePath(App.Path) & "etc\icons\style_engine.png")

    End With

End Sub

Public Function Name() As String

    Name = mClassId

End Function

Public Function IsSystem() As Boolean

    IsSystem = (mFlags And &H80000000)

End Function

Public Function Description() As String

    If Not (mObject Is Nothing) Then _
        Description = mObject.Description

End Function

Public Function IsConfigurable() As Boolean

    IsConfigurable = mSupportsEngineConfig

End Function

Public Sub Configure()

    If (mState <> SN_SE_LOADED) Or (Not mSupportsEngineConfig) Then _
        Exit Sub

Dim hWnd As Long

    g_Debug "WARNING: style engine configuration is not currently supported", LEMON_LEVEL_CRITICAL

    hWnd = mObject.GetConfigWindow(".global")
    If IsWindow(hWnd) <> 0 Then
        ' /* display it */
            
    End If

End Sub

Public Function StyleList() As BTagList

    Set StyleList = new_BTagList()

    If ISNULL(mObject) Then _
        Exit Function

Dim psi As style_info
Dim i As Long

    With mObject
        If .CountStyles Then
            For i = 1 To .CountStyles
                .StyleAt i, psi
                StyleList.Add new_BTagItem(psi.Name, psi.IconPath)
    
            Next i
        End If
    End With

End Function
