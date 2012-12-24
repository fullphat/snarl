VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const ENG_DATE = "19-Jan-2012"

Dim mGotSAPI As Boolean
Dim mLastErr As String
'Dim mPanel As BPrefsPanel

Dim mIndex As Long
Dim mCount As Long
Dim mSchemes As String

Dim mVoices As BTagList

Implements IStyleEngine

'Private Function ISSEngine_ConfigRequested(ByVal StyleAndScheme As String) As melon.M_RESULT
'
''    If (mPanel Is Nothing) Then
''        Set mPanel = New BPrefsPanel
''        With mPanel
''            .SetHandler Me
''            .SetTitle App.FileDescription & " Preferences"
''            .SetWidth 300
''            .AddPage new_BPrefsPage("General", load_image_obj(g_MakePath(App.Path) & "icon.png"), New TGeneralPage)
''            .Go
'''    g_SetWindowIconToAppResourceIcon mPanel.hWnd
''
''        End With
''    End If
''
''    ISSEngine_ConfigRequested = mPanel.hWnd
'
'End Function
'
'Private Function ISSEngine_CreateInstance(ByVal StyleAndScheme As String) As libSnarlStyleEngine.ISSInstance
'Dim pInstance As StyleInstance
'Dim sz() As String
'
'    sz() = Split(StyleAndScheme, "/")
'    If UBound(sz()) < 1 Then _
'        Exit Function
'
'    Set pInstance = New StyleInstance
'    pInstance.Scheme = LCase$(sz(1))
'    Set ISSEngine_CreateInstance = pInstance
'
'End Function
'
'Private Function ISSEngine_GetInfo() As melon.MMessage
'Dim pm As CTempMsg
'
'    Set pm = New CTempMsg
'    If mLastErr <> "" Then _
'        pm.Add "last_error", mLastErr
'
'    pm.Add "version", App.Major
'    pm.Add "revision", App.Revision
'    pm.Add "path", App.Path
'    pm.Add "author", App.CompanyName
'    pm.Add "copyright", App.LegalCopyright
'
'    Set ISSEngine_GetInfo = pm
'
'End Function
'
'Private Function ISSEngine_GetStyle(Style As libSnarlStyleEngine.SNARL_STYLE) As Boolean
'
'    If Style.Name <> App.FileDescription Then _
'        Exit Function
'
'Dim pm As CTempMsg
'Dim i As Long
'
'    With Style
'        .Name = App.FileDescription
'        .Author = App.CompanyName
'        .Copyright = App.LegalCopyright
'        .Date = LIB_DATE
'        .Email = "snarl@fullphat.net"
'        .Flags = SNARL_STYLE_WINDOWLESS Or SNARL_STYLE_HAS_SETTINGS
'        .Path = g_MakePath(App.Path)
'        .Revision = 0
'        .URL = "www.fullphat.net"
'        .Version = 1
'
'        ' /* convert the schemes list into a message */
'
'        Set pm = New CTempMsg
'        pm.What = gSchemes
'
'        For i = 1 To gSchemes
'            pm.Add CStr(i), gScheme(i).Name
'
'        Next i
'
'        Set .Schemes = pm
'
'    End With
'
'    ISSEngine_GetStyle = True
'
'End Function
'
'Private Function ISSEngine_Initialize() As melon.M_RESULT
'
'    On Error Resume Next
'
''Dim pVoiceList As ISpeechObjectTokens
'Dim pv As SpVoice
''Dim i As Long
'
'    Err.Clear
'    Set pv = New SpVoice
'    If Err.Number <> 0 Then
'        mLastErr = "Couldn't create Microsoft SAPI object"
'        ISSEngine_Initialize = M_FAILED
'
'    Else
'
'        ' /* load up schemes */
'
'        uGetSchemes
'        mCount = 1
'
'    End If
'
'End Function
'
'Private Sub ISSEngine_Uninitialize()
'
''    ReDim mStyle(0)
''    mCount = 0
'
'End Sub

'Private Sub uAdd(ByRef Scheme As T_SCHEME)
'
'    gSchemes = gSchemes + 1
'    ReDim Preserve gScheme(gSchemes)
'    LSet gScheme(gSchemes) = Scheme
'
'    ' /* convert schemes to a message */
'
'
'
''Dim pm As CTempMsg
''Dim sz() As String
''Dim c As Long
''Dim i As Long
''
''    Set pm = New CTempMsg
''    sz() = Split(Schemes, "|")
''
''    c = UBound(sz()) + 1
''    pm.What = c
''
''    If c > 0 Then
''        For i = 1 To c
''            pm.Add CStr(i), sz(i - 1)
''
''        Next i
''
''    End If
''
''    ' /* add the style */
''
''    mCount = mCount + 1
''    ReDim Preserve mStyle(mCount)
''    With mStyle(mCount)
''        .Name = Name
''        Set .Schemes = pm
''        .Author = "full phat products"
''        .Copyright = App.LegalCopyright
''        .Date = LIB_DATE
''        .Email = "snarl@fullphat.net"
''        .Path = g_MakePath(App.Path)
''        .Revision = App.Revision
''        .URL = "www.fullphat.net"
''        .Version = App.Major
''
''    End With
'
'End Sub

'Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
'
'End Sub
'
'Private Sub KPrefsPanel_Quit()
'Dim sz As String
'Dim i As Long
'
'    sz = uGetConfigPath()
'    If sz <> "" Then
'        With New ConfigFile
'            .File = sz
'
'            For i = 1 To gSchemes
'                With .AddSectionObj("scheme")
'                    .Add "name", gScheme(i).Name
'                    .Add "voice", gScheme(i).VoiceToUse
'                    .Add "rate", CStr(gScheme(i).VoiceRate)
'
'                End With
'            Next i
'
'            .Save
'
'        End With
'    End If
'
'    Set mPanel = Nothing
'
'End Sub
'
'Private Sub KPrefsPanel_Ready()
'Dim pc As BControl
'
'    If mPanel.Find("lb>schemes", pc) Then _
'        pc.SetValue "1"
'
'End Sub
'
'Private Sub KPrefsPanel_Selected(ByVal Command As String)
'
'End Sub

'Private Sub uGetSchemes()
'Dim sz As String
'
'    gSchemes = 0
'    ReDim gScheme(0)
'
'    sz = uGetConfigPath()
'    If sz = "" Then _
'        Exit Sub
'
'Dim pSection As ConfigSection
'Dim pScheme As T_SCHEME
'
'    With New ConfigFile
'        .File = sz
'        If .Load() Then
'            ' /* get schemes */
'
'            .Rewind
'            Do While .GetNextSection(pSection)
'                With pSection
'                    If .Name = "scheme" Then
'                        pScheme.Name = .GetValueWithDefault("name", "")
'                        pScheme.VoiceToUse = .GetValueWithDefault("voice", "")
'
'                        sz = .GetValueWithDefault("rate", "0")
'                        If (Val(sz) >= -10) And (Val(sz) <= 10) Then _
'                            pScheme.VoiceRate = Val(sz)
'
'                    End If
'                End With
'
'                uAdd pScheme
'
'            Loop
'
'        End If
'
'    End With
'
'End Sub
'
'Private Function uGetConfigPath() As String
'Dim sz As String
'
'    If Not g_GetSystemFolder(CSIDL_APPDATA, sz) Then _
'        Exit Function
'
'    uGetConfigPath = g_MakePath(sz) & "full phat\snarl\etc\.speech"
'
'End Function

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = IIf(mGotSAPI, 1, 0)

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance

    If Not mGotSAPI Then _
        Exit Function

    Set IStyleEngine_CreateInstance = New StyleInstance

    ' /* add it to our list to keep the object open */

    mVoices.Add IStyleEngine_CreateInstance

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = ENG_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = "Speaks the notification"

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    l3OpenLog "%APPDATA%\full phat\snarl\sapi.log", True

    On Error Resume Next

Dim pVoice As SpVoice
Dim pSot As ISpeechObjectTokens
Dim i As Long

    g_Debug "IStyleEngine.Initialise()", LEMON_LEVEL_PROC_ENTER
    g_Debug "checking SAPI..."

    Err.Clear
    Set pVoice = New SpVoice
    mGotSAPI = (Err.Number = 0)

    If mGotSAPI Then
        ' /* each voice is a scheme */
        Set mVoices = new_BTagList()
        g_Debug "success; loading voices..."

        Set pSot = pVoice.GetVoices()
        For i = 0 To pSot.Count - 1
            g_Debug g_Quote(pSot.Item(i).GetDescription())
            mSchemes = mSchemes & pSot.Item(i).GetDescription() & "|"

        Next i

        mSchemes = g_SafeLeftStr(mSchemes, Len(mSchemes) - 1)
        mLastErr = ""
        g_Debug "done"

    Else
        g_Debug "failed to create an SpVoice object (" & Err.Description & ")", LEMON_LEVEL_CRITICAL
        mLastErr = "Couldn't create Microsoft SAPI object"
        IStyleEngine_Initialize = M_FAILED

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function IStyleEngine_LastError() As String

    IStyleEngine_LastError = mLastErr

End Function

Private Function IStyleEngine_Name() As String

    IStyleEngine_Name = App.Title

End Function

Private Function IStyleEngine_Path() As String

    IStyleEngine_Path = App.Path

End Function

Private Function IStyleEngine_Revision() As Long

    IStyleEngine_Revision = App.Revision

End Function

Private Sub IStyleEngine_StyleAt(ByVal Index As Long, Style As libSnarlStyles.style_info)

    g_Debug "IStyleEngine.StyleAt()", LEMON_LEVEL_PROC_ENTER

    If Not mGotSAPI Then
        g_Debug "can't provide style - SAPI not loaded", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If

    If Index <> 1 Then
        g_Debug "can't provide style for index #" & CStr(Index), LEMON_LEVEL_WARNING Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If

    g_Debug "asked for style at index #" & CStr(Index)

    With Style
        .Copyright = App.LegalCopyright
        .Date = ENG_DATE
        .Description = "Speaks the notification (Beta)"
        .Flags = S_STYLE_IS_WINDOWLESS
        .Path = g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "")
        .Major = 1
        .Minor = 0
        .Name = "Speech"
        .IconPath = .Path & "icon.png"
        .Schemes = mSchemes
        .SupportEmail = "info@fullphat.net"
        .URL = "www.fullphat.net"

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub IStyleEngine_TidyUp()

    g_Debug "IStyleEngine.TidyUp(): nothing to do"

End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function
