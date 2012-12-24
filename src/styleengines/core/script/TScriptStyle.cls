VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TScriptStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mHelpers As THelpers
Dim mObj As ScriptStyle
Dim mInfo As style_info
Dim mView As mfxView
Dim mScriptFile As String

Dim WithEvents theScript As ScriptControl
Attribute theScript.VB_VarHelpID = -1

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mInfo.Name

End Function

Private Function BTagItem_Value() As String
End Function

Public Function InitFrom(ByVal Path As String) As Boolean
Dim szCode As String
Dim sz As String
Dim i As Integer

    g_Debug "TScriptStyle.InitFrom()", LEMON_LEVEL_PROC_ENTER

    Set theScript = New ScriptControl
    theScript.Language = "VBScript"
    mScriptFile = g_MakePath(Path) & "script.vbs"

    g_Debug "loading " & g_MakePath(Path) & "script.vbs..."
    i = FreeFile()
    Open g_MakePath(Path) & "script.vbs" For Input As #i
    Do While Not EOF(i)
        Line Input #i, sz
        szCode = szCode & sz & vbCrLf

    Loop

    Close #i

    g_Debug "adding code..."
    If Not uSafeAddCode(szCode) Then
        g_Debug "Syntax error in script code", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    If Not uHasFunc("style_Init") Then
        g_Debug "Must have style_Init() function", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    Set mObj = New ScriptStyle
    mObj.bSetPath Path, "" '//style_GetSnarlConfigPath()
    theScript.AddObject "style", mObj, False

    mInfo.Flags = S_STYLE_IS_CONFIGURABLE

Dim ph As THelpers
Dim szErr As String

    If CallProc("style_Init") Then
        If mObj.Info.Error <> 0 Then
            szErr = mObj.Info.Reason
            If szErr = "" Then _
                szErr = "Style returned error " & CStr(mObj.Info.Error) & " but didn't supply a failure reason"

            g_Debug "style_Init() failed: " & szErr, LEMON_LEVEL_CRITICAL

        ElseIf (mObj.Info.Name = "") Or (mObj.Info.Format = "") Then
            Error = "Style didn't set info->name or info->format"
            g_Debug Error, LEMON_LEVEL_CRITICAL

        Else

            Select Case mObj.Info.Format
            Case "redirect"
                If Not HasProc("style_Notify") Then
                    g_Debug "redirect styles must support style_Notify() hook", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
                    Exit Function

                End If

                mInfo.Flags = mInfo.Flags Or S_STYLE_IS_WINDOWLESS

            Case "display"
                If Not HasProc("style_Draw") Then
                    g_Debug "display styles must support style_Draw() hook", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
                    Exit Function

                End If
                
                Set mView = New mfxView
                theScript.AddObject "view", mView, False
                theScript.AddObject "gfx", New libmgraphics21.globals, True
                Set mHelpers = New THelpers
                mHelpers.bSet mView
                theScript.AddObject "helpers", mHelpers, True

            Case Else
                Error = "Invalid style format '" & mObj.Info.Format & "'"
                g_Debug Error, LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
                Exit Function
            
            End Select

            ' /* build info */

            With mInfo
                .Copyright = mObj.Info.Copyright
                .Description = mObj.Info.Description
                .Flags = .Flags Or S_STYLE_MULTI_INSTANCE Or S_STYLE_V42_CONTENT
                .IconPath = mObj.Info.Icon
                .Major = mObj.Info.Version
                .Minor = mObj.Info.Revision
                .Name = mObj.Info.Name
                .Path = Path
                .Schemes = Replace$(mObj.Info.Schemes, ";", "|")
                .SupportEmail = ""
                .URL = mObj.Info.SupportURL

            End With

            If HasProc("style_Pulse") Then _
                mInfo.Flags = mInfo.Flags Or S_STYLE_PULSE_NEEDED

            g_Debug "style initialised ok: Type=" & mObj.Info.Format & " Name='" & mObj.Info.Name & "' (V" & mInfo.Major & "." & mInfo.Minor & ") flags=0x" & g_HexStr(mInfo.Flags), LEMON_LEVEL_INFO
            mObj.bSetOwner Me
            InitFrom = True

        End If

    Else
        g_Debug "CallProc(style_Init) failed", LEMON_LEVEL_CRITICAL
        Set mObj = Nothing

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Sub TidyUp()

    Me.CallProc "style_TidyUp"

    ' /* detach */

    If NOTNULL(mObj) Then _
        mObj.bSetOwner Nothing

    If NOTNULL(theScript) Then
        theScript.Reset
        Set theScript = Nothing

    End If

    If NOTNULL(mHelpers) Then _
        mHelpers.bSet Nothing

    Set mHelpers = Nothing
    Set mView = Nothing

End Sub

Private Function uSafeAddCode(ByVal Code As String) As Boolean

    On Error Resume Next

    Err.Clear
    theScript.AddCode Code
    uSafeAddCode = (Err.Number = 0)

End Function

Public Function HasProc(ByVal Name As String, Optional ByRef Index As Long) As Boolean

    If (theScript Is Nothing) Then
        Debug.Print "TScriptExtension.HasProc(): script not loaded"
        Exit Function

    End If

    If (Name = "") Or (theScript.Procedures.Count = 0) Then
        Debug.Print "TScriptExtension.HasProc(): bad arg"
        Exit Function

    End If

    Name = LCase$(Name)

Dim n As Long

    With theScript.Procedures
        If .Count > 0 Then
            For n = 1 To .Count
                If LCase$(.Item(n).Name) = Name Then
                    Index = n
                    HasProc = True
                    Exit Function

                End If
            Next n
        End If
    End With

'    Debug.Print "TScriptExtension.HasProc(): '" & Name & "' not found"

End Function

Private Function uHasFunc(ByVal Name As String) As Boolean
Dim n As Long

    If Not HasProc(Name, n) Then _
        Exit Function

    uHasFunc = theScript.Procedures.Item(n).HasReturnValue

End Function

Public Function CallProc(ByVal Name As String, Optional ByVal Arg1 As String, Optional ByVal Arg2 As String) As Boolean

    g_Debug "TScriptStyle.CallProc()", LEMON_LEVEL_PROC_ENTER

    On Error Resume Next

    If Not HasProc(Name) Then
        g_Debug "Proc '" & Name & "()' not found", LEMON_LEVEL_WARNING Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    Err.Clear
    If Arg1 = "" Then
        theScript.Run Name

    ElseIf Arg2 = "" Then
        theScript.Run Name, Arg1

    Else
        theScript.Run Name, Arg1, Arg2

    End If

    If Err.Number <> 0 Then _
        g_Debug "'" & Name & "()' failed: " & Err.Description, LEMON_LEVEL_WARNING

    CallProc = (Err.Number = 0)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

'Private Sub thePulseTimer_Pulse()
'
'    Me.CallProc "ext_Pulse"
'
'End Sub
'
'Public Sub SetPulseRate(ByVal Milliseconds As Long)
'
'    If (Not mWantsPulse) Or (Milliseconds < 100) Then _
'        Exit Sub
'
'    mPulseRate = Milliseconds
'
'    If mEnabled Then _
'        Set thePulseTimer = new_BTimer(mPulseRate)
'
'End Sub
'
'Public Function PulseRate() As Long
'
'    PulseRate = mPulseRate
'
'End Function

Public Function Info() As style_info

    LSet Info = mInfo

End Function

Public Function UpdateContent(ByRef Content As notification_info, ByVal RealIcon As String) As mfxBitmap

    g_Debug "TScriptStyle.UpdateContent()", LEMON_LEVEL_PROC_ENTER

    g_Debug "setting Notification object content..."
    mObj.bUpdateContent Content, RealIcon

    If (mInfo.Flags And S_STYLE_IS_WINDOWLESS) Then
        ' /* redirect */
        g_Debug "is redirect style - calling style_Notify()..."
        CallProc "style_Notify"

    Else
        ' /* display */
        g_Debug "is display style - calling style_Draw()..."
        mView.Clear
        CallProc "style_Draw"

        Set UpdateContent = mView.AsBitmap()

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function Name() As String

    Name = BTagItem_Name()

End Function

Public Function ScriptPath() As String

    ScriptPath = mScriptFile

End Function

Private Sub theScript_Error()

    With theScript.Error
        MsgBox "Error loading script " & g_Quote(mScriptFile) & vbCrLf & vbCrLf & _
               .Description & " at line " & .Line & ", column " & .Column & " (" & g_Quote(.Text) & ")" & vbCrLf & vbCrLf & "Error type: " & .Source, _
               vbExclamation Or vbOKOnly, "Scripted Style Error"

    End With

End Sub
