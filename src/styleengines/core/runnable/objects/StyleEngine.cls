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

    '/*********************************************************************************************
    '/
    '/  File:           StyleEngine.cls
    '/
    '/  Description:    Provides the Runnable style engine ingress point
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Const LIB_DATE = "19-Jan-2012"

Dim mStyles As BTagList
Dim mLastErr As String
Dim mConfig As CConfFile

Implements IStyleEngine

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = mStyles.CountItems

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance
Dim pri As TRunnableInstance
Dim i As Long

    ' /* look it up (translates friendly name to signature) */

    i = mStyles.IndexOf(StyleName)
    If i Then
        Set pri = New TRunnableInstance
        pri.Init mStyles.TagAt(i), g_SafeLong(mConfig.ValueOf("timeout"))
        Set IStyleEngine_CreateInstance = pri

    Else
        g_Debug "IStyleEngine.CreateInstance(): '" & StyleName & "' not found", LEMON_LEVEL_CRITICAL

    End If

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = "Manages styles provided by standalone runnable processes"

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long
Dim ps As TRunnableStyle
Dim i As Long

    If StyleName = ".global" Then
        MsgBox "This style engine will support a user interface in a future release.  " & _
               "For now, configurable features can be modified directly via the settings " & _
               "file (located at '" & g_MakePath(style_GetSnarlStylesPath(False)) & "runnable\runnable.conf').", _
               vbInformation Or vbOKOnly, "Runnable Style Engine"

    Else
        i = mStyles.IndexOf(StyleName, False)
        If i > 0 Then
            Set ps = mStyles.TagAt(i)
            ShellExecute 0, "open", ps.Executable, "-configure", vbNullString, 5

        End If

    End If

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    l3OpenLog "%APPDATA%\full phat\snarl\" & "runnable.log", True
    g_Debug "IStyleEngine.Initialise()", LEMON_LEVEL_PROC_ENTER

    ' /* load up our global config */

    Set mConfig = New CConfFile

Dim sz As String

    If style_GetSnarlStylesPath2(False, sz) Then
        sz = g_MakePath(sz) & "runnable\settings.conf"
        g_Debug "loading config from '" & sz & "'..."
        mConfig.SetTo sz, False

    Else
        g_Debug "user style path missing", LEMON_LEVEL_WARNING

    End If

    With mConfig
        .AddIfMissing "timeout", "750"
        
        
    End With

    Select Case g_SafeLong(mConfig.ValueOf("timeout"))
    Case 200 To 1500
    
    Case Else
        g_Debug "timeout of " & mConfig.ValueOf("timeout") & " is invalid, resetting...", LEMON_LEVEL_WARNING
        mConfig.Update "timeout", "750"

    End Select

    g_Debug "Initialise(): using timeout of " & mConfig.ValueOf("timeout") & "ms"

    IStyleEngine_Initialize = uGetStyles()

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function IStyleEngine_LastError() As String

    IStyleEngine_LastError = mLastErr

End Function

Private Function IStyleEngine_Name() As String

    IStyleEngine_Name = "runnable.styleengine"

End Function

Private Function IStyleEngine_Path() As String

    IStyleEngine_Path = App.Path

End Function

Private Function IStyleEngine_Revision() As Long

    IStyleEngine_Revision = App.Revision

End Function

Private Sub IStyleEngine_StyleAt(ByVal Index As Long, Style As libSnarlStyles.style_info)
Dim pr As TRunnableStyle

    If Index = 0 Then
        ' /* provide info about ourselves */
        With Style
            .Flags = S_STYLE_IS_CONFIGURABLE

        End With

    Else
        Set pr = mStyles.TagAt(Index)
        pr.Populate Style

    End If

End Sub

Private Sub IStyleEngine_TidyUp()

    Set mStyles = Nothing
    l3CloseLog

End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function

Private Function uGetStyles() As M_RESULT

    g_Debug "StyleEngine.uGetStyles()", LEMON_LEVEL_PROC_ENTER

Dim sz As String

    Set mStyles = new_BTagList()
    mLastErr = ""

    ' /* if in IDE, from local path */

    g_Debug "local..."
    If g_IsIDE() Then _
        uGetFromFolder g_MakePath(App.Path) & "styles"

    ' /* load styles from All Users first */

    g_Debug "all users..."
    If style_GetSnarlStylesPath2(True, sz) Then _
        uGetFromFolder g_MakePath(sz) & "runnable"

    ' /* load styles from user folder last */

    g_Debug "this user..."
    If style_GetSnarlStylesPath2(False, sz) Then _
        uGetFromFolder g_MakePath(sz) & "runnable"

    If mStyles.CountItems = 0 Then
        g_Debug "no styles found", LEMON_LEVEL_CRITICAL
        mLastErr = "No suitable styles found"
        uGetStyles = M_NOT_FOUND

    Else
        g_Debug CStr(mStyles.CountItems) & " style(s) loaded", LEMON_LEVEL_INFO
        uGetStyles = M_OK

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub uGetFromFolder(ByRef Path As String)

    g_Debug "StyleEngine.uGetFromFolder()", LEMON_LEVEL_PROC_ENTER

    g_Debug "looking in '" & Path & "'..."

    If Not g_IsFolder(Path) Then
        g_Debug "not a folder", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Sub

    End If

Dim pr As TRunnableStyle
Dim sz As String

    With New CFolderContent2
        If .SetTo(Path) Then
            .Rewind
            Do While .GetNextFile(sz)
                If g_IsFolder(sz) Then
                    Set pr = New TRunnableStyle
                    If pr.SetTo(sz) Then
                        If mStyles.IndexOf(pr.Name) = 0 Then
                            mStyles.Add pr
                            g_Debug "added " & IIf(pr.IsRedirect, "redirect", "display") & " " & g_Quote(pr.Name), LEMON_LEVEL_INFO

                        Else
                            g_Debug "not added: a style called " & g_Quote(pr.Name) & " is already loaded", LEMON_LEVEL_WARNING
                            
                        End If
                    End If
                End If
            Loop
        Else
            g_Debug "invalid path", LEMON_LEVEL_CRITICAL

        End If

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

'Private Function uStyleExists(ByVal ShortName As String) As Boolean
'Dim ps As TRunnableStyle
'
'    With mStyles
'        .Rewind
'        Do While .GetNextTag(ps) = B_OK
'            If ps.ShortName = ShortName Then
'                g_Debug "uStyleExists(): can't add '" & ShortName & "': a style with this name already exists", LEMON_LEVEL_CRITICAL
'                uStyleExists = True
'                Exit Do
'
'            End If
'
'        Loop
'    End With
'
'End Function


