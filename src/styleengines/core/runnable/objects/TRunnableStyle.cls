VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRunnableStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

        '/*********************************************************************************************
        '/
        '/  File:           TRunnableStyle.cls
        '/
        '/  Description:    Represents a particular runnable style
        '/
        '/  © 2011 full phat products
        '/
        '/  This file may be used under the terms of the Simplified BSD Licence
        '/
        '*********************************************************************************************/

Dim mConfig As CConfSection
Dim mSig As String                      ' // folder prefix: should be defined in the config file
Dim mPath As String                     ' // path as discovered by the style engine
Dim mExecutable As String               ' // "style.exe" or whatever was provided by the config (path removed)

Dim mSchemes As CConfSection

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mConfig.GetValueWithDefault("name", mSig)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mPath

End Function

Public Function SetTo(ByVal Path As String) As Boolean

    ' /* must have:
    '       1. an executable (style.exe) [alternatively runnable.conf can point to the exe to run]
    '       2. a config file (runnable.conf)
    ' */

    g_Debug "TRunnableStyle.SetTo()", LEMON_LEVEL_PROC_ENTER

Dim pc As CConfFile3
Dim ps As CConfSection
Dim szSig As String

    szSig = g_FilenameFromPath(Path)
    g_Debug "signature: " & g_Quote(szSig)

    With New CFolderContent2
        If .SetTo(Path) Then
            If .Exists("runnable.conf") Then
                Set pc = New CConfFile3
                pc.SetFile g_MakePath(Path) & "runnable.conf"
                If pc.Load Then
                    If pc.Find("general", ps) Then
                        Set mConfig = ps
                        
                        ' /* required stuff */

                        mSig = szSig
                        mPath = g_MakePath(Path)
                        mExecutable = mPath & g_FilenameFromPath(ps.GetValueWithDefault("style_exe", "style.exe"))
                        SetTo = g_Exists(mExecutable)

                        If SetTo Then
                            ' /* optional stuff */
                            If pc.Find("schemes", ps) Then
                                Set mSchemes = ps

                            Else
                                g_Debug "style " & g_Quote(mSig) & " says it has no schemes", LEMON_LEVEL_WARNING
                                Set mSchemes = New CConfSection

                            End If
                        Else
                            g_Debug "missing style exe " & g_Quote(mExecutable), LEMON_LEVEL_CRITICAL
                        
                        End If
                    Else
                        g_Debug "missing [general] section from " & g_Quote(Path), LEMON_LEVEL_CRITICAL

                    End If
                Else
                    g_Debug "failed to load runnable.conf from " & g_Quote(Path), LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "no runnable.conf in " & g_Quote(Path), LEMON_LEVEL_CRITICAL

            End If
        Else
            g_Debug "bad path " & g_Quote(Path), LEMON_LEVEL_CRITICAL

        End If

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Sub Populate(ByRef StyleInfo As style_info)
Dim szSchemes As String
Dim pe As CConfEntry

    ' /* get schemes */

    With mSchemes
        .Rewind
        Do While .NextEntry(pe)
            szSchemes = szSchemes & pe.Name & "|"

        Loop

        szSchemes = g_SafeLeftStr(szSchemes, Len(szSchemes) - 1)

    End With

    ' /* get general info */

    With StyleInfo
        .Copyright = mConfig.GetValueWithDefault("copyright")
        .Date = mConfig.GetValueWithDefault("release_date")

        .Description = mConfig.GetValueWithDefault("description")
        If .Description = "" Then _
            .Description = "Runnable style"

'        Else
'            .Description = .Description & " (runnable style)"
'
'        End If

        .Flags = S_STYLE_MULTI_INSTANCE Or S_STYLE_V42_CONTENT

        If mConfig.GetValueWithDefault("dropshadow", "1") = "0" Then _
            .Flags = .Flags Or S_STYLE_NO_DROPSHADOW

        If mConfig.GetValueWithDefault("configurable", "0") = "1" Then _
            .Flags = .Flags Or S_STYLE_IS_CONFIGURABLE

        If Me.IsRedirect() Then
            .Flags = .Flags Or S_STYLE_IS_WINDOWLESS Or S_STYLE_DONT_QUEUE
            If Me.IsScreenRedirect Then _
                .Flags = .Flags Or S_STYLE_REDIRECT_TO_SCREEN

        End If

        .IconPath = g_MakePath(mPath) & g_FilenameFromPath(mConfig.GetValueWithDefault("icon", "icon.png"))
        If Not g_Exists(.IconPath) Then _
            .IconPath = ""

        .Major = g_SafeLong(mConfig.GetValueWithDefault("major", "0"))
        .Minor = g_SafeLong(mConfig.GetValueWithDefault("minor", "0"))
        .Name = mConfig.GetValueWithDefault("name", mSig)
        .Path = mPath
        .Schemes = szSchemes
        .SupportEmail = mConfig.GetValueWithDefault("email")
        .URL = mConfig.GetValueWithDefault("url")

    End With

End Sub

Public Function Executable() As String

    Executable = mExecutable

End Function

'Public Function ShortName() As String
'
'    ShortName = LCase$(g_FilenameFromPath(mPath))
'
'End Function

Public Function Path() As String

    Path = mPath

End Function

Public Function IsRedirect() As Boolean

    If NOTNULL(mConfig) Then _
        IsRedirect = (mConfig.GetValueWithDefault("redirect", "0") = "1") Or (IsScreenRedirect())

End Function

Public Function IsScreenRedirect() As Boolean

    If NOTNULL(mConfig) Then _
        IsScreenRedirect = (mConfig.GetValueWithDefault("redirect_to_screen", "0") = "1")

End Function

Public Function Name() As String

    Name = BTagItem_Name()

End Function
