VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyleRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TStyleRoster.cls
    '/
    '/  Description:    Manages style engines and the styles they provide
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mList As BTagList

    ' /* full list of all supported styles and schemes */
Dim mItem() As TStyle
Dim mItems As Long
Dim mIndex As Long

Dim mRefs As Long
Dim mVersions As CConfFile          ' // R2.2: list of engines and the installed version
Dim mStartup As CConfFile           ' // R2.4.2: list of engines and their boot-up state

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT
Dim pse As TStyleEngine
Dim i As Long

    mRefs = mRefs - 1
    If mRefs = 0 Then
        ' /* run tidyup code */



        ' /* unload style engines */


        With mList
            .Rewind
            Do While .GetNextTag(pse) = B_OK
                pse.Unload
                pse.TidyUp

            Loop
        
        End With

        Set mList = Nothing

        mItems = 0
        ReDim mItem(0)

    End If

    If g_IsWinXPOrBetter() Then _
        g_FlushLibs

End Function

Private Sub MLibrary_Initialize()
    ' /* not currently used */
End Sub

Private Function MLibrary_Magic() As Long

    MLibrary_Magic = &H23232323

End Function

Private Function MLibrary_Open() As melon.M_RESULT
Dim pFolder As storage_kit.Node
Dim sz As String
Dim i As Long

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        Set mList = new_BTagList()

        ' /* read our boot config, if it exists */

        Set mStartup = New CConfFile
        If g_GetUserFolderPath(sz) Then
            If mStartup.SetTo(g_MakePath(sz) & "etc\styleengines.boot", True) Then
                g_Debug "TStyleRoster.Open(): startup list loaded"

            Else
                g_Debug "TStyleRoster.Open(): startup list not found"

            End If

        End If

        uGetVersionList
        uGetStyleEngine "<internal>.styleengine", &H80000000

        ' /* R2.5.1: style link files are now in the installation folder... */

        g_GetAppFolderNode pFolder, "styles"
        If pFolder.IsFolder() Then
            pFolder.ReadContents
            For i = 1 To pFolder.CountNodes
                If LCase$(pFolder.NodeAt(i).Extension = "styleengine") Then _
                    uGetStyleEngine pFolder.NodeAt(i).Filename

            Next i

        Else
            g_Debug "TStyleRoster.Open(): '" & g_MakePath(pFolder.File) & "' is not a folder", LEMON_LEVEL_WARNING

        End If

        ' /* ...but we also look in the user folder as well for compatability reasons */

        g_GetUserFolder pFolder, False, "styles"
        If pFolder.IsFolder() Then
            pFolder.ReadContents
            For i = 1 To pFolder.CountNodes
                If LCase$(pFolder.NodeAt(i).Extension = "styleengine") Then _
                    uGetStyleEngine pFolder.NodeAt(i).Filename

            Next i

        Else
            g_Debug "TStyleRoster.Open(): '" & g_MakePath(pFolder.File) & "' is not a folder", LEMON_LEVEL_WARNING

        End If
    End If

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()
    ' /* not currently used */
End Sub

Private Function MRoster_CountItems() As Long

    g_Debug "TStyleRoster: Call to MRoster->CountItems", LEMON_LEVEL_WARNING

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long

    g_Debug "TStyleRoster: Call to MRoster->FindItem", LEMON_LEVEL_WARNING

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.mObject

    g_Debug "TStyleRoster: Call to MRoster->ItemAt", LEMON_LEVEL_WARNING

End Function

Private Function MRoster_NextItem(Item As melon.mObject) As Boolean

    g_Debug "TStyleRoster: Call to MRoster->NextItem", LEMON_LEVEL_WARNING

End Function

Private Sub MRoster_Rewind()

    g_Debug "TStyleRoster: Call to MRoster->Rewind", LEMON_LEVEL_WARNING

End Sub

Private Property Get MVersionInfo_Date() As String

    MVersionInfo_Date = "28-May-2011"

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "style.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

Private Function uGetStyleEngine(ByVal ClassId As String, Optional ByVal Flags As Long) As Boolean

    If mList.IndexOf(ClassId, False) <> 0 Then
        g_Debug "TStyleRoster.uGetStyleEngine(): '" & ClassId & "' is already in roster", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pse As TStyleEngine

    Set pse = New TStyleEngine
    pse.Init ClassId, Flags

    If mStartup.GetValueWithDefault(ClassId) <> "ignore" Then
        ' /* if it loads okay, update our version tracking record */
        If (pse.Load) And (pse.State = SN_SE_LOADED) Then _
            uCheckVersion pse.Engine, False

    End If

    ' /* add it to the engine list */

    mList.Add pse

    ' /* if it was loaded ok, add the styles it provides */

    If pse.State = SN_SE_LOADED Then _
        uAddStyles pse.Engine

    uGetStyleEngine = True

End Function

Private Sub uAddStyles(ByRef Engine As IStyleEngine)

    ' /* get the style names - **assumes engine has been intialised** */

    If Engine.CountStyles < 1 Then _
        Exit Sub

Dim pStyleInfo As style_info
Dim pStyle As TStyle
Dim i As Long

    For i = 1 To Engine.CountStyles
        Engine.StyleAt i, pStyleInfo
        If (pStyleInfo.Name <> "") And (IndexOf(pStyleInfo.Name) = 0) Then

            ' /* configure the TStyle object */
            Set pStyle = New TStyle
            pStyle.bInit Engine, pStyleInfo

            ' /* add it to the roster */
            uAddStyleToRoster pStyle

        Else
            g_Debug "uAddStyles(): style '" & pStyle.Name & "' is either invalid or already exists", LEMON_LEVEL_WARNING

        End If

    Next i

End Sub

Private Sub uAddStyleToRoster(ByRef Style As TStyle)
Dim i As Long
Dim j As Long

    ' /* add it alpha-sorted */

    If mItems Then
        For i = 1 To mItems
            If LCase$(Style.Name) < LCase$(mItem(i).Name) Then
                mItems = mItems + 1
                ReDim Preserve mItem(mItems)
                For j = mItems To (i + 1) Step -1
                    Set mItem(j) = mItem(j - 1)

                Next j

                Set mItem(i) = Style
                Exit Sub

            End If
        Next i
    End If

    mItems = mItems + 1
    ReDim Preserve mItem(mItems)
    Set mItem(mItems) = Style

    g_Debug "TStyleRoster.uAddStyleToRoster(): added style '" & Style.Name & "'"

End Sub

Public Sub Rewind()

    mIndex = 1

End Sub

Friend Function GetNextStyle(ByRef Style As TStyle) As Boolean

    If (mIndex < 1) Or (mIndex > mItems) Then _
        Exit Function

    Set Style = mItem(mIndex)
    mIndex = mIndex + 1
    GetNextStyle = True

End Function

Public Function CountStyles() As Long

    CountStyles = mItems

End Function

Friend Function StyleAt(ByVal Index As Long) As TStyle

    If (Index > 0) And (Index <= mItems) Then _
        Set StyleAt = mItem(Index)

End Function

Public Function GetStyleList(Optional ByRef Result As Boolean, Optional ByVal MarkDefault As Boolean = True) As String
Dim i As Long

    MarkDefault = False

    If mItems Then
        For i = 1 To mItems
            GetStyleList = GetStyleList & mItem(i).Name & IIf((MarkDefault) And (LCase$(mItem(i).Name) = uDefaultStyle), " (Default)", "") & "|"

        Next i

        GetStyleList = g_SafeLeftStr(GetStyleList, Len(GetStyleList) - 1)

    End If

    If GetStyleList = "" Then _
        GetStyleList = "<invalid list>"

    Result = (GetStyleList <> "<invalid list>")

End Function

'Public Function GetNonWindowStyleList(Optional ByRef Result As Boolean) As String
'Dim pt As TStyle
'Dim i As Long
'Dim j As Long
'
'    If mItems Then
'        For i = 1 To mItems
'            Set pt = mItem(i)
'            If (pt.Flags And S_STYLE_IS_WINDOWLESS) Then
'                If pt.CountSchemes Then
'                    For j = 1 To pt.CountSchemes
'                        GetNonWindowStyleList = GetNonWindowStyleList & pt.Name & ": " & pt.SchemeAt(j) & "#?" & _
'                                                LCase$(pt.Name & "/" & pt.SchemeAt(j)) & "|"
'
'                    Next j
'                End If
'            End If
'        Next i
'
'        GetNonWindowStyleList = g_SafeLeftStr(GetNonWindowStyleList, Len(GetNonWindowStyleList) - 1)
'
'    End If
'
'    If GetNonWindowStyleList = "" Then _
'        GetNonWindowStyleList = "<None>"
'
'    Result = (GetNonWindowStyleList <> "<None>")
'
'End Function

'Public Sub SetNonWindowStyleIcons(ByRef Control As BControl)
'
'    If (mItems = 0) Or (Control Is Nothing) Then _
'        Exit Sub
'
'Dim sz As String
'Dim i As Long
'Dim j As Long
'Dim n As Long
'
'    For i = 1 To mItems
'        If (mItem(i).Flags And S_STYLE_IS_WINDOWLESS) Then
'            sz = Me.StyleAt(i).IconPath
'            If Not g_Exists(sz) Then _
'                sz = g_MakePath(App.Path) & "etc\icons\style.png"
'
'            If mItem(i).CountSchemes Then
'                For j = 1 To mItem(i).CountSchemes
'                    n = n + 1
'                    prefskit_SetItem Control, n, "image-file", sz
'
'                Next j
'            End If
'        End If
'    Next i
'
'End Sub

Public Function IndexOf(ByVal Style As String) As Long
Dim i As Long

    ' /* expects just a style name - not a style/scheme pair! */

    If mItems = 0 Then _
        Exit Function

    Style = LCase$(Style)
    For i = 1 To mItems
        If LCase$(mItem(i).Name) = Style Then
            IndexOf = i
            Exit Function

        End If
    Next i

End Function

Public Function CountEngines() As Long

    CountEngines = mList.CountItems

End Function

Friend Function EngineAt(ByVal Index As Long) As TStyleEngine

    If (Index > 0) And (Index <= mList.CountItems) Then _
        Set EngineAt = mList.TagAt(Index)

End Function

Public Function Find(ByVal Name As String, ByRef Style As TStyle) As Boolean
Dim i As Long

    i = Me.IndexOf(LCase$(Name))
    If i Then
        Set Style = mItem(i)
        Find = True

    End If

End Function






'Friend Function RefreshStyle(ByVal Name As String, ByRef Style As T_SNARL_STYLE_2) As Boolean
'Dim i As Long
'
'    ' /* do we have the style? if not, can't do anything... */
'
'    i = Me.IndexOf(Name)
'    If i = 0 Then _
'        Exit Function
'
'Dim ps As SNARL_STYLE
'
'    ' /* got the style so call ISSEngine.GetStyle() but with the style's name
'    '    in SNARL_STYLE->Name.  This tells the engine that we're requesting
'    '    an update to a specific style; not just enumerating the list */
'
'    ps.Name = Name
'    If mItem(i).Engine.GetStyle(ps) Then
'        LSet mItem(i).Info = ps
'        uGetSchemes ps.Schemes, mItem(i)        ' // convert the schemes message
'        LSet Style = mItem(i)
'        RefreshStyle = True
'
'    End If
'
'End Function

Public Function Load(ByVal ClassId As String, ByVal NotifyVersionChange As Boolean, ByVal NotifyItHappened As Boolean) As Boolean
Dim pse As TStyleEngine

    If Not mList.Find(ClassId, pse) Then
        g_Debug "TStyleRoster.Load(): style engine '" & ClassId & "' not found", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    If pse.State = SN_SE_LOADED Then
        g_Debug "TStyleRoster.Load(): style engine '" & ClassId & "' is already running", LEMON_LEVEL_WARNING
        Exit Function

    End If

    If (pse.Load()) And (pse.State = SN_SE_LOADED) Then
        ' /* if it was loaded ok, add the styles it provides */
        uCheckVersion pse.Engine, NotifyVersionChange
        uAddStyles pse.Engine
        frmAbout.bNotifyStyleEnginesChanged

        uWriteStartupList
        If NotifyItHappened Then _
            g_PrivateNotify SNARL_CLASS_GENERAL, , ClassId & " started successfully", , ".good"

        Load = True

    Else
        g_Debug "TStyleRoster.Load(): style engine '" & ClassId & "' failed to load", LEMON_LEVEL_CRITICAL
        If NotifyItHappened Then _
            g_PrivateNotify SNARL_CLASS_GENERAL, , ClassId & " failed to start", , "!system-warning"

    End If

End Function

Public Function Unload(ByVal Engine As String, ByVal NotifyItHappened As Boolean) As Boolean
Dim pse As TStyleEngine

    If Not mList.Find(Engine, pse) Then
        g_Debug "TStyleRoster.Unload(): '" & Engine & "' not in list", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    If pse.IsSystem Then
        g_Debug "TStyleRoster.Unload(): cannot unload system style engine", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    If pse.State = SN_SE_NOT_LOADED Then
        g_Debug "TStyleRoster.Unload(): '" & Engine & "' is not running", LEMON_LEVEL_WARNING
        Exit Function

    End If

Dim j As Long

    ' /* find and remove all styles associated with this engine */

    If mItems Then
        For j = mItems To 1 Step -1
            If mItem(j).Engine.Name = Engine Then _
                uRemoveStyle j

        Next j
    End If

    ' /* zap and unload the engine code itself */

    pse.Unload

    If g_IsWinXPOrBetter() Then _
        CoFreeUnusedLibrariesEx 0, 0

    frmAbout.bNotifyStyleEnginesChanged

    uWriteStartupList
    If NotifyItHappened Then _
        g_PrivateNotify SNARL_CLASS_GENERAL, , Engine & " stopped successfully", , ".good"

    Unload = True

End Function

Private Sub uRemoveStyle(ByVal Index As Long)

    If (Index < 1) Or (Index > mItems) Then _
        Exit Sub

    ' /* trash the pointer to the style engine */

    mItem(Index).bZapEngine

Dim i As Long

    If Index < mItems Then
        For i = Index To (mItems - 1)
            Set mItem(i) = mItem(i + 1)

        Next i

    End If

    mItems = mItems - 1
    ReDim Preserve mItem(mItems)

End Sub

Private Sub uGetVersionList()
Dim szPath As String

    If Not g_GetUserFolderPath(szPath) Then _
        Exit Sub

    Set mVersions = New CConfFile
    If mVersions.SetTo(g_MakePath(szPath) & "etc\.styleengines") Then
        g_Debug "TStyleRoster.uGetVersionList(): version list loaded"

    Else
        g_Debug "TStyleRoster.uGetVersionList(): version list not found", LEMON_LEVEL_WARNING

    End If

End Sub

Private Sub uCheckVersion(ByRef Engine As IStyleEngine, ByVal Notify As Boolean)

    If (Engine Is Nothing) Or (mVersions Is Nothing) Then _
        Exit Sub

Dim szVerStr As String
Dim bChanged As Boolean
Dim sz As String

    ' /* 'CheckVersion' is set by snPrivateLoadExtension().  Here we check the current
    '    version of the extension compared to the one we're loading.  If the versions
    '    are different we show a notification telling the user that the extension
    '    was upgraded (or, possibly, downgraded) */

    szVerStr = CStr(Engine.Version) & "." & CStr(Engine.Revision)

    If Not mVersions.FindEntry(Engine.Name, sz) Then
        ' /* new */
        mVersions.Add Engine.Name, szVerStr
        mVersions.Save

        If Notify Then _
            g_PrivateNotify SNARL_CLASS_GENERAL, "Style engine installed", _
                            "Name: " & Engine.Name & vbCrLf & _
                            "Version: " & szVerStr, _
                            -1, _
                            g_MakePath(App.Path) & "etc\icons\style_engine.png"

    Else
        ' /* known */

'Dim nMaj As Long
'Dim nMin As Long
'Dim i As Long
'
'        ' /* decode version string */
'
'        i = InStr(sz, ".")
'        If i > 0 Then
'            nMaj = Val(g_SafeLeftStr(sz, i - 1))
'            nMin = Val(g_SafeRightStr(sz, Len(sz) - i))
'
'            Debug.Print nMaj & " > " & nMin

        If szVerStr <> sz Then
            Debug.Print "changed"

            If Notify Then
    
'                If (Engine.Version > nMaj) Or ((Engine.Version = nMaj) And (Engine.Revision > nMin)) Then
'                End If

                g_PrivateNotify SNARL_CLASS_GENERAL, "Style engine updated", _
                                "Name: " & Engine.Name & vbCrLf & _
                                "Previous version: " & sz & vbCrLf & _
                                "New version: " & szVerStr, _
                                , _
                                g_MakePath(App.Path) & "etc\icons\style_engine.png"

            End If

            mVersions.Update Engine.Name, szVerStr
            mVersions.Save

        End If

    End If

End Sub

Private Function uDefaultStyle() As String

    uDefaultStyle = LCase$(style_GetStyleName(g_ConfigGet("default_style")))

End Function

Private Function uDefaultScheme() As String

    uDefaultScheme = LCase$(style_GetSchemeName(g_ConfigGet("default_style")))

End Function

Public Function EngineList() As BTagList

    Set EngineList = mList

End Function

Private Sub uWriteStartupList()
Dim pse As TStyleEngine
Dim szPath As String

    If Not g_GetUserFolderPath(szPath) Then _
        Exit Sub

    With New CConfFile
        .SetTo g_MakePath(szPath) & "etc\styleengines.boot", False
        .MakeEmpty

        mList.Rewind
        Do While mList.GetNextTag(pse) = B_OK
            .Add pse.Name, IIf(pse.State = SN_SE_NOT_LOADED, "ignore", "load")

        Loop

        .Save
        g_Debug "TStyleRoster.uWriteStartupList(): startup list written to '" & .Filename & "'"

    End With

End Sub

Public Sub Restart()

    MLibrary_Close
    MLibrary_Open

End Sub

'Public Function GetNonWindowStyleList2(Optional ByRef Result As Boolean) As String
'Dim pt As TStyle
'Dim i As Long
'
'    If mItems Then
'        For i = 1 To mItems
'            Set pt = mItem(i)
'            If (pt.Flags And S_STYLE_IS_WINDOWLESS) Then _
'                GetNonWindowStyleList2 = GetNonWindowStyleList2 & pt.Name & "#?" & LCase$(pt.Name) & "|"
'
'        Next i
'
'        GetNonWindowStyleList2 = g_SafeLeftStr(GetNonWindowStyleList2, Len(GetNonWindowStyleList2) - 1)
'
'    End If
'
'    If GetNonWindowStyleList2 = "" Then _
'        GetNonWindowStyleList2 = "<None>"
'
'    Result = (GetNonWindowStyleList2 <> "<None>")
'
'End Function
'
'Public Sub SetNonWindowStyleIcons2(ByRef Control As BControl)
'
'    If (mItems = 0) Or (Control Is Nothing) Then _
'        Exit Sub
'
'Dim sz As String
'Dim i As Long
'Dim n As Long
'
'    For i = 1 To mItems
'        If (mItem(i).Flags And S_STYLE_IS_WINDOWLESS) Then
'            sz = Me.StyleAt(i).IconPath
'            If Not g_Exists(sz) Then _
'                sz = g_MakePath(App.Path) & "etc\icons\style.png"
'
'            n = n + 1
'            prefskit_SetItem Control, n, "image-file", sz
'
'        End If
'    Next i
'
'End Sub

Public Function GetSchemeList(ByVal Style As String) As String
Static i As Long

    i = uIndexOf(Style)
    If i = 0 Then _
        Exit Function

Static j As Long

    With mItem(i)
        If .CountSchemes Then
            For j = 1 To .CountSchemes
                GetSchemeList = GetSchemeList & .SchemeAt(j) & "#?" & LCase$(.SafeSchemeAt(j)) & "|"

            Next j
        End If
    End With

    GetSchemeList = g_SafeLeftStr(GetSchemeList, Len(GetSchemeList) - 1)

End Function

Private Function uIndexOf(ByVal Name As String) As Long
Static i As Long

    If mItems Then
        Name = LCase$(Name)
        For i = 1 To mItems
            If LCase$(mItem(i).Name) = Name Then
                uIndexOf = i
                Exit Function

            End If
        Next i
    End If

End Function

Friend Function RedirectTo(ByVal StyleAndScheme As String, ByRef Info As T_NOTIFICATION_INFO) As Boolean

    ' /* returns True if the notification should be reflected back to the screen */

    g_Debug "TStyleRoster.RedirectTo()", LEMON_LEVEL_PROC_ENTER
 
    ' /* if StyleAndScheme is empty, it's already been provided in Info */
    If StyleAndScheme = "" Then _
        StyleAndScheme = Info.StyleName & "/" & Info.SchemeName
 
    StyleAndScheme = LCase$(StyleAndScheme)
 
    ' /* get the appropriate instance for the notification before we do anything else
    '    we must pass a valid style instance to CSnarlWindow, even if the style
    '    roster is invalid */

Dim i As Long

    i = Me.IndexOf(style_GetStyleName(StyleAndScheme))
    If i = 0 Then
        g_Debug "redirector " & g_Quote(Info.StyleName) & " not found", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    ' /* some safety checks */

Dim ps As TStyle

    Set ps = mItem(i)
    If (ps.Flags And S_STYLE_IS_WINDOWLESS) = 0 Then
        g_Debug g_Quote(ps.Name) & " is not a redirector", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If
    
Dim psi As IStyleInstance
    
    Set psi = ps.CreateInstance()
    If ISNULL(psi) Then
        g_Debug "failed to create new instance of redirector " & g_Quote(Info.StyleName), LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        Exit Function

    End If

    ' /* get the path to the style - we need this for style-provided icons and sounds */

    Info.IconPath = g_TranslateIconPath(Info.IconPath, ps.Path)

Dim sTitle As String
Dim sText As String

    If (ps.Flags And S_STYLE_V42_CONTENT) Then
        g_Debug "wants unabridged content..."
        ' /* title becomes sender name, text is set to the original packed string */
        sTitle = Info.ClassObj.App.Name
        sText = Info.OriginalContent

    Else
        ' /* pre-V42 model: <title> is either title or app/title pair, <text> is text */
        g_Debug "wants translated content..."
        sText = Info.Text
        sTitle = Info.Title

        ' /* if S_STYLE_WANT_APP_NAME is set, prefix title with the sending app name */

        If (ps.Flags And S_STYLE_WANT_APP_NAME) Then
            g_Debug "wants application name..."
            sTitle = Info.ClassObj.App.Name & "#?" & sTitle

        End If
    End If

Dim pContent As notification_info

    g_Debug "sending to " & g_Quote(ps.Name) & "..."
    LSet pContent = ps.CreateNotificationInfo(sTitle, sText, Info.IconPath, _
                                              style_GetSchemeName(StyleAndScheme), _
                                              IIf(Info.Priority > 0, S_NOTIFICATION_IS_PRIORITY, 0))

    psi.UpdateContent pContent

    ' /* R2.4 Beta 4: if the style set notificaton_info->flags to -1 we should pass the notification
    '    back around for on-screen display */

    If pContent.Flags = -1 Then _
        g_Debug "redirect indicated we should reflect notification to screen..."

    RedirectTo = (pContent.Flags = -1)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function SafeGetStyle(ByVal StyleToUse As String, ByRef Style As TStyle, ByRef Instance As IStyleInstance) As Boolean

    ' /* try to create a new TStyle from the provided Style name */

Dim i As Long

    i = Me.IndexOf(style_GetStyleName(StyleToUse))
    If i Then
        Set Style = mItem(i)
        SafeGetStyle = True

    Else
        ' /* not found? use default */
        Set Style = New TStyle

    End If

    ' /* create a new instance from the style */

    Set Instance = Style.CreateInstance()

    ' /* didn't work? use an internal style instance */

    If (Instance Is Nothing) Then _
        Set Instance = New TInternalStyle

End Function

Public Function FindEngine(ByVal Name As String, ByRef Engine As TStyleEngine) As Boolean

    FindEngine = mList.Find(Name, Engine, False)

End Function

