VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type T_INFO
    Engine As IStyleEngine
    SchemeCount As Long
    Scheme() As String
    Info As style_info

End Type

Dim mContent As T_INFO
Dim mhWndOwner As Long
Dim mhWndConfig As Long
Dim mIcon As MImage

Public Event StylePrefsWindowClosed()

Implements BWndProcSink

Private Sub Class_Initialize()

    Set mContent.Engine = New TInternalStyle
    mContent.Info.Path = g_MakePath(App.Path) & "etc\default_theme"

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    On Error Resume Next

    If uMsg = WM_DESTROY Then
        ' /* config window has gone */
        window_subclass hWnd, Nothing
        If mhWndOwner <> 0 Then
'            EnableWindow mhWndOwner, -1
            g_WindowToFront mhWndOwner, True
            mhWndConfig = 0
            mhWndOwner = 0

        End If
    
        RaiseEvent StylePrefsWindowClosed

    End If

End Function

Friend Sub bInit(ByRef Engine As IStyleEngine, ByRef Info As style_info)
Dim sz As String

    With mContent
        Set .Engine = Engine
        LSet .Info = Info

    End With

    ' /* get the schemes */
    uGetSchemes

    ' /* cache the icon */
    sz = Info.IconPath
    If Not g_Exists(sz) Then _
        sz = g_MakePath(Info.Path) & Info.Name & ".png"

    If Not g_Exists(sz) Then _
        sz = g_MakePath(Info.Path) & "icon.png"

    ' /* use default icon */
    If Not g_Exists(sz) Then
        If Me.IsRedirect Then
            sz = g_MakePath(App.Path) & "etc\icons\forward.png"

        Else
            sz = g_MakePath(App.Path) & "etc\icons\style.png"
        
        End If

    End If

    Set mIcon = load_image_obj(sz)

End Sub

Friend Sub bZapEngine()

    Set mContent.Engine = Nothing

End Sub

Private Sub uGetSchemes()
Dim sz() As String
Dim c As Long
Dim i As Long

    On Error Resume Next

    With mContent
        .SchemeCount = 0
        ReDim .Scheme(0)

    End With

    sz() = Split(mContent.Info.Schemes, "|")
    c = UBound(sz()) + 1
    If c > 0 Then
        For i = 0 To (c - 1)
            sz(i) = trim(sz(i))
            If sz(i) <> "" Then
                With mContent
                    .SchemeCount = .SchemeCount + 1
                    ReDim Preserve .Scheme(.SchemeCount)
                    .Scheme(.SchemeCount) = sz(i)

                End With
            End If
        Next i
    End If

    With mContent
        If .SchemeCount = 0 Then
            g_Debug "TStyle.uGetSchemes(): style '" & mContent.Info.Name & "' didn't provide any valid schemes", LEMON_LEVEL_WARNING
            .SchemeCount = 1
            ReDim .Scheme(.SchemeCount)
            .Scheme(.SchemeCount) = "<Default>"

        End If

    End With

End Sub

'Public Function Info() As style_info
'
'    LSet Info = mInfo
'
'End Function

Public Function CreateInstance() As IStyleInstance

    If Not (mContent.Engine Is Nothing) Then _
        Set CreateInstance = mContent.Engine.CreateInstance(mContent.Info.Name)

End Function

Public Function CreateNotificationInfo(ByVal Title As String, ByVal Text As String, ByVal Icon As String, ByVal Scheme As String, ByVal Flags As S_NOTIFICATION_FLAGS) As notification_info
Dim pInfo As notification_info

    With pInfo
        .Title = Title
        .Text = Text
        .Icon = Icon
        .Scheme = Scheme
        .Flags = Flags

    End With

    LSet CreateNotificationInfo = pInfo

End Function

Public Function CountSchemes() As Long

    CountSchemes = mContent.SchemeCount

End Function

Public Function SchemeAt(ByVal Index As Long) As String

    If (Index > 0) And (Index <= mContent.SchemeCount) Then _
        SchemeAt = mContent.Scheme(Index)

End Function

Public Function IconPath() As String

    IconPath = mContent.Info.IconPath

    ' /* backwards compatability */

    If Not g_Exists(IconPath) Then _
        IconPath = g_MakePath(mContent.Info.Path) & mContent.Info.Name & ".png"

    If Not g_Exists(IconPath) Then _
        IconPath = g_MakePath(mContent.Info.Path) & "icon.png"

    ' /* still no good? let caller decide what to do */

    If Not g_Exists(IconPath) Then _
        IconPath = ""

End Function

Public Function Name() As String

    Name = mContent.Info.Name

End Function

Public Function Flags() As S_STYLE_FLAGS

    Flags = mContent.Info.Flags

End Function

Public Function SchemeIndex(ByVal Scheme As String) As Long
Dim i As Long

    Scheme = LCase$(Scheme)

    With mContent
        If .SchemeCount Then
            For i = 1 To .SchemeCount
                If LCase$(.Scheme(i)) = Scheme Then
                    SchemeIndex = i
                    Exit Function

                End If
            Next i
        End If
    End With

End Function
'
'Friend Function IndexOfScheme2(ByVal StyleIndex As Long, ByVal Scheme As String) As Long
'Dim i As Long
'
'    If (StyleIndex < 1) Or (StyleIndex > mItems) Then _
'        Exit Function
'
'    Scheme = LCase$(Scheme)
'
'    With mItem(StyleIndex)
'        If .SchemeCount Then
'            For i = 1 To .SchemeCount
'                If LCase$(.Scheme(i)) = Scheme Then
'                    IndexOfScheme2 = i
'                    Exit Function
'
'                End If
'            Next i
'        End If
'    End With
'
'End Function


Public Function Path() As String

    Path = g_MakePath(mContent.Info.Path)

End Function

Public Function Description() As String

    Description = mContent.Info.Description

End Function

Public Function Engine() As IStyleEngine

    Set Engine = mContent.Engine

End Function

Public Function Copyright() As String

    Copyright = mContent.Info.Copyright

End Function

Public Function URL() As String

    URL = mContent.Info.URL

End Function

Public Function Release() As String

    Release = CStr(mContent.Info.Major) & "." & CStr(mContent.Info.Minor)

End Function

Public Function GetSchemeList(Optional ByVal Refresh As Boolean) As String

    If Refresh Then
        uUpdateStyleInfo
        uGetSchemes

    End If

Dim i As Long

    With mContent
        If .SchemeCount Then
            For i = 1 To .SchemeCount
                GetSchemeList = GetSchemeList & IIf(.Scheme(i) = "", "<default>", .Scheme(i)) & "#?" & LCase$(Me.SafeSchemeAt(i)) & "|"

            Next i
        End If
    End With

    GetSchemeList = g_SafeLeftStr(GetSchemeList, Len(GetSchemeList) - 1)

    If GetSchemeList = "" Then _
        GetSchemeList = "<Default>"

End Function

Private Sub uUpdateStyleInfo()
Dim ps As style_info
Dim i As Long

    With mContent.Engine
        For i = 1 To .CountStyles
            .StyleAt i, ps
            If ps.Name = mContent.Info.Name Then
                LSet mContent.Info = ps
                Exit Sub

            End If
        Next i

    End With

    Debug.Print "TStyle.uUpdateStyleInfo(): couldn't find style '" & mContent.Info.Name & "'"

End Sub

Public Function SafeSchemeAt(ByVal Index As Long) As String

    If (Index > 0) And (Index <= mContent.SchemeCount) Then _
        SafeSchemeAt = mContent.Scheme(Index)

    If SafeSchemeAt = "<Default>" Then _
        SafeSchemeAt = ""

End Function

Public Function IsRedirect() As Boolean

    If mContent.Info.Flags And S_STYLE_IS_WINDOWLESS Then _
        IsRedirect = ((mContent.Info.Flags And S_STYLE_REDIRECT_TO_SCREEN) = 0)

End Function

Public Function ShowPrefs(ByVal hWndOwner As Long) As Boolean

    If ISNULL(mContent.Engine) Then _
        Exit Function

    If mhWndConfig <> 0 Then
        g_WindowToFront mhWndConfig, True
        ShowPrefs = True
        Exit Function

    End If


Dim hWndConfig As Long

    hWndConfig = mContent.Engine.GetConfigWindow(mContent.Info.Name)
    If IsWindow(hWndConfig) = 0 Then
        g_Debug "TStyle.ShowPrefs(): GetConfigWindow() returned invalid window " & g_HexStr(hWndConfig), LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* valid window */

    window_subclass hWndConfig, Me
    mhWndConfig = hWndConfig

    If hWndOwner <> 0 Then
        ' /* centre it on its owner */
        mhWndOwner = g_GetTopLevel(hWndOwner)
        SetWindowLong hWndConfig, GWL_HWNDPARENT, mhWndOwner
        g_MoveWindow hWndConfig, , , , , True, mhWndOwner

    Else
        g_MoveWindow hWndConfig, , , , True

    End If

    g_ShowWindow hWndConfig, True, True
    ShowPrefs = True

End Function

Public Function SafeIconPath() As String

    SafeIconPath = Me.IconPath
    If Not g_Exists(SafeIconPath) Then _
        SafeIconPath = g_MakePath(App.Path) & "etc\icons\style.png"

End Function

Public Function NameEx() As String
Dim sz As String

'    sz = mContent.Info.Copyright
'    If sz <> "" Then _
'        sz = " (" & sz & ")"

    NameEx = mContent.Info.Name & " " & CStr(mContent.Info.Major) & "." & CStr(mContent.Info.Minor) & sz

End Function

Public Function DoSchemePreview(ByVal Scheme As String, ByVal IsPriority As Boolean, Optional ByVal Percent As Integer = -1, Optional ByVal HidePrev As Boolean = True) As Boolean

    ' /* this handles external requests to Snarl to display a notification in a particular
    '    style and scheme - only the SNARL_PREVIEW_SCHEME message handler calls this */

    If (g_NotificationRoster Is Nothing) Then _
        Exit Function

    If Scheme = "" Then
        ' /* if no scheme, use "<Default>" */
        Scheme = "<Default>"

    Else
        ' /* otherwise, supplied scheme must exist */
        If Me.SchemeIndex(Scheme) = 0 Then _
            Exit Function

    End If

Dim pInfo As T_NOTIFICATION_INFO

    With pInfo
        .Title = Me.Name & IIf(Scheme = "<Default>", "", "/" & Scheme) & IIf(IsPriority, " (Priority)", "")
        .Text = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
        .Timeout = -1
        .IconPath = Me.SafeIconPath()
        .StyleName = mContent.Info.Name
        .SchemeName = LCase$(Scheme)
        .Position = SN_SP_DEFAULT_POS
        .Priority = IIf(IsPriority, 1, 0)
        Set .ClassObj = New TAlert
        .APIVersion = App.Major
        .IntFlags = App.Major
        
        If HidePrev Then _
        .CustomUID = "style-preview" & IIf(IsPriority, "-priority", "")
        
        With New BPackedData
            If Percent > -1 Then _
                .Add "value-percent", CStr(Percent)
                
            .Add "value-date-packed", Format$(Now(), "YYYYMMDDHHNNSS")

            pInfo.OriginalContent = .AsString()

        End With

        If HidePrev Then _
            g_NotificationRoster.Hide 0, .CustomUID, App.ProductName, ""

    End With

    DoSchemePreview = (g_NotificationRoster.Add(pInfo, Nothing, False) <> 0)

End Function

Public Function SchemeIconPath(ByVal Index As Long) As String
Dim szf As String
Dim sz As String

    If (Index > 0) And (Index <= mContent.SchemeCount) Then
        sz = g_MakePath(mContent.Info.Path)
        szf = g_MakeFilename(mContent.Scheme(Index))
        SchemeIconPath = sz & szf & ".icon"
        If g_Exists(SchemeIconPath) Then _
            Exit Function

        SchemeIconPath = sz & szf & ".png"
        If g_Exists(SchemeIconPath) Then _
            Exit Function

    End If

    SchemeIconPath = mContent.Info.IconPath

    If Not g_Exists(SchemeIconPath) Then
        If Me.IsRedirect Then
            SchemeIconPath = g_MakePath(App.Path) & "etc\icons\forward.png"

        Else
            SchemeIconPath = g_MakePath(App.Path) & "etc\icons\style.png"
        
        End If
    End If

End Function

Public Function SchemeIconIcon(ByVal Index As Long) As MImage

    Set SchemeIconIcon = load_image_obj(Me.SchemeIconPath(Index))

End Function

Public Function IsConfigurable() As Boolean

    IsConfigurable = ((mContent.Info.Flags And S_STYLE_IS_CONFIGURABLE) <> 0)

End Function

Public Function Icon() As MImage

    Set Icon = mIcon

End Function
