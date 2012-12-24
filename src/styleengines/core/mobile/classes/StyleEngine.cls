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

    ' /*
    '
    '   StyleEngine.cls --
    '
    '   © 2010 full phat products.  Some Rights Reserved.
    '   Licenced as per the Simplified BSD (http://www.opensource.org/licenses/bsd-license.php)
    '
    '   Notes
    '   -----
    '
    '   Standardised V42 entries accepted:
    '
    '   Generic groups
    '   --------------
    '   value-:         must be a numeric value, may be a decimal or integer
    '   label-:         may be alphanumeric
    '   colour-:        may use HTML notation (i.e. #RRGGBB) or rgb notation (i.e. r,g,b)
    '   data-:          free for custom use
    '
    '   Defined Entries
    '   ---------------
    '   value-percent:      must be a numeric value in the range 0 to 100.  May be fractional although styles may round to
    '                       nearest integer at their discretion.
    '
    '   label-subtext:      text which is typically displayed inferior to the notification text.
    '   label-tagline:      a short piece of text typically displayed in a prominent way.
    '
    '   colour-background:  requested background colour.
    '   colour-text:        requested text colour.
    '   colour-subtext:     requested subtext colour.
    '
    ' */

Private Const LIB_DATE = "21-Feb-2012"

Private Const STYLE_NAME_TOASTY = "Toasty"
Private Const STYLE_NAME_NMA = "Notify My Android"
Private Const STYLE_NAME_WEB_FORWARD = "WebForward"
Private Const STYLE_NAME_OUTLOOK_LOG = "Outlook Logger"

Dim mLastErr As String
Dim mStyle() As TStyle
Dim mCount As Long

Dim mPanel As BPrefsPanel
Dim mPage As BPrefsPage
Dim mNoWriteConfig As Boolean

Dim mCurrent As TStyle
Dim mWebForward As BTagList

Dim WithEvents theAddEditPanel As TWebForwardAddEdit
Attribute theAddEditPanel.VB_VarHelpID = -1

Implements KPrefsPage
Implements KPrefsPanel
Implements IStyleEngine

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = mCount

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance
Dim pWF As TWebForwardInstance
Dim px As TGenericInstance

    Select Case StyleName
    Case STYLE_NAME_TOASTY
        Set px = New TGenericInstance
        px.SetStyle mStyle(uIndexOf(StyleName))
        Set IStyleEngine_CreateInstance = px
        Exit Function

    Case STYLE_NAME_NMA
        Set px = New TGenericInstance
        px.IsNotifyMyAndroid = True
        px.SetStyle mStyle(uIndexOf(StyleName))
        Set IStyleEngine_CreateInstance = px
        Exit Function

    Case STYLE_NAME_WEB_FORWARD
        Set pWF = New TWebForwardInstance
        pWF.Init mWebForward
        Set IStyleEngine_CreateInstance = pWF
        Exit Function

    Case STYLE_NAME_OUTLOOK_LOG
        Set IStyleEngine_CreateInstance = New TOutlookLogInstance

    End Select

'    If Not (pStyle Is Nothing) Then
'        pStyle.SetStyle mStyle(uIndexOf(StyleName))
'        Set IStyleEngine_CreateInstance = pStyle
'
'    End If

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = "Provides a suite of styles that send notification content via different Internet transport protocols and APIs"

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long
Dim pp As BPrefsPage
Dim pm As CTempMsg
Dim pc As BControl

    If (mPanel Is Nothing) Then
        Set mCurrent = mStyle(uIndexOf(StyleName))
        If (mCurrent Is Nothing) Then
            MsgBox "Error setting style", vbCritical, App.Title
            Exit Function

        End If

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetWindow 1

            If mCurrent.Name = STYLE_NAME_WEB_FORWARD Then
                .SetWidth 380
                .SetTitle "WebForward 2.0"
                Set pp = new_BPrefsPage("Forwards", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "webforward.png"), Me)
                With pp
                    .SetMargin 0
                    Set pm = New CTempMsg
                    pm.Add "item-height", 36&
                    pm.Add "plus_minus", 1&
                    Set pc = new_BPrefsControl("listbox", "wf_list", , , , pm)
                    pc.SizeTo 0, (7 * 36) + 2
                    .Add pc
                    .Add new_BPrefsControl("fancyplusminus", "wf_add_remove")
                End With
                .AddPage pp

                Set pp = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "about.png"), Me)
                With pp
                    .SetMargin 24
                    .Add new_BPrefsControl("banner", "", "The Concept")
                    .Add new_BPrefsControl("label", "", "This style allows you to forward the content of notifications to web-based services which then typically display the notification on a device such as a mobile phone or tablet.  Each scheme is an individual WebForward, allowing for granular control over which notifications are sent where.")
                    .Add new_BPrefsControl("banner", "", "Quick Guide")
                    .Add new_BPrefsControl("label", "", "Each WebForward consists of a base URL and an optional query which are combined to form a GET or POST HTTP request.  If the POST method is used, the request body can also be specified.")
                    .Add new_BPrefsControl("label", "", "Any elements enclosed with '#' characters will be replaced with the correspondingly named notification component (e.g. #title# will be replaced with the notification title).  See the User Guide for more information.")
                    .Add new_BPrefsControl("fancybutton2", "do_read_more", "User Guide")
                End With
                .AddPage pp

            Else
                .SetWidth 300
                .SetTitle StyleName & " Preferences"
                Set pp = new_BPrefsPage("", , Me)
                With pp
                    Select Case mCurrent.Name
                    Case STYLE_NAME_NMA
                        .SetMargin 80
                        .Add new_BPrefsControl("fancyedit", "api_key", mCurrent.ReadSetting("api_key"), "API keys:")
                        .Add new_BPrefsControl("label", "", "Multiple keys can be provided, seperate each key with a comma.")
    
                    Case STYLE_NAME_TOASTY
                        .SetMargin 80
                        .Add new_BPrefsControl("fancyedit", "device_id", mCurrent.ReadSetting("device_id"), "Device ID:")

                    End Select
                End With

                .AddPage pp

            End If
            
            .Go

'    g_SetWindowIconToAppResourceIcon mPanel.hWnd

        End With

    End If

    IStyleEngine_GetConfigWindow = mPanel.hWnd

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    gID = 63

    ' /* create the user folder */

    create_directory_exists g_MakePath(style_GetSnarlStylesPath()) & "webforward"

    ' /* add our styles */

    uAdd STYLE_NAME_TOASTY, _
         "Forwards the notification content to a Windows Phone 7 device via Toasty.  See http://supertoasty.com/ for more information.", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_IS_WINDOWLESS Or S_STYLE_WANT_APP_NAME Or S_STYLE_V42_CONTENT, _
         "", _
         1, 1

    uAdd STYLE_NAME_NMA, _
         "Forwards the notification content to an Android device via Notify My Android.  You will need to register an API key at http://nma.usk.bz first.", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_IS_WINDOWLESS Or S_STYLE_WANT_APP_NAME, _
         "", _
         1, 1

    uAdd STYLE_NAME_WEB_FORWARD, _
         "Forwards the notification content to a specific URL using a customizable template.", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_IS_WINDOWLESS Or S_STYLE_V42_CONTENT, _
         uGetWebForwardSchemes(), _
         2, 0

    uAdd STYLE_NAME_OUTLOOK_LOG, _
         "Creates the notification as an email item in the Inbox in Microsoft Outlook.  Note that Outlook must be running for the item to be created successfully.", _
         S_STYLE_MULTI_INSTANCE Or S_STYLE_IS_WINDOWLESS Or S_STYLE_V42_CONTENT, _
         "Inbox", _
         0, 1


    IStyleEngine_Initialize = M_OK

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

    If (Index < 1) Or (Index > mCount) Then _
        Exit Sub

    mStyle(Index).GetInfo Style

    If Style.Name = STYLE_NAME_WEB_FORWARD Then _
        Style.Schemes = uGetWebForwardSchemes()

End Sub

Private Sub IStyleEngine_TidyUp()
End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function

Private Sub KPrefsPage_AllAttached()

    uUpdateList

End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim sz As String
Dim i As Long

    Debug.Print Control.GetName() & " > " & Value

    Select Case Control.GetName
    Case "wf_add_remove"
        ' /* webforwarder +/- button */
        Select Case Value
        Case "-"
            ' /* remove selected item */
            i = Val(prefskit_GetValue(Control.Page.Panel, "wf_list"))
            If i <> 0 Then
                DeleteFile mWebForward.TagAt(i).Name                                ' // delete the file
                uGetWebForwardSchemes                                               ' // reload schemes
                uUpdateList

            End If

        Case "+"
            ' /* open add panel */
            Set theAddEditPanel = New TWebForwardAddEdit
            theAddEditPanel.Go mPanel.hWnd, Nothing

        End Select

        Exit Sub

    End Select

    ' /* write settings back to config file */

    mCurrent.WriteSetting Control.GetName, Value

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pf As CConfFile2
Dim i As Long

    Select Case Control.GetName
    Case "wf_list"
        i = Val(Control.GetValue)
        If i <> 0 Then
            Set pf = New CConfFile2
            pf.SetTo mWebForward.TagAt(i).Name
            Set theAddEditPanel = New TWebForwardAddEdit
            theAddEditPanel.Go mPanel.hWnd, pf

        End If

'            ShellExecute Control.Page.Panel.hWnd, "open", , vbNullString, vbNullString, 5

    Case "do_read_more"
        ShellExecute Control.Page.Panel.hWnd, "open", g_MakePath(App.Path) & "userguide\index.html", vbNullString, vbNullString, 5

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
'    mPage.SetMargin 140

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing
    Set mCurrent = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub uAdd(ByVal StyleName As String, ByVal Description As String, ByVal Flags As S_STYLE_FLAGS, ByVal Schemes As String, ByVal Major As Long, ByVal Minor As Long) ', ByRef Defaults As BPackedData, Optional ByRef CustomPrefsPage As KPrefsPage)
Dim ps As TStyle

    Set ps = New TStyle
    ps.Init StyleName, Description, Flags, Schemes, Major, Minor, Nothing, LIB_DATE ', CustomPrefsPage

    mCount = mCount + 1
    ReDim Preserve mStyle(mCount)
    Set mStyle(mCount) = ps

End Sub

Private Function uIndexOf(ByVal Name As String) As Long
Static i As Long

    For i = 1 To mCount
        If mStyle(i).Name = Name Then
            uIndexOf = i
            Exit Function

        End If
    Next i

End Function

'Private Function uCreatePackedStr(ByVal Content As String) As BPackedData
'
'    Set uCreatePackedStr = New BPackedData
'    uCreatePackedStr.SetTo Content
'
'End Function

Private Function uGetWebForwardSchemes() As String
Dim pf As CConfFile2
Dim sz As String

    Set mWebForward = new_BTagList()

    With New CFolderContent2
        If .SetTo(g_MakePath(style_GetSnarlStylesPath()) & "webforward") Then
            .Rewind
            Do While .GetNextFile(sz)
                Debug.Print "webforward: " & sz
                If g_GetExtension(sz, True) = "webforward" Then
                    g_Debug "uGetWebForwardSchemes(): checking '" & sz & "'..."
                    Set pf = New CConfFile2
                    If pf.SetTo(sz, True) Then
                        sz = pf.GetValueWithDefault("name")
                        If sz <> "" Then
                            Debug.Print "name=" & sz
                            mWebForward.Add pf
                            uGetWebForwardSchemes = uGetWebForwardSchemes & sz & "|"

                        Else
                            g_Debug "uGetWebForwardSchemes(): ...invalid content"

                        End If

                    Else
                        g_Debug "uGetWebForwardSchemes(): ...invalid file"

                    End If

                End If
            Loop

        End If

    End With

    uGetWebForwardSchemes = g_SafeLeftStr(uGetWebForwardSchemes, Len(uGetWebForwardSchemes) - 1)

'    If uGetWebForwardSchemes = "" Then _
        uGetWebForwardSchemes = "(None defined)"

End Function

'Private Function uSchemesList() As String
'
'    If (mWebForward Is Nothing) Then _
'        Exit Function
'
'Dim pf As CConfFile2
'
'    With mWebForward
'        .Rewind
'        Do While .GetNextTag(pf) = B_OK
'            uSchemesList = uSchemesList & pf.GetValueWithDefault("name") & "#?0#?" & g_SafeLeftStr(pf.GetValueWithDefault("url") & _
'                           Replace$(pf.GetValueWithDefault("template"), "|", "¦"), 70, True) & "|"
'
'        Loop
'
'    End With
'
'    uSchemesList = g_SafeLeftStr(uSchemesList, Len(uSchemesList) - 1)
'
'End Function

'Private Function uCreateNewScheme() As String
'Dim sz As String
'Dim sz2 As String
'Dim i As Long
'
'    sz = g_MakePath(style_GetSnarlStylesPath()) & "webforward"
'    If Not g_Exists(sz) Then _
'        Exit Function
'
'    sz = sz & "\"
'
'    Do
'        i = i + 1
'        sz2 = sz & "NewForwarder" & CStr(i) & ".webforward"
'
'    Loop While g_Exists(sz2)
'
'    On Error Resume Next
'
'Dim n As Integer
'
'    n = FreeFile()
'    Open sz2 For Output As #n
'
'    Print #n, "name=New Forwarder " & CStr(i)
'    Print #n, "url=<server_url_here>"
'    Print #n, "template=<forwarder_template_here>"
'
'    Close #n
'
'    uCreateNewScheme = sz2
'
'End Function

Private Sub theAddEditPanel_Quit()

    Set theAddEditPanel = Nothing
    uGetWebForwardSchemes                                       ' // reload schemes
    uUpdateList

End Sub

Private Sub uUpdateList()
Dim pc As BControl

    If (Not mPanel.Find("wf_list", pc)) Or (ISNULL(mWebForward)) Then _
        Exit Sub

Dim pf As CConfFile2
Dim szIcon As String
Dim sz As String
Dim i As Long

    With mWebForward
        .Rewind
        Do While .GetNextTag(pf) = B_OK
            sz = sz & pf.GetValueWithDefault("name") & "#?0#?" & g_SafeLeftStr(pf.GetValueWithDefault("url") & _
                      Replace$(pf.GetValueWithDefault("template"), "|", "¦"), 70, True) & "|"

        Loop

        sz = g_SafeLeftStr(sz, Len(sz) - 1)
        pc.SetText sz

        If sz <> "" Then
            .Rewind
            Do While .GetNextTag(pf) = B_OK
                i = i + 1
                szIcon = g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & g_MakeFilename2(pf.GetValueWithDefault("name")) & ".icon"
                If Not g_Exists(szIcon) Then _
                    szIcon = g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "webforward.png"

                prefskit_SetItem pc, i, "image-file", szIcon
            
            Loop
        End If
    End With

End Sub
