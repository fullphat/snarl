VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRunFilePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRunFilePanel.cls
    '/
    '/  Description:    Provides the Run File style configuration panel
    '/
    '/  Created:        13-Jan-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mPanel As BPrefsPanel
Dim mSelected As String
Dim mSelIndex As Long

Dim WithEvents mAddEditPanel As TRunFileEditPanel
Attribute mAddEditPanel.VB_VarHelpID = -1

Dim mType As String
Dim mConfigFile As ConfigFile

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pc As BControl

    Select Case Control.GetName()
    Case "list"
        mSelIndex = Val(Value)
        mSelected = prefskit_GetItem(Control, "data")



    Case "add_remove"
        If Value = "+" Then
            uAddNewEntry

        Else
            If mSelIndex <> 0 Then _
                uRemEntry mSelIndex

        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim prf As TRunFileScheme
Dim i As Long

    Select Case Control.GetName
    Case "list"
        Select Case mType
        Case "r"
            ' /* run files */
            Set prf = gRunFiles.TagAt(Val(Control.GetValue))
            Set mAddEditPanel = New TRunFileEditPanel
            mAddEditPanel.Go mPanel.hWnd, prf

        Case "n"
            ' /* network forwarder */

        End Select



'    Case "test_forward"
''        i = mConfigFile.FindSection(mSelected)
''        If i Then
''            If Not g_SubsRoster.DoForward(mConfigFile.SectionAt(i)) Then _
''                MsgBox "Failed to forward notification.  Check the protocol and host names are valid.", vbExclamation Or vbOKOnly, "Network Forwarder"
''
''        End If


    Case "go_user_guide"
        ShellExecute Control.Page.Panel.hWnd, "open", g_MakePath(App.Path) & "etc\RunFileUserGuide\index.html", vbNullString, vbNullString, 5

    End Select

'        ShellExecute Control.Page.Panel.hWnd, "open", prf.File, vbNullString, vbNullString, 5

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
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

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Function Go(ByVal PanelType As String, ByVal Title As String, ByRef ConfigFile As ConfigFile) As Long

    Set mConfigFile = ConfigFile
    mType = PanelType

Dim pPage As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle Title & " Settings"
            .SetWidth 400
            .SetWindow 1

            Select Case mType

            Case "r"
                ' /* Run File */
                Set pPage = new_BPrefsPage("Templates", load_image_obj(g_MakePath(App.Path) & "etc\icons\runfile.png"), Me)
                With pPage
                    .SetMargin 0
                    Set pm = New CTempMsg
                    pm.Add "item-height", 36&
                    pm.Add "plus_minus", 1&
                    Set pc = new_BPrefsControl("listbox", "list", , , , pm)
                    pc.SizeTo 0, (6 * 36) + 2
                    .Add pc
                    .Add new_BPrefsControl("fancyplusminus", "add_remove")
'                .Add new_BPrefsControl("fancytoolbar", "toolbar_options", "Add...|Remove")

                End With
                .AddPage pPage

                Set pPage = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & "etc\icons\generic_about.png"), Me)
                With pPage
                    .SetMargin 0
    '                Set pm = New CTempMsg
    '                pm.Add "item-height", 36&
    '                pm.Add "plus_minus", 1&
    '                Set pc = new_BPrefsControl("listbox", "list", , , , pm)
    '                pc.SizeTo 0, (6 * 36) + 2
    '                .Add pc
    '                .Add new_BPrefsControl("fancyplusminus", "add_remove")
    '
    '
    ''                .Add new_BPrefsControl("fancytoolbar", "toolbar_options", "Add...|Remove")
    
                End With
                .AddPage pPage


'            Case "n"
'                ' /* Network Forwarder */
'
'                Set pPage = new_BPrefsPage("Destinations", load_image_obj(g_MakePath(App.Path) & "etc\icons\network.png"), Me)
'                With pPage
'                    .SetMargin 0
'                    Set pm = New CTempMsg
'                    pm.Add "item-height", 36&
'                    pm.Add "plus_minus", 1&
'                    Set pc = new_BPrefsControl("listbox", "list", , , , pm)
'                    pc.SizeTo 0, (6 * 36) + 2
'                    .Add pc
'                    .Add new_BPrefsControl("fancyplusminus", "add_remove")
'                    .Add new_BPrefsSeparator
'                    .Add new_BPrefsControl("fancybutton2", "test_forward", "Test")
'
'                End With
'                .AddPage pPage
'
'                Set pPage = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & "etc\icons\generic_about.png"), Me)
'                With pPage
'                    .SetMargin 0
'
'                End With
'                .AddPage pPage

            End Select
            
            .Go


            uRefreshList

        End With

    End If

    Go = mPanel.hWnd

End Function

Private Sub uAddNewEntry()
Dim prf As TRunFileScheme

    Select Case mType
    Case "r"
        Set prf = New TRunFileScheme
        If prf.CreateNew Then
            Set mAddEditPanel = New TRunFileEditPanel
            mAddEditPanel.Go mPanel.hWnd, prf
            g_RunFileLoadSchemes
    
        End If

    End Select

    uRefreshList

End Sub

Private Sub uRemEntry(ByVal Index As Long)
Dim prf As TRunFileScheme

    Select Case mType
    Case "r"
        Set prf = gRunFiles.TagAt(Index)
        DeleteFile prf.File
        gRunFiles.Remove Index

    End Select

    uRefreshList

End Sub

Private Sub uRefreshList()
Dim pc As BControl

    If Not mPanel.Find("list", pc) Then _
        Exit Sub

Dim n As Long

    n = Val(pc.GetValue())

    Select Case mType
    Case "r"
        uGetCurrentRunFiles pc

'    Case "n"
'        uGetForwards pc

    End Select

    pc.SetValue CStr(Max(n, 1))

End Sub

Private Sub mAddEditPanel_Done()

    uRefreshList
    
End Sub

Private Sub uGetCurrentRunFiles(ByRef ListControl As BControl)
Dim szType As String
Dim pt As BTagItem
Dim sz As String
Dim i As Long

    ' /* set content */

    With gRunFiles
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            Select Case pt.Value
            Case "1"
                szType = "Static template"

            Case "2"
                szType = "Custom template"

            Case "3"
                szType = "Unabridged"

            Case Else
                szType = "Unknown type"

            End Select

            sz = sz & pt.Name & "#?0#?" & szType & "|"

        Loop
    End With

    sz = g_SafeLeftStr(sz, Len(sz) - 1)
    ListControl.SetText sz

    If sz <> "" Then
        ' /* set icons */
        With gRunFiles
            .Rewind
            Do While .GetNextTag(pt) = B_OK
                i = i + 1

                Select Case pt.Value
                Case "1"
                    szType = "-static"
    
                Case "2"
                    szType = "-custom"
                    
                Case "3"
                    szType = "-unabridged"
    
                Case Else
                    szType = ""
    
                End Select

                prefskit_SetItemObject ListControl, i, "image-object", load_image_obj(g_MakePath(App.Path) & "etc\icons\runfile" & szType & ".png")
    
            Loop
        End With
    End If

End Sub

'Private Sub uGetForwards(ByRef ListControl As BControl)
'Dim ps As ConfigSection
'Dim szList As String
'
'    With mConfigFile
'        .Load
'        .Rewind
'
'        Do While .GetNextSection(ps)
'            If ps.GetValueWithDefault("type", "") = "forwarder" Then
'                ' /* name, hosts, protocol */
'                If (ps.GetValueWithDefault("name", "") <> "") And (ps.Name <> "") Then _
'                    szList = szList & ps.GetValueWithDefault("name", "") & "#?" & ps.Name & "#?" & _
'                                      "To " & ps.GetValueWithDefault("hosts", "<none>") & " using " & ps.GetValueWithDefault("protocol", "<unknown protocol>") & "|"
'
'            End If
'        Loop
'
'        szList = g_SafeLeftStr(szList, Len(szList) - 1)
'        ListControl.SetText szList
'
'Dim szIcon As String
'Dim i As Long
'
'        If szList <> "" Then
'            .Rewind
'            Do While .GetNextSection(ps)
'                If ps.GetValueWithDefault("type", "") = "forwarder" Then
'                    If (ps.GetValueWithDefault("name", "") <> "") And (ps.Name <> "") Then
'                        i = i + 1
'                        szIcon = LCase$(ps.GetValueWithDefault("protocol", ""))
'                        Select Case szIcon
'                        Case "snp", "gntp"
'
'                        Case Else
'                            szIcon = "unknown"
'
'                        End Select
'
'                        prefskit_SetItem ListControl, i, "image-file", g_MakePath(App.Path) & "etc\icons\protocol-" & szIcon & ".png"
'
'                    End If
'                End If
'            Loop
'        End If
'    End With
'
'End Sub
