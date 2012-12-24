VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWindowSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Dim mWindows As BTagList
Dim mRules As BTagList

Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Public Event WindowAppeared(ByRef MatchingRule As TRule, ByVal Title As String, ByVal Class As String, ByVal hWnd As Long)
Public Event WindowDisappeared(ByRef MatchingRule As TRule, ByVal Title As String, ByVal Class As String, ByVal hWnd As Long)

Private Sub Class_Initialize()

    Set mWindows = new_BTagList()
    Set mRules = new_BTagList()

End Sub

Private Sub Class_Terminate()

    Set theTimer = Nothing

End Sub

Private Sub theTimer_Pulse()
'Dim i As Long
'
'    ' /* check gone windows */
'
'    With mWindows
'        If .CountItems Then
'            For i = .CountItems To 1 Step -1
'                If Not g_IsAppWindow(Val(.TagAt(i).Name), True) Then
'                    ' /* window gone */
'                    uCompare .TagAt(i), False
'
'
''                    RaiseEvent WindowDisappeared(Val(.TagAt(i).Name))
'
''                    ' /* remove the notification if it's still around */
''
''                    If mToken <> 0 Then _
''                        snDoRequest "hide?app-sig=" & App.ProductName & _
''                                    "&uid=" & .TagAt(i).Name
''
'''                    If mToken <> 0 Then _
'''                        snDoRequest "notify?app-sig=" & App.ProductName & _
'''                                    "&class=" & WINDOW_DISAPPEARED & _
'''                                    "&uid=" & .TagAt(i).Name & _
'''                                    "&replace-uid=" & .TagAt(i).Name & _
'''                                    "&title=Window disappeared" & _
'''                                    "&text=" & .TagAt(i).Value & _
'''                                    "&icon=" & g_MakePath(App.Path) & "gone.png"
''
'                    .Remove i
'
'                End If
'            Next i
'        End If
'
'    End With
'
'    ' /* check new windows */
'
'Dim pWindow As TWindow
'Dim lPid As Long
'Dim h As Long
'Dim c As Long
'
'    c = g_AppWindows_Refresh(True)
'
'    If c Then
'        For i = 1 To c
'            h = g_AppWindows_WindowAt(i)
'            GetWindowThreadProcessId h, lPid
'
'            If (lPid <> GetCurrentProcessId()) And (Not mWindows.Find(CStr(h), Nothing)) Then
'                ' /* if we don't already have it, and it wasn't created by us, add it */
'                Set pWindow = New TWindow
'                pWindow.Init h, g_ClassName(h), g_WindowText(h)
'                mWindows.Add pWindow
'
'                ' /* does it match any rules? */
'                uCompare pWindow, True
'
'            End If
'        Next i
'    End If

End Sub

Private Sub uCompare(ByRef aWindow As TWindow, ByVal Arrived As Boolean)
Dim pr As TRule

'    With mRules
'        .Rewind
'        Do While .GetNextTag(pr) = B_OK
'            Debug.Print pr.Class & " / " & aWindow.Class & " / " & (pr.Class Like aWindow.Class)
'            If (aWindow.Title Like pr.Title) And (aWindow.Class Like pr.Class) Then
'                If Arrived Then
'                    RaiseEvent WindowAppeared(pr, aWindow.Title, aWindow.Class, aWindow.Handle)
'
'                Else
'                    RaiseEvent WindowDisappeared(pr, aWindow.Title, aWindow.Class, aWindow.Handle)
'
'                End If
'            End If
'        Loop
'
'    End With

End Sub

Public Sub Go()

    ' /* load config */

Dim pcf As CConfFile3
Dim pcs As CConfSection
Dim pr As TRule

    Set pcf = New CConfFile3
    With pcf
        .SetFile Form1.GetConfigPath(True) & "windows.conf"
        .Load

        Do While .GetNextSection(pcs)
            If pcs.Name = "rule" Then
                Set pr = New TRule
                If pr.SetTo(pcs.GetValueWithDefault("guid"), pcs.GetValueWithDefault("title"), pcs.GetValueWithDefault("class")) Then _
                    mRules.Add pr

            End If

        Loop

    End With

    ' /* if there are no rules, create a default all-inclusive one */

    If mRules.CountItems = 0 Then
        Set pr = New TRule
        pr.SetTo "", "*", "*"
        mRules.Add pr
        uWriteConfig

    End If

Dim sz As String

Dim i As Long
Dim c As Long
Dim h As Long

    ' /* get current windows */

    c = g_AppWindows_Refresh(True)
    If c Then
        For i = 1 To c
            h = g_AppWindows_WindowAt(i)
            sz = g_WindowText(h)
            If sz = "" Then _
                sz = "<null>"

            sz = sz & "\n" & g_ClassName(h)
            mWindows.Add new_BTagItem(CStr(h), sz)

        Next i

    End If

    frmSettings.UpdateWindowWatchList

    ' /* start snooping */

    Set theTimer = new_BTimer(250)

End Sub

Private Sub uWriteConfig()
Dim pcs As CConfSection
Dim pcf As CConfFile3
Dim pr As TRule

    Set pcf = New CConfFile3
    pcf.SetFile Form1.GetConfigPath(True) & "windows.conf"

    With mRules
        .Rewind
        Do While .GetNextTag(pr) = B_OK
            Set pcs = New CConfSection
            With pcs
                .SetName "rule"
                .Add "guid", pr.Guid
                .Add "title", pr.Title
                .Add "class", pr.Class

            End With

            pcf.Add pcs

        Loop

    End With

    pcf.Save

End Sub

Public Function Rules() As BTagList

    Set Rules = mRules

End Function
