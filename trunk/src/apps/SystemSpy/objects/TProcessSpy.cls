VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TProcessSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim mProcs As BTagList

Dim mGeneral As CConfSection
Dim mInclude2 As CConfSection
Dim mExclude2 As CConfSection

Dim mConfig As CConfFile3

Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Public Event ProcessStarted(ByRef Process As TProcess)
Public Event ProcessStopped(ByRef Process As TProcess)

Private Sub Class_Initialize()

    Set mProcs = new_BTagList()

End Sub

Private Sub Class_Terminate()

    Set theTimer = Nothing

End Sub

Private Sub theTimer_Pulse()

    ' /* check new processes */

    uRefreshProcessList True

End Sub

'Private Sub uCompare(ByRef aWindow As TWindow, ByVal Arrived As Boolean)
'Dim pr As TRule
'
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
'
'End Sub

Public Sub Go(Optional ByVal RefreshInterval As Long = 1000)

    ' /* load config */

    Set mConfig = New CConfFile3
    With mConfig
        .SetFile Form1.GetConfigPath(True) & "process.conf"
        .Load

        ' /* [general] */

        If Not .Find("general", mGeneral) Then
            Set mGeneral = New CConfSection
            With mGeneral
                .SetName "general"
                .Add "mode", "1"

            End With

            mConfig.Add mGeneral

        End If

Dim sz As String

        With mGeneral
            sz = .GetValueWithDefault("mode", "1")
            If (sz <> "0") And (sz <> "1") Then _
                sz = "0"

        End With


        ' /* [include] */

        If Not .Find("include", mInclude2) Then
            Set mInclude2 = New CConfSection
            mInclude2.SetName "include"
            mConfig.Add mInclude2

        End If

        ' /* [exclude] */

        If Not .Find("exclude", mExclude2) Then
            Set mExclude2 = New CConfSection
            mExclude2.SetName "exclude"
            mConfig.Add mExclude2

        End If


        mConfig.Save

    End With

'    With frmSettings
'        .zUpdateExcludeProcessList
'        .zUpdateIncludeProcessList
'
'    End With

    ' /* get current baseline */

    uRefreshProcessList False

    ' /* start snooping */

    Set theTimer = new_BTimer(RefreshInterval)

End Sub

Private Sub uRefreshProcessList(ByVal NotifyChanges As Boolean)
Static c As Long

    c = g_GetProcessCount(False, False)
    If c = 0 Then _
        Exit Sub            ' // this would be serious...

Dim pt As TProcess
Static i As Long

    With mProcs
        If .CountItems Then
            For i = .CountItems To 1 Step -1
                Set pt = .TagAt(i)
                If g_FindProcess(pt.Pid) = 0 Then
                    If (NotifyChanges) And (uSnooping(pt.Name)) Then _
                        RaiseEvent ProcessStopped(pt)

                    mProcs.Remove i

                End If
            Next i
        End If
    End With

Dim pp As GMPROCESSINFO

    For i = 1 To c
        g_ProcessAt i, pp
        If mProcs.IndexOf(CStr(pp.ProcessId)) = 0 Then
            Set pt = New TProcess
            pt.bInit pp
            mProcs.Add pt
            If (NotifyChanges) And (uSnooping(pt.Name)) Then _
                RaiseEvent ProcessStarted(pt)

        End If
    Next i

End Sub

Public Function IncludeList() As CConfSection

    Set IncludeList = mInclude2

End Function

Public Function ExcludeList() As CConfSection

    Set ExcludeList = mExclude2

End Function

Public Function AddNewIncludeProcess(ByVal Name As String) As Boolean

    If Name = "" Then _
        Exit Function

    Name = LCase$(Name)

    If mInclude2.IndexOf(Name) = 0 Then
        mInclude2.Add Name
        mConfig.Save
        AddNewIncludeProcess = True

    End If

End Function

Public Function AddNewExcludeProcess(ByVal Name As String) As Boolean

    If Name = "" Then _
        Exit Function

    Name = LCase$(Name)

    If mExclude2.IndexOf(Name) = 0 Then
        mExclude2.Add Name
        mConfig.Save
        AddNewExcludeProcess = True

    End If

End Function

Public Function Mode() As Long

    If NOTNULL(mGeneral) Then _
        Mode = g_SafeLong(mGeneral.GetValueWithDefault("mode", "1"))

    If (Mode < 0) Or (Mode > 1) Then _
        Mode = 1

End Function

Public Sub SetMode(ByVal ProcessMode As Long)

    If NOTNULL(mGeneral) Then
        mGeneral.Update "mode", CStr(ProcessMode)
        mConfig.Save

    End If

End Sub

Private Function uSnooping(ByVal Name As String) As Boolean

    Name = LCase$(Name)
    If Mode = 0 Then
        ' /* inclusive */
        uSnooping = (mInclude2.IndexOf(Name) <> 0)

    Else
        ' /* exclusive */
        uSnooping = (mExclude2.IndexOf(Name) = 0)

    End If

End Function
