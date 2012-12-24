VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TWindow.cls
    '/
    '/  Description:    Handles messages directed to the extension's window
    '/
    '/  Created:        31-Mar-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* Notes:
    '
    '   We need a native window here simply because VB forms don't seem to receive
    '   posted broadcast messages for some reason.
    '
    ' */

Private Const WM_CREATE = &H1
Private Const WM_TIMER = &H113
Private Const WM_POWERBROADCAST = &H218
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim mhWnd As Long

Dim mCurrentPower As B_POWER_STATUS
Dim mDoneLowPowerWarning As Boolean
Dim mLowPowerThreshold As Long

Public Event OnPowerStatusChanged(ByVal Connected As Boolean, ByVal CurrentChargePercent As Integer)
Public Event OnLowPowerWarning(ByVal CurrentChargePercent As Integer)

Dim mDoneLowRAMWarning As Boolean
Dim mLowRAMThreshold As Long
Dim mDoneNoRAMWarning As Boolean

Public Event OnPhysicalRAMLow(ByVal PercentRemaining As Integer)
Public Event OnPhysicalRAMGone()

Dim mHighCPUThreshold As Long           ' // percent
Dim mHighCPUDuration As Long            ' // seconds
Dim mCPUIsTracking As Boolean           ' // currently tracking high CPU
Dim mCPUTrackingStarted As Long         ' // tick when tracking was started
Dim mDoneHighCPUWarning As Boolean

Public Event OnCPUMonitoringStarted()
Public Event OnCPUMonitoringStopped()
Public Event OnCPUThresholdReached(ByVal LoadPercent As Integer, ByVal Duration As Integer)

Dim mLowDiskThreshold As Long           ' // percent
Dim mDiskDoneList As BTagList
Dim mWatchRem As Boolean
Dim mWatchOpt As Boolean
Public Event OnDiskThresholdReached(ByVal FreePercent As Integer, ByVal Path As String)

Dim mNumState As Boolean
Dim mCapState As Boolean
Dim mSclState As Boolean
Public Event OnLockKeyPressed(ByVal Key As B_MODIFIER_KEYS, ByVal Enabled As Boolean)

Dim mIP() As String
Dim mIPCount As Long
Public Event LocalIPAddressChanged(ByVal Details As String)
Public Event ExternalIPAddressChanged(ByVal IPAddress As String)
Dim mCurrentExternalIP As String
Dim WithEvents theExternalIPRequest As CHTTPRequest3
Attribute theExternalIPRequest.VB_VarHelpID = -1
Dim mIgnoreFirstIPChange As Boolean

Private Const TIMER_60SECS = 94
Private Const TIMER_DISK_CAPACITY = 95
Private Const TIMER_IE_DOWNLOAD = 96
Private Const TIMER_FAST = 97
Private Const TIMER_SYS_LOAD = 98
Private Const TIMER_SLOW = 99

Dim mIEWindowList As BTagList
Public Event IEDownloadCompleted()

Public Event ClipboardChanged()
Dim mBlockClipboardChange As Boolean

Implements MWndProcSink

Private Sub Class_Initialize()

'    EZRegisterClass App.ProductName
'    mhWnd = EZAddWindow(App.ProductName, Me)
'    Set mDiskDoneList = new_BTagList()
'
'    mCapState = uGetKeyLockState(vbKeyCapital)
'    mNumState = uGetKeyLockState(vbKeyNumlock)
'    mSclState = uGetKeyLockState(vbKeyScrollLock)
'    mBlockClipboardChange = True
'
'    mIgnoreFirstIPChange = True
'    uCheckExternalIP
'
'    uScanIPTable
'
'    Set mIEWindowList = new_BTagList()

End Sub

Private Sub Class_Terminate()

'    EZRemoveWindow mhWnd
'    EZUnregisterClass App.ProductName

End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim fClipboard As Boolean

    ' /* from the clipboard? */

    If g_ClipboardHandle(hWnd, uMsg, wParam, lParam, fClipboard) Then
        If fClipboard Then
            If mBlockClipboardChange Then
                mBlockClipboardChange = False

            Else
                RaiseEvent ClipboardChanged

            End If
        End If

        MWndProcSink_WndProc = True
        Exit Function

    End If

Static dMemPercent As Double
Static n As B_POWER_STATUS
Static m As Long

    Select Case uMsg

    Case WM_CREATE
        mCurrentPower = get_power_status()


    Case WM_POWERBROADCAST
        n = get_power_status()
        If n <> mCurrentPower Then
            If n = B_POWER_STATE_CONNECTED Then
                RaiseEvent OnPowerStatusChanged(True, get_battery_charge_percent())

            ElseIf n = B_POWER_STATE_DISCONNECTED Then
                RaiseEvent OnPowerStatusChanged(False, get_battery_charge_percent())

            End If

            mCurrentPower = n

        End If


'    Case WM_USER + 1
'        mLowPowerThreshold = wParam
'        mDoneLowPowerWarning = False
'        SetTimer hWnd, 1, 100, 0



    Case WM_TIMER
        Select Case wParam
        Case 1
            ' /* % battery check */
            m = get_battery_charge_percent()
            If mDoneLowPowerWarning Then
                If m > mLowPowerThreshold Then _
                    mDoneLowPowerWarning = False

            Else
                If m <= mLowPowerThreshold Then
                    RaiseEvent OnLowPowerWarning(get_battery_charge_percent())
                    mDoneLowPowerWarning = True

                End If

            End If


        Case 2
            ' /* % physical RAM check */
            
            dMemPercent = g_PercentMemInUse(True)

            If mDoneLowRAMWarning Then
                ' /* see if it's back above the warning limit */
                If dMemPercent > mLowRAMThreshold Then
                    g_Debug "[RAM Check]: has returned above safe threshold (now " & CStr(dMemPercent) & "%)"
                    mDoneLowRAMWarning = False

                End If

            Else
                ' /* is it at or below the warning limit */
                If dMemPercent <= mLowRAMThreshold Then
                    g_Debug "[RAM Check]: has dropped below safe threshold (now " & CStr(dMemPercent) & "%)"
                    RaiseEvent OnPhysicalRAMLow(dMemPercent)
                    mDoneLowRAMWarning = True

                End If

            End If

        Case TIMER_FAST
            ' /* fast timer - 100ms */
            uCheckKeys
            uScanIPTable

        Case TIMER_SYS_LOAD
            uCheckCPU

        Case TIMER_DISK_CAPACITY
            uCheckDisks

        Case TIMER_IE_DOWNLOAD
            uCheckIEWindows

        Case TIMER_SLOW
            ' /* slow timer - every 5 seconds */
            uCheckNoRAM

        Case TIMER_60SECS
            uCheckExternalIP

        End Select

    Case Else


    End Select

End Function

Public Function IsValid() As Boolean

    IsValid = (mhWnd <> 0)

End Function

Public Sub SetEnabled(ByVal Enabled As Boolean)

    If Enabled Then
        SetTimer mhWnd, TIMER_SLOW, 5000, 0
        SetTimer mhWnd, TIMER_FAST, 100, 0
        SetTimer mhWnd, TIMER_60SECS, 60000, 0

    End If

End Sub

Public Sub ConfigChanged(ByVal Item As String, ByVal Value As String)
'Dim b As Boolean
Dim n As Long

    g_Debug "TWindow.ConfigChanged(): " & Item & "=" & Value

    Select Case Item
    Case "cpu.enabled"
        If Value = "0" Then
            KillTimer mhWnd, TIMER_SYS_LOAD
            g_Debug "TWindow.ConfigChanged(): system load monitoring stopped"

        ElseIf Value = "1" Then
            SetTimer mhWnd, TIMER_SYS_LOAD, 1000, 0
            g_Debug "TWindow.ConfigChanged(): system load monitoring started"

        End If

    Case "cpu.duration"
        mHighCPUDuration = Val(Value)
        mDoneHighCPUWarning = False

    Case "cpu.high_mark"
        n = MAX(MIN(Val(Value), 100), 0)
        mHighCPUThreshold = n
        mDoneHighCPUWarning = False


    Case "disk.enabled"
        If Value = "0" Then
            KillTimer mhWnd, TIMER_DISK_CAPACITY
            g_Debug "TWindow.ConfigChanged(): disk capacity monitoring stopped"

        ElseIf Value = "1" Then
            mDiskDoneList.MakeEmpty
            SetTimer mhWnd, TIMER_DISK_CAPACITY, 5000, 0
            g_Debug "TWindow.ConfigChanged(): disk capacity monitoring started"

        End If

    Case "disk.low_mark"
        n = MAX(MIN(Val(Value), 100), 0)
        mLowDiskThreshold = n
        mDiskDoneList.MakeEmpty

    Case "disk.watch_optical"
        mWatchOpt = (Value = "1")
        
    Case "disk.watch_removable"
        mWatchRem = (Value = "1")

    Case "other.watch_clipboard"
        If Value = "1" Then
            ' /* enabled */
            g_Debug "TWindow.ConfigChanged(): clipboard snooping enabled"
            mBlockClipboardChange = True
            g_ClipboardInit mhWnd

        Else
            ' /* disabled */
            g_ClipboardUninit
            g_Debug "TWindow.ConfigChanged(): clipboard snooping disabled"

        End If

    Case "other.watch_ie"
        If Value = "0" Then
            KillTimer mhWnd, TIMER_IE_DOWNLOAD
            g_Debug "TWindow.ConfigChanged(): ie download monitoring stopped"

        ElseIf Value = "1" Then
            SetTimer mhWnd, TIMER_IE_DOWNLOAD, 500, 0
            g_Debug "TWindow.ConfigChanged(): ie download monitoring started"

        End If



    End Select

'    b = (Config.GetValueWithDefault("cpu.enabled", "") = "1")
'
'    If b <> mWatchingLoad Then
'        g_Debug "ConfigChanged(): cpu.enabled is now " & CStr(b)
'
'        If b Then
'            ' /* start watching load */
'            mHighCPUDuration = Val(Config.GetValueWithDefault("cpu.duration"))
'
'            n = Val(Config.GetValueWithDefault("cpu.high_mark"))
'            n = MAX(MIN(n, 100), 0)
'            mHighCPUThreshold = n
'
'            mDoneHighCPUWarning = False
'
'        Else
'            ' /* stop watching load */
''            mHighCPUThreshold = -1
'
'        End If
'
'        mWatchingLoad = b
'
'    End If

End Sub

Public Sub SetLowPowerThreshold(ByVal Value As Long)

    g_Debug "TWindow.SetLowPowerThreshold(): " & CStr(Value), LEMON_LEVEL_PROC

    KillTimer mhWnd, 1

    If Value > 100 Then
        Value = 100

    ElseIf Value < 0 Then
        Value = 0

    End If

    mLowPowerThreshold = Value
    mDoneLowPowerWarning = False

    If Value > 0 Then _
        SetTimer mhWnd, 1, 1000, 0

End Sub

Public Sub SetLowRAMThreshold(ByVal Value As Long)

    g_Debug "TWindow.SetLowRAMThreshold(): " & CStr(Value), LEMON_LEVEL_PROC

    KillTimer mhWnd, 2

    If Value > 100 Then
        Value = 100

    ElseIf Value < 0 Then
        Value = 0

    End If

    mLowRAMThreshold = Value
    mDoneLowRAMWarning = False

    If Value > 0 Then _
        SetTimer mhWnd, 2, 10000, 0

End Sub

'Public Sub SetLowDiskThreshold(ByVal Value As Long)
'
'    g_Debug "TWindow.SetLowDiskThreshold(): " & CStr(Value), LEMON_LEVEL_PROC
'
'    If Value > 100 Then
'        Value = 100
'
'    ElseIf Value < 0 Then
'        Value = 0
'
'    End If
'
'    mLowDiskThreshold = Value
'    mDiskDoneList.MakeEmpty
'
'End Sub

'Public Sub SetWatchRemovable(ByVal Enabled As Boolean)
'
'    mWatchRem = Enabled
'
'End Sub

'Public Sub SetWatchOptical(ByVal Enabled As Boolean)
'
'    mWatchOpt = Enabled
'
'End Sub

'Public Sub SetHighCPUThreshold(ByVal Value As Long)
'
'    g_Debug "TWindow.SetHighCPUThreshold(): " & CStr(Value), LEMON_LEVEL_PROC
'
'    If Value > 100 Then
'        Value = 100
'
'    ElseIf Value < 0 Then
'        Value = 0
'
'    End If
'
'    mHighCPUThreshold = Value
'    mDoneHighCPUWarning = False
'
'End Sub

'Public Sub SetHighCPUDuration(ByVal Value As Long)
'
'    g_Debug "TWindow.SetHighCPUDuration(): " & CStr(Value), LEMON_LEVEL_PROC
'
''    If Value > 100 Then
''        Value = 100
''
''    ElseIf Value < 0 Then
''        Value = 0
''
''    End If
'
'    mHighCPUDuration = Value
'    mDoneHighCPUWarning = False
'
'End Sub

Private Sub uCheckNoRAM()

    If mDoneNoRAMWarning Then
        ' /* see if it's back in the black... */
        If g_PercentMemInUse(True) > 0 Then _
            mDoneNoRAMWarning = False

    Else
        ' /* is it at or below zero */
        If g_PercentMemInUse() <= 0 Then
            RaiseEvent OnPhysicalRAMGone
            mDoneNoRAMWarning = True

        End If

    End If

End Sub

Private Sub uCheckCPU()
Dim dLoad As Double

    dLoad = krnl_GetCurrentTotalCPUPercent()

    If dLoad > mHighCPUThreshold Then
        ' /* over the threshold */

        g_Debug "uCheckCPU(): current load (" & CStr(dLoad) & "%) is over threshold (" & CStr(mHighCPUThreshold) & "%)"

        If Not mCPUIsTracking Then
            ' /* first time (not tracking) */
            mCPUTrackingStarted = GetTickCount()
            mCPUIsTracking = True
            g_Debug "uCheckCPU(): started tracking..."
            RaiseEvent OnCPUMonitoringStarted

        ElseIf Not mDoneHighCPUWarning Then
            ' /* already tracking... */
            If (GetTickCount() - mCPUTrackingStarted) > (mHighCPUDuration * 1000) Then
                ' /* ...and past the cut-off duration */
                g_Debug "uCheckCPU(): tracking duration exceeded"
                RaiseEvent OnCPUThresholdReached(mHighCPUThreshold, mHighCPUDuration)
                mDoneHighCPUWarning = True

            End If

        End If

    ElseIf mCPUIsTracking Then
        ' /* load has dropped below the threshold and we were monitoring, so stop */
        g_Debug "uCheckCPU(): stopped tracking..."
        RaiseEvent OnCPUMonitoringStopped
        mCPUIsTracking = False
        mDoneHighCPUWarning = False

    End If

End Sub

Private Sub uCheckDisks()
Dim szPath As String
Dim i As Integer

    For i = Asc("C") To Asc("Z")

        szPath = Chr$(i) & ":\"

        Select Case GetDriveType(szPath)
        Case DRIVE_FIXED
            uCheckDisk szPath

        Case DRIVE_REMOVABLE
            If mWatchRem Then _
                uCheckDisk szPath

        Case DRIVE_CDROM
            If mWatchOpt Then _
                uCheckDisk szPath
        
        End Select

    Next i

End Sub

Private Sub uCheckDisk(ByVal Path As String)
Dim n As Integer

    n = g_GetDiskFreePercent(Path)
    Debug.Print Path & " = " & CStr(n)
    If n < 0 Then _
        Exit Sub

    If (n <= mLowDiskThreshold) And (mDiskDoneList.IndexOf(Path) = 0) Then
        ' /* on or below threshold and not already notified - notify and add to list */
        mDiskDoneList.Add new_BTagItem(Path, "")
        RaiseEvent OnDiskThresholdReached(n, Path)

    ElseIf (mDiskDoneList.IndexOf(Path) <> 0) And (n > mLowDiskThreshold) Then
        ' /* notified but now above threshold - remove from list */
        Debug.Print "sysinfo-disk_monitor: " & Path & " is now above threshold"
        mDiskDoneList.Remove mDiskDoneList.IndexOf(Path)

    End If

End Sub

Private Sub uCheckKeys()
Static b As Boolean

    b = uGetKeyLockState(vbKeyCapital)
    If b <> mCapState Then
        mCapState = b
        g_Debug "uCheckKeys(): B_CAPS_LOCK changed to " & CStr(b)
        RaiseEvent OnLockKeyPressed(B_CAPS_LOCK, b)

    End If

    b = uGetKeyLockState(vbKeyNumlock)
    If b <> mNumState Then
        mNumState = b
        g_Debug "uCheckKeys(): B_NUM_LOCK changed to " & CStr(b)
        RaiseEvent OnLockKeyPressed(B_NUM_LOCK, b)

    End If

    b = uGetKeyLockState(vbKeyScrollLock)
    If b <> mSclState Then
        mSclState = b
        g_Debug "uCheckKeys(): B_SCROLL_LOCK changed to " & CStr(b)
        RaiseEvent OnLockKeyPressed(B_SCROLL_LOCK, b)

    End If

End Sub

Private Function uGetKeyLockState(ByVal KeyCode As Long) As Boolean
Dim b As Integer

    b = GetKeyState(KeyCode)
    uGetKeyLockState = (b And 1)

End Function

Private Sub uScanIPTable()
Dim fWasRemoved As Boolean
Dim fWasAdded As Boolean
Dim szText As String
Dim ip() As String
Dim sz As String
Dim c As Long
Dim i As Long

    sz = get_ip_address_table()
    ip() = Split(sz, " ")
    c = UBound(ip())

'    Debug.Print sz

    ' /* tidy the list */

    For i = 0 To c
        If (ip(i) = "0.0.0.0") Or (ip(i) = "127.0.0.1") Then _
            ip(i) = ""

    Next i

    szText = ""

    ' /* look for missing addresses */

    If mIPCount Then
        For i = mIPCount To 1 Step -1
            If InStr(sz, mIP(i)) = 0 Then
                g_Debug "uScanIPTable(): " & mIP(i) & " removed"
                szText = szText & mIP(i) & " was removed" & vbCrLf
                fWasRemoved = True
                uRemoveIPAddr i

            End If
        Next i
    End If

    ' /* look for additions */

    For i = 0 To c
        If ip(i) <> "" Then
            If uFindIPAddr(ip(i)) = 0 Then
                g_Debug "uScanIPTable(): " & ip(i) & " added"
                szText = szText & ip(i) & " was added" & vbCrLf
                uAddIPAddr ip(i)
                fWasAdded = True

            End If
        End If

    Next i

    If (fWasRemoved) Or (fWasAdded) Then _
        RaiseEvent LocalIPAddressChanged(szText)

'"Current configured addresses are:" & vbCrLf & Replace$(sz, " ", vbCrLf)
'            snShowMessageEx Me.CLASS_1, "IP Address Change", szText, 10, g_MakePath(App.Path) & "icon.png", Me.hWnd, &H401

End Sub

Private Function uFindIPAddr(ByVal ipaddr As String) As Long

    If mIPCount = 0 Then _
        Exit Function

Dim i As Long

    For i = 1 To mIPCount
        If mIP(i) = ipaddr Then
            uFindIPAddr = i
            Exit Function

        End If
    Next i

End Function

Private Sub uRemoveIPAddr(ByVal Index As Long)

    mIP(Index) = ""
    g_Debug "uRemoveIPAddr(): removed #" & CStr(Index)

End Sub

Private Sub uAddIPAddr(ByVal ipaddr As String)

    mIPCount = mIPCount + 1
    ReDim Preserve mIP(mIPCount)
    mIP(mIPCount) = ipaddr
    g_Debug "uAddIPAddr(): added '" & ipaddr & "'"

End Sub

Private Sub uCheckIEWindows()
Static pt As BTagItem
Static h As Long
Static c As Long
Static i As Long
Static p As Long

    ' /* gone windows */

    With mIEWindowList
        If .CountItems Then
            For i = .CountItems To 1 Step -1
                h = Val(.TagAt(i).Name)
                If IsWindow(h) = 0 Then
                    Debug.Print "removed " & g_HexStr(h)
                    .Remove i
            
                End If
            Next i
        End If
    End With

    ' /* new windows */

    c = g_AppWindows_Refresh(True)
    If c Then
        For i = 1 To c
            h = g_AppWindows_WindowAt(i)
            If (g_ClassName(h) = "#32770") And (g_WindowText(h) = "Download complete") And (mIEWindowList.IndexOf(CStr(h)) = 0) Then
                GetWindowThreadProcessId h, p
                If g_GetProcessFriendlyName2(p) = "Internet Explorer" Then
                    Debug.Print "found window 0x" & g_HexStr(h)
                    mIEWindowList.Add new_BTagItem(CStr(h), "")
                    RaiseEvent IEDownloadCompleted

                End If
            End If
        Next i
    End If

End Sub

Private Sub theExternalIPRequest_Completed(ByVal WasSuccessful As Boolean)
Dim bNotify As Boolean
Dim sz As String

    g_Debug "theExternalIPRequest.Completed(): success=" & CStr(WasSuccessful)

    If WasSuccessful Then
        sz = theExternalIPRequest.ResponseBody()
        g_Debug "theExternalIPRequest.Completed(): address is " & sz

        If mCurrentExternalIP <> sz Then
            mCurrentExternalIP = sz
            If mIgnoreFirstIPChange Then
                mIgnoreFirstIPChange = False

            Else
                g_Debug "theExternalIPRequest.Completed(): notifying change"
                RaiseEvent ExternalIPAddressChanged(sz)

            End If
        End If
    End If

End Sub

Private Sub uCheckExternalIP()

    g_Debug "uCheckExternalIP(): checking external IP address..."

    Set theExternalIPRequest = New CHTTPRequest3
    theExternalIPRequest.GetContent "http://automation.whatismyip.com/n09230945.asp"

End Sub
