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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2
Private Const WM_POWERBROADCAST = &H218

Private Const PBT_APMSUSPEND = 4
'Private Const PBT_APMPOWERSTATUSCHANGE = 10
Private Const PBT_APMRESUMEAUTOMATIC = 18

'Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Class      GUID Device                             Description
'CDROM      4D36E965-E325-11CE-BFC1-08002BE10318    CD/DVD/Blu-ray drives
'DiskDrive  4D36E967-E325-11CE-BFC1-08002BE10318    Hard drives
'Display    4D36E968-E325-11CE-BFC1-08002BE10318    Video adapters
'FDC        4D36E969-E325-11CE-BFC1-08002BE10318    Floppy controllers
'FloppyDisk 4D36E980-E325-11CE-BFC1-08002BE10318    Floppy drives
'HDC        4D36E96A-E325-11CE-BFC1-08002BE10318    Hard drive controllers
'HIDClass   745A17A0-74D3-11D0-B6FE-00A0C90F57DA    Some USB devices
'1394       6BDD1FC1-810F-11D0-BEC7-08002BE2092F    IEEE 1394 host controller
'Image      6BDD1FC6-810F-11D0-BEC7-08002BE2092F    Cameras and scanners
'Keyboard   4D36E96B-E325-11CE-BFC1-08002BE10318    Keyboards
'Modem      4D36E96D-E325-11CE-BFC1-08002BE10318    Modems
'Mouse      4D36E96F-E325-11CE-BFC1-08002BE10318    Mice and pointing devices
'Media      4D36E96C-E325-11CE-BFC1-08002BE10318    Audio and video devices
'Net        4D36E972-E325-11CE-BFC1-08002BE10318    Network adapters
'Ports      4D36E978-E325-11CE-BFC1-08002BE10318    Serial and parallel ports
'SCSIAdapter 4D36E97B-E325-11CE-BFC1-08002BE10318   SCSI and RAID controllers
'System     4D36E97D-E325-11CE-BFC1-08002BE10318    System buses, bridges, etc.
'USB        36FC9E60-C465-11CF-8056-444553540000    USB host controllers and hubs


'Const GUID_DEVINTERFACE_USB_DEVICE = "A5DCBF10-6530-11D2-901F-00C04FB951ED"
'Const GUID_DEVINTERFACE_DISK = "53F56307-B6BF-11D0-94F2-00A0C91EFB8B"
'Const GUID_DEVINTERFACE_PARTITION = "53F5630A-B6BF-11D0-94F2-00A0C91EFB8B"
'Const GUID_DEVINTERFACE_VOLUME = "53F5630D-B6BF-11D0-94F2-00A0C91EFB8B"
'Const GUID_DEVINTERFACE_HID = "4D1E55B2-F16F-11CF-88CB-001111000030"
'Const GUID_DEVINTERFACE_MOUSE = "378DE44C-56EF-11D1-BC8C-00A0C91405DD"
'Const GUID_DEVINTERFACE_KEYBOARD = "884B96C3-56EF-11D1-BC8C-00A0C91405DD"
'Const GUID_NDIS_LAN_CLASS = "AD498944-762F-11D0-8DCB-00C04FC3358C"
'Const GUID_DEVINTERFACE_CDROM = "53F56308-B6BF-11D0-94F2-00A0C91EFB8B"
'Const GUID_KSCATEGORY_CAPTURE = "65E8773D-8F56-11D0-A3B9-00A0C9223196"
'Const GUID_KSCATEGORY_VIDEO = "6994AD05-93EF-11D0-A3CC-00A0C9223196"
'Const GUID_DEVCLASS_1394 = "6BDD1FC1-810F-11D0-BEC7-08002BE2092F"
'Const GUID_DEVINTERFACE_MODEM = "2C7089AA-2E0E-11D1-B114-00C04FC2AAE4"
'Const GUID_DEVINTERFACE_COMPORT = "86E0D1E0-8089-11D0-9CE4-08003E301F73"
'
'    ' /* made-up names */
'Const GUID_DEVCLASS_IMAGE = "6BDD1FC6-810F-11D0-BEC7-08002BE2092F"

Dim mhDevNotify As Long

Private Type T_DEVICEINSTANCEINFO
    Description As String
    FriendlyName As String
    LocationInformation As String
    Manufacturer As String
    Class As String
    DeviceClass As String           ' // HID, etc.
    DeviceClassGUID As String
    Service As String               ' // service used by the device (e.g. "kbdhid")
    Driver As String                ' // driver instance (e.g. "{4D36E96B-E325-11CE-BFC1-08002BE10318}\0002")

End Type

Private Type T_DEVICEINFO
    DeviceType As Long
    Path As String                  ' // for volume types
    Flags As Long                   ' // ditto

    InstanceInfo As T_DEVICEINSTANCEINFO

End Type

Dim mhWnd As Long
Dim mToken As Long

'Dim mLastConnected As String
'Dim mLastDisconnected As String

'Dim mPanel As BPrefsPanel
'Dim WithEvents theAddEditPanel As TAddEditPanel

'Dim mClasses As CConfFile

'Implements KPrefsPage
'Implements KPrefsPanel
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static bIgnoreDeviceChanges As Boolean
'Static bWaitForNextPowerChange As Boolean
Dim n As Long

    On Error Resume Next

    Select Case uMsg
    Case WM_CREATE
        Form1.Add "WM_CREATE"
        mhWnd = hWnd
        uStart


    Case WM_DESTROY
        Form1.Add "WM_DESTROY"
        UnregisterDeviceNotification mhDevNotify
        If mToken > 0 Then _
            snarl_unregister mToken

'        If Not (mPanel Is Nothing) Then _
            mPanel.Quit


    Case WM_CLOSE
        PostQuitMessage 0


    Case WM_DEVICECHANGE
        If Not bIgnoreDeviceChanges Then _
            uDeviceChange wParam, lParam

        ReturnValue = -1
        BWndProcSink_WndProc = True


    Case WM_POWERBROADCAST
        Select Case wParam
        Case PBT_APMSUSPEND
            Form1.List1.AddItem "PBT_APMSUSPEND: ignoring further device notifications"
            bIgnoreDeviceChanges = True
'            bWaitForNextPowerChange = False

        Case PBT_APMRESUMEAUTOMATIC
            Form1.List1.AddItem "PBT_APMRESUMEAUTOMATIC: waiting for a while..."
            SetTimer hWnd, 1, 3000, 0
'            bWaitForNextPowerChange = True

'        Case PBT_APMPOWERSTATUSCHANGE
'            If bWaitForNextPowerChange Then
'                Form1.List1.AddItem "PBT_APMPOWERSTATUSCHANGE: allowing further device notifications"
'                bWaitForNextPowerChange = False
'                bIgnoreDeviceChanges = False
'
'            End If

        Case Else
            Form1.List1.AddItem "PBT_0x" & g_HexStr(wParam, 4)
            Debug.Print "WM_POWERBROADCAST: " & g_HexStr(wParam)

        End Select

        ReturnValue = -1
        BWndProcSink_WndProc = True


    Case WM_TIMER
        If wParam = 1 Then
            Form1.List1.AddItem "WM_TIMER: allowing further device notifications"
            bIgnoreDeviceChanges = False
            KillTimer hWnd, wParam

        End If


    Case snBroadcastMsg()
        Select Case wParam
        Case SNARL_BROADCAST_LAUNCHED
            Form1.Add "Snarl launched"
            uRegisterWithSnarl

        Case SNARL_BROADCAST_QUIT
            Form1.Add "Snarl quit"
            mToken = 0      ' // no point in unregistering...
            UnregisterDeviceNotification mhDevNotify

        End Select


    Case snAppMsg()
        Select Case wParam
'        Case SNARL_APP_PREFS
'            Form1.Add "_APP_PREFS"
'            uDoPrefs

        Case SNARLAPP_DO_PREFS
            If MsgBox("Quit SnarlWare?", vbYesNo Or vbQuestion, _
                      "SnarlWare " & App.Major & "." & App.Minor & " (Build " & App.Revision & ")") = vbYes Then _
                        PostQuitMessage 0

        Case SNARLAPP_DO_ABOUT
            snDoRequest "notify?app-sig=" & App.ProductName & _
                        "&title=SnarlWare " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments & _
                        "&text=" & App.LegalCopyright & vbCrLf & "Build " & CStr(App.Revision) & _
                        "&icon=" & g_MakePath(App.Path) & "icon.png"

        End Select

    End Select

End Function

'Private Sub KPrefsPage_AllAttached()
'End Sub
'
'Private Sub KPrefsPage_Attached()
'End Sub
'
'Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
'Dim i As Long
'
'    Select Case Control.GetName()
'
'    Case "task_list"
'        mSelected = Val(Value)
'
'    Case "toolbar"
'        Select Case Val(Value)
'        Case 1
'            ' /* add new */
'            Set theAddEditPanel = New TAddEditPanel
'            theAddEditPanel.Go mPanel.hwnd, uCreateNewTask(), True
'
'
'        Case 2
'            ' /* edit selected */
'            If mSelected = 0 Then _
'                Exit Sub
'
'            Set theAddEditPanel = New TAddEditPanel
'            theAddEditPanel.Go mPanel.hwnd, mConfig.SectionAt(mSelected), False
'
'
'        Case 4
'            ' /* delete selected */
'            If mSelected = 0 Then _
'                Exit Sub
'
'            i = mConfig.IndexOf(mConfig.SectionAt(mSelected).Name)
'            If i Then
'                If mConfig.RemoveSection(i, True) Then
'                    uSetTaskList i
'                    uAddTasksToSnarl
'
'                End If
'            End If
'
'        End Select
'
'    Case Else
'
'    End Select
'
'End Sub
'
'Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
'
'    Select Case Control.GetName()
'    Case "task_list"
'        If mSelected = 0 Then _
'            Exit Sub
'
'        Set theAddEditPanel = New TAddEditPanel
'        theAddEditPanel.Go mPanel.hwnd, mConfig.SectionAt(mSelected), False
'
'    End Select
'
'End Sub
'
'Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
'End Sub
'
'Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
'End Sub
'
'Private Sub KPrefsPage_Destroy()
'End Sub
'
'Private Sub KPrefsPage_Detached()
'End Sub
'
'Private Function KPrefsPage_hWnd() As Long
'End Function
'
'Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
'End Sub

Private Sub uRegisterWithSnarl()
Dim pdbdi As DEV_BROADCAST_DEVICEINTERFACE
Dim szGuid As String
Dim szVal As String

'    mToken = sn41RegisterApp(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", mhWnd, WM_NOTIFICATION, SNARL_APP_IS_WINDOWLESS Or SNARL_APP_HAS_ABOUT)

    mToken = snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", , mhWnd, WM_NOTIFICATION, SNARLAPP_IS_WINDOWLESS Or SNARLAPP_HAS_PREFS Or SNARLAPP_HAS_ABOUT)

    If mToken > 0 Then
        Form1.List1.Tag = CStr(mToken)
        Form1.Add "Registered with Snarl (token=" & CStr(mToken) & ")"

'        uLoadClasses

'        With mClasses
'            .Rewind
'            Do While .GetEntry(szGuid, szVal)
'                snDoRequest "addclass?token=" & CStr(mToken) & "&id=" & szGuid & "&name=" & uNameFromEntry(szVal)
'
'            Loop
'
'        End With

        snDoRequest "addclass?token=" & CStr(mToken) & "&id=dcon&name=Device connected"
        snDoRequest "addclass?token=" & CStr(mToken) & "&id=ddis&name=Device disconnected"

        With pdbdi
            .dbcc_size = Len(pdbdi)
            .dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE

        End With

        mhDevNotify = RegisterDeviceNotification(mhWnd, pdbdi, DEVICE_NOTIFY_WINDOW_HANDLE Or DEVICE_NOTIFY_ALL_INTERFACE_CLASSES)

    Else
        Form1.Add "Error registering with Snarl (" & Abs(mToken) & ")"

    End If

End Sub

'Private Sub uDoPrefs()
'
'    If Not (mPanel Is Nothing) Then _
'        Exit Sub
'
'Dim aPage As BPrefsPage
'Dim aMsg As CTempMsg
'Dim pc As BControl
'
'    Set mPanel = New BPrefsPanel
'    With mPanel
'        .SetHandler Me
'        .SetTitle "SnarlTasks"
'        .SetWidth 380
'
'        Set aPage = new_BPrefsPage("", Nothing, Me)
'        With aPage
'            .SetMargin 0
'            .Add new_BPrefsControl("banner", "", "Tasks")
'
'            Set aMsg = New CTempMsg
'            aMsg.Add "item-height", 42&
'            Set pc = new_BPrefsControl("listbox", "task_list", "", "", "0", aMsg)
'            pc.SizeTo 0, 300
'            .Add pc
'
'            .Add new_BPrefsControl("fancytoolbar", "toolbar", "New...|Edit||Delete")
'
''        Set pm = New CTempMsg
''        pm.Add "text", "Use the same settings for all notification classes?"
''        pm.Add "align", 1
''        pm.Add "inset_by", 0
''        .Add new_BPrefsControl("fancytoggle", "ft>use_same", , , "0", pm)
'''    mPage.Add new_BPrefsControl("label", "", "If enabled, Snarl will apply the settings defined for the <Other> or <All> class to all notifications generated by the application.")
''
''        .Add new_BPrefsControl("banner", "", "Notification classes")
''
''        Set pm = New CTempMsg
''        pm.Add "checkboxes", 1
''        .Add new_BPrefsControl("listbox", "lb>classes", "", "", 1, pm)
''        .Add new_BPrefsControl("fancybutton2", "fb>cfg_class", "Configure...")
'
'
'
'
'        End With
'        .AddPage aPage
'
'
''        uSetTaskList
''        If mConfig.CountSections > 0 Then _
''            mSelected = 1
'
'        .Go
'
''        g_SetWindowIconToAppResourceIcon .hWnd
'
'        ' /* options:
'        '       mark item as read when replying or forwarding?
'        '       group multiple new items into a single notification?
'        '       open item if notification is clicked?
'        '
'        ' */
'
'    End With
'
'End Sub

Private Sub uStart()

    If snIsSnarlRunning() Then
        uRegisterWithSnarl

    Else
        Form1.Add "Snarl not running, waiting..."

    End If

End Sub

Private Sub uDeviceChange(ByVal wParam As Long, ByVal lParam As Long)
Dim pdbh As DEV_BROADCAST_HDR
Dim pDevice As T_DEVICEINFO
Dim szDevice As String
Dim szIcon As String

    Select Case wParam
    Case DBT_DEVNODES_CHANGED

    Case DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE

        Form1.Add IIf(wParam = DBT_DEVICEARRIVAL, "DBT_DEVICEARRIVAL", "DBT_DEVICEREMOVECOMPLETE")

        CopyMemory pdbh, ByVal lParam, Len(pdbh)
        Form1.Add "  DEV_BROADCAST_HDR->dbch_devicetype=" & CStr(pdbh.dbch_devicetype) & IIf(pdbh.dbch_devicetype = DBT_DEVTYP_DEVICEINTERFACE, " (DBT_DEVTYP_DEVICEINTERFACE)", "")

        If uGetDeviceInfo(pdbh.dbch_devicetype, lParam, pDevice) Then
            If pDevice.DeviceType = DBT_DEVTYP_DEVICEINTERFACE Then
                Form1.Add "  Identified!  Base type=" & pDevice.InstanceInfo.Class & " UID=" & pDevice.InstanceInfo.Driver

'                szIcon = g_MakePath(App.Path) & "icons\" & pDevice.InstanceInfo.Class & ".png"

                ' /* prefer "FriendlyName"... */
                szDevice = pDevice.InstanceInfo.FriendlyName

                ' /* ...then "LocationInformation" */
                If szDevice = "" Then _
                    szDevice = pDevice.InstanceInfo.LocationInformation

                ' /* ...but use Description if necessary */
                If szDevice = "" Then
                    If Not gVerboseMode Then
                        Form1.Add "  Not displaying: verbose mode not enabled"
                        Exit Sub

                    Else
                        szDevice = pDevice.InstanceInfo.Description

                    End If
                End If

                ' /* add debug info */

                If gDebugMode Then _
                    szDevice = szDevice & IIf(szDevice <> "", vbCrLf & vbCrLf, "") & _
                               "Class: " & pDevice.InstanceInfo.Class & " (" & pDevice.InstanceInfo.DeviceClassGUID & ")" & vbCrLf & _
                               "FriendlyName: " & pDevice.InstanceInfo.FriendlyName & vbCrLf & _
                               "Description: " & pDevice.InstanceInfo.Description & vbCrLf & _
                               "UID: " & pDevice.InstanceInfo.Driver

                ' /* icon */

                szIcon = g_MakePath(App.Path) & "icons\" & pDevice.InstanceInfo.Class & ".png"
                If Not g_Exists(szIcon) Then _
                    szIcon = g_MakePath(App.Path) & "icons\unknown.png"

                Debug.Print ">> " & snDoRequest("notify?app-sig=" & App.ProductName & _
                            "&id=" & IIf(wParam = DBT_DEVICEARRIVAL, "dcon", "ddis") & _
                            "&title=Device " & IIf(wParam = DBT_DEVICEARRIVAL, "connected", "disconnected") & _
                            "&text=" & szDevice & _
                            "&icon=" & szIcon & _
                            "&uid=" & pDevice.InstanceInfo.Driver)

            Else
                Form1.Add "  not a DBT_DEVTYP_DEVICEINTERFACE"

            End If

        Else
            Form1.Add " failed to retrieve info"

        End If

    Case Else
        Debug.Print "WM_DEVICECHANGE: class=0x" & g_HexStr(wParam)

    End Select

End Sub

Private Function uGetDeviceInfo(ByVal DeviceType As Long, ByVal ptr As Long, ByRef DeviceInfo As T_DEVICEINFO) As Boolean
Dim pdbdi As DEV_BROADCAST_DEVICEINTERFACER
Dim pvol As DEV_BROADCAST_VOLUME
Dim pInfo As T_DEVICEINSTANCEINFO

    ' /* pre-set as unknown */
    DeviceInfo.DeviceType = -1
    LSet DeviceInfo.InstanceInfo = pInfo

    Select Case DeviceType

    Case DBT_DEVTYP_VOLUME
        ' // logical volume
        CopyMemory pvol, ByVal ptr, Len(pvol)
        Debug.Print "  DBT_DEVTYP_VOLUME: " & g_HexStr(pvol.dbcv_unitmask) & " " & g_HexStr(pvol.dbcv_flags)
        DeviceInfo.Path = g_BitMaskToDrive(pvol.dbcv_unitmask)
        DeviceInfo.Flags = pvol.dbcv_flags

        If (pvol.dbcv_flags And DBTF_NET) Then
            ' /* network drive */
            Debug.Print "  DBTF_NET"

        Else
            ' /* physical drive */
            Debug.Print "  (physical)"

        End If

        uGetDeviceInfo = True


    Case DBT_DEVTYP_DEVICEINTERFACE
        ' /* device interface class
        '       ->dbcc_classguid is the device class guid
        '       ->decc_name is a packed registry path
        ' */

        CopyMemory pdbdi, ByVal ptr, Len(pdbdi)
'        Debug.Print "  DBT_DEVTYPE_DEVICEINTERFACE"
'        Debug.Print "  guid: " & g_StringFromGuid(pdbdi.dbcc_classguid)
        Debug.Print "  name: " & StrConv(pdbdi.dbcc_name, 0)

        Form1.Add "  DBT_DEVTYPE_DEVICEINTERFACE->dbcc_classguid=" & g_StringFromGuid(pdbdi.dbcc_classguid)
        Form1.Add "  DBT_DEVTYPE_DEVICEINTERFACE->dbcc_name=" & StrConv(pdbdi.dbcc_name, 0)

        ' /* extract information from the packed registry path */

        If uGetDeviceInstanceInfo(StrConv(pdbdi.dbcc_name, 0), pInfo) Then
            LSet DeviceInfo.InstanceInfo = pInfo
            DeviceInfo.DeviceType = DeviceType
            DeviceInfo.InstanceInfo.DeviceClassGUID = g_GuidToString(pdbdi.dbcc_classguid)

'            ' /* convert the class GUID into something human-readable */
'
'            If mClasses.Exists(DeviceInfo.InstanceInfo.DeviceClassGUID) Then
'                DeviceInfo.InstanceInfo.DeviceClass = uGetClassName(mClasses.ValueOf(DeviceInfo.InstanceInfo.DeviceClassGUID))
'
'            Else
'                Form1.Add " * class name not found (guid=" & DeviceInfo.InstanceInfo.DeviceClassGUID & ", class=" & DeviceInfo.InstanceInfo.Class & ")"
'
'            End If

            Debug.Print DeviceInfo.InstanceInfo.DeviceClass
            uGetDeviceInfo = True

        Else
            Form1.Add " unable to open registry path"

        End If

    Case Else
        Debug.Print "uDecodeDeviceType(): unknown type " & DeviceType

    End Select

End Function

Private Function uGetDeviceInstanceInfo(ByVal dbcc_name As String, ByRef Info As T_DEVICEINSTANCEINFO) As Boolean
Dim szClass As String

'    szClass = "\\?\HID#Vid_046e&Pid_5500&MI_01&Col02#7&26ec423b&0&0001#{4d1e55b2-f16f-11cf-88cb-001111000030}"

    ' /* remove \\?\ prefix */

    szClass = StrReverse(dbcc_name)
    szClass = g_SafeLeftStr(szClass, Len(szClass) - 4)

    ' /* extract class name */

Dim i As Long

    i = InStr(szClass, "#")
    If i Then _
        szClass = g_SafeRightStr(szClass, Len(szClass) - i)

    szClass = StrReverse(szClass)
    szClass = Replace$(szClass, "#", "\")

    i = InStr(szClass, "\")
    If i Then _
        Info.DeviceClass = g_SafeLeftStr(szClass, i - 1)

Dim hKey As Long

    If reg_OpenKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Enum\" & szClass, hKey) Then

        ' /* read what we can from the registry - it's unclear if any of these are mandatory */

        reg_GetValue hKey, "Mfg", Info.Manufacturer
        reg_GetValue hKey, "DeviceDesc", Info.Description
        reg_GetValue hKey, "Class", Info.Class
        reg_GetValue hKey, "LocationInformation", Info.LocationInformation  ' // not always present
        reg_GetValue hKey, "FriendlyName", Info.FriendlyName                ' // not always present
        reg_GetValue hKey, "Service", Info.Service
        reg_GetValue hKey, "Driver", Info.Driver

        Form1.Add "  Mfg=" & Info.Manufacturer
        Form1.Add "  DeviceDesc=" & Info.Description
        Form1.Add "  Class=" & Info.Class
        Form1.Add "  LocationInformation=" & Info.LocationInformation
        Form1.Add "  FriendlyName=" & Info.FriendlyName
        Form1.Add "  Service=" & Info.Service
        Form1.Add "  Driver=" & Info.Driver
 
        RegCloseKey hKey
        uGetDeviceInstanceInfo = True

    Else
        Debug.Print "Error opening HKLM\SYSTEM\CurrentControlSet\Enum\" & szClass

    End If


'    hDevInfo = SetupDiGetClassDevs(0, szClass, 0, DIGCF_ALLCLASSES)
'    If hDevInfo <> INVALID_HANDLE_VALUE Then
'
'Dim pspDevInfoData As SP_DEVINFO_DATA
'Dim buf As String
'Dim nSize As Long
'
'        i = 0
'        pspDevInfoData.cbSize = Len(pspDevInfoData)
'        Debug.Print SetupDiEnumDeviceInfo(hDevInfo, i, pspDevInfoData)
'        Debug.Print g_ApiError()
'        buf = String$(1024, 0)
'        SetupDiGetDeviceInstanceId hDevInfo, pspDevInfoData, buf, Len(buf), nSize
'
'        Debug.Print g_TrimStr(buf)
'
'        SetupDiDestroyDeviceInfoList hDevInfo
'
'    Else
'        Debug.Print "failed: SetupDiGetClassDevs()"
'
'    End If




'SP_DEVINFO_DATA* pspDevInfoData =
'        (SP_DEVINFO_DATA*)HeapAlloc(GetProcessHeap(), 0, sizeof(SP_DEVINFO_DATA));
'    pspDevInfoData->cbSize = sizeof(SP_DEVINFO_DATA);
'    for(int i=0; SetupDiEnumDeviceInfo(hDevInfo,i,pspDevInfoData); i++)
'    {
'        DWORD DataT ;
'        DWORD nSize=0 ;
'        TCHAR buf[MAX_PATH];
'
'        if ( !SetupDiGetDeviceInstanceId(hDevInfo, pspDevInfoData, buf, sizeof(buf), &nSize) ) {
'            AfxMessageBox(CString("SetupDiGetDeviceInstanceId(): ")
'                + _com_error(GetLastError()).ErrorMessage(), MB_ICONEXCLAMATION);
'            break;
'        }
'
'        if ( szDevId == buf ) {
'            // device found
'            if ( SetupDiGetDeviceRegistryProperty(hDevInfo, pspDevInfoData,
'                SPDRP_FRIENDLYNAME, &DataT, (PBYTE)buf, sizeof(buf), &nSize) ) {
'                // do nothing
'            } else if ( SetupDiGetDeviceRegistryProperty(hDevInfo, pspDevInfoData,
'                SPDRP_DEVICEDESC, &DataT, (PBYTE)buf, sizeof(buf), &nSize) ) {
'                // do nothing
'            } else {
'                lstrcpy(buf, _T("Unknown"));
'            }
'            // update UI
'            // .....
'            // .....
'            break;
'        }
'    }

End Function

'Private Function uGetCustomClassName(ByVal ClassGUID As String) As String
'
'    With New CConfFile
'        If .SetTo(g_MakePath(App.Path) & "customclasses.conf") Then _
'            uGetCustomClassName = .ValueOf(UCase$(ClassGUID))
'
'    End With
'
'End Function

'Private Sub uLoadClasses()
'
'    Set mClasses = New CConfFile
'    mClasses.SetTo g_MakePath(App.Path) & "customclasses.conf"
'
'End Sub

Private Function uNameFromEntry(ByVal Entry As String) As String
Dim s() As String

    On Error Resume Next

    s = Split(Entry, ",")
    uNameFromEntry = s(1)

End Function

Private Function uGetClassName(ByVal Entry As String) As String
Dim s() As String

    On Error Resume Next

    s = Split(Entry, ",")
    uGetClassName = s(0)

End Function

