VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFolderSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mFolders As BTagList

Public Event FileRenamed(ByVal Was As String, ByVal Now As String)
Public Event FileCreated(ByVal Path As String)
Public Event FileDeleted(ByVal Path As String)
Public Event FolderRenamed(ByVal Was As String, ByVal Now As String)
Public Event FolderCreated(ByVal Path As String)
Public Event FolderDeleted(ByVal Path As String)

Public Sub Go()

    Set mFolders = new_BTagList()

    ' /* load config */

Dim pcf As CConfFile3
Dim pcs As CConfSection
Dim pfw As TFolderWatch

    Set pcf = New CConfFile3
    With pcf
        .SetFile Form1.GetConfigPath(True) & "folders.conf"
        .Load
        Do While .GetNextSection(pcs)
            If pcs.Name = "watch" Then
                Set pfw = New TFolderWatch
                If pfw.SetTo(pcs.GetValueWithDefault("path"), Val(pcs.GetValueWithDefault("flags")), pcs.GetValueWithDefault("guid"), pcs.GetValueWithDefault("recurse") = "1") Then _
                    mFolders.Add pfw

            End If

        Loop

    End With

    If mFolders.CountItems = 0 Then
        Set pfw = New TFolderWatch
        pfw.SetTo "c:\"
        mFolders.Add pfw

    End If

End Sub

Public Sub ShellChangeNotify(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Dim pWatch As TFolderWatch

    If Not mFolders.Find(CStr(uMsg), pWatch) Then
        Debug.Print "error: watch " & CStr(uMsg) & " not found"
        Exit Sub

    End If

Static psns As SHNOTIFYSTRUCT

    Debug.Print "foldersnoop: " & g_HexStr(wParam) & " " & g_HexStr(lParam) & " [" & g_SHNotifyStr(lParam) & "]"
    CopyMemory psns, ByVal wParam, Len(psns)
    uShellChangeEvent lParam, g_GetPathFromPIDL(psns.dwItem1), g_GetPathFromPIDL(psns.dwItem2), pWatch


End Sub

Private Sub uShellChangeEvent(ByVal EventId As Long, ByVal Path1 As String, ByVal Path2 As String, ByRef Watch As TFolderWatch)

    Debug.Print " 1> " & Path1
    Debug.Print " 2> " & Path2

    Select Case EventId

'    Case SHCNE_MEDIAINSERTED, SHCNE_MEDIAREMOVED, SHCNE_DRIVEREMOVED, SHCNE_DRIVEADD, SHCNE_NETSHARE, SHCNE_NETUNSHARE
        ' /* not interested in these */

'Public Const SHCNE_UPDATEIMAGE = &H8000&     '(G) An image in the system image list has changed.
'Public Const SHCNE_DRIVEADDGUI = &H10000     '(G) A drive has been added and the shell should create a new window for the drive.
'Public Const SHCNE_FREESPACE = &H40000       '(G) The amount of free space on a drive has changed.
'Public Const SHCNE_EXTENDED_EVENT = &H4000000 '(G) Not currently used.
'Public Const SHCNE_ASSOCCHANGED = &H8000000   '(G) A file type association has changed.

'        Exit Sub

    Case SHCNE_RENAMEITEM
        If (Watch.Flags And FWF_FILE_RENAME) Then _
            RaiseEvent FileRenamed(Path1, Path2)

    Case SHCNE_CREATE
        If (Watch.Flags And FWF_FILE_CREATE) Then _
            RaiseEvent FileCreated(Path1)

    Case SHCNE_DELETE
        If (Watch.Flags And FWF_FILE_DELETE) Then _
            RaiseEvent FileDeleted(Path1)

    Case SHCNE_RENAMEFOLDER
        If (Watch.Flags And FWF_FOLDER_RENAME) Then _
            RaiseEvent FolderRenamed(Path1, Path2)

    Case SHCNE_MKDIR
        If (Watch.Flags And FWF_FOLDER_CREATE) Then _
            RaiseEvent FolderCreated(Path1)

    Case SHCNE_RMDIR
        If (Watch.Flags And FWF_FOLDER_DELETE) Then _
            RaiseEvent FolderDeleted(Path1)


''Public Const SHCNE_ATTRIBUTES = &H800        '(D) The attributes of an item or folder have changed.
''Public Const SHCNE_UPDATEDIR = &H1000        '(D) The contents of an existing folder have changed,
''                                '    but the folder still exists and has not been renamed.
'
'
'    Case SHCNE_UPDATEITEM
'        szTitle = "Item changed"
'        szText = "Attributes for " & Path1 & " were changed"
'        szUid = Path1
'
'        If g_IsFolder(Path1) Then
'            If (Watch.Flags And FWF_FOLDER_CHANGE) = 0 Then _
'                Exit Sub
'
'            szIcon = "folder"
'            szClass = FOLDER_UPDATED
'
'        Else
'            If (Watch.Flags And FWF_FILE_CHANGE) = 0 Then _
'                Exit Sub
'
'            szIcon = "file"
'            szClass = FILE_UPDATED
'
'        End If

    Case Else
        Debug.Print "not implemented"
        Exit Sub

    End Select

End Sub

Private Sub uWriteConfig()
Dim pfw As TFolderWatch
Dim pcf As CConfFile3
Dim pcs As CConfSection

    Set pcf = New CConfFile3
    pcf.SetFile Form1.GetConfigPath(True) & "folders.conf"

    With mFolders
        .Rewind
        Do While .GetNextTag(pfw) = B_OK
            Set pcs = New CConfSection
            With pcs
                .SetName "watch"
                .Add "guid", pfw.Guid
                .Add "path", pfw.Path
                .Add "flags", pfw.Flags
                .Add "recurse", pfw.RecurseAsString

            End With

            pcf.Add pcs

        Loop

    End With

    pcf.Save

End Sub

Public Sub Remove(ByVal Index As Long)

    If (Index < 0) Or (Index > mFolders.CountItems) Then _
        Exit Sub

    mFolders.Remove Index
    uWriteConfig

End Sub

Public Sub Add(ByVal Path As String)
Dim pw As TFolderWatch

    Set pw = New TFolderWatch
    pw.SetTo Path
    mFolders.Add pw
    uWriteConfig

End Sub

Public Function List() As BTagList

    Set List = mFolders

End Function
