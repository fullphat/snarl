VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFolderWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum E_FOLDER_WATCH_FLAGS
    FWF_FOLDER_CREATE = 1
    FWF_FOLDER_DELETE = 2
    FWF_FOLDER_RENAME = 4
    FWF_FOLDER_CHANGE = 8

    FWF_FILE_CREATE = &H100
    FWF_FILE_DELETE = &H200
    FWF_FILE_RENAME = &H400
    FWF_FILE_CHANGE = &H400

    FWF_ALL_FOLDER = FWF_FOLDER_CREATE Or FWF_FOLDER_DELETE Or FWF_FOLDER_RENAME
    FWF_ALL_FILE = FWF_FILE_CREATE Or FWF_FILE_DELETE Or FWF_FILE_RENAME
    FWF_ALL = FWF_ALL_FOLDER Or FWF_ALL_FILE

End Enum

Dim mFlags As E_FOLDER_WATCH_FLAGS

Dim mHandle As Long
Dim mPath As String
Dim mPidl As Long
Dim mMsg As Long
Dim mGuid As String
Dim mRecurse As Boolean

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mMsg)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mPath

End Function

Public Function SetTo(ByVal Path As String, Optional ByVal Flags As E_FOLDER_WATCH_FLAGS = FWF_ALL, Optional ByVal Guid As String, Optional ByVal Recurse As Boolean) As Boolean

    If (Path = "") Or (Flags = 0) Then
        Debug.Print "TFolderWatch.SetTo(): bad path or flags"
        Exit Function

    End If

Dim ps As SHChangeNotifyEntry
Dim hsn As Long

    mPidl = SHSimpleIDListFromPath(StrConv(Path, vbUnicode))
    If mPidl = 0 Then
        Debug.Print "TFolderWatch.SetTo(): '" & Path & "' is invalid"
        Exit Function

    End If

    ' /* complete the single PIDLSTRUCT */

    With ps
        .pidl = mPidl
        .bWatchSubFolders = CLng(mRecurse)

    End With

Dim dw As Long

    dw = Form1.NextFreeMsg()
    mHandle = SHChangeNotifyRegister(Form1.hWnd, SHCNRF_ShellLevel, SHCNE_DISKEVENTS, dw, 1, ps)

    If mHandle <> 0 Then
        Debug.Print "TFolderWatch.SetTo(): '" & Path & "' added using 0x" & g_HexStr(dw, 4) & " flags=0x" & g_HexStr(Flags, 4)
        mFlags = Flags
        mPath = Path
        mMsg = dw
        mRecurse = Recurse

        If Guid = "" Then
            mGuid = g_CreateGUID(True)

        Else
            mGuid = Guid

        End If

'        snDoRequest "addclass?app-sig=" & App.ProductName & "&id=" & mGuid & "&name=" & g_FormattedMidStr(mPath, 48) & " (" & Me.FlagsAsString & ")"

        SetTo = True

    End If

End Function

Private Sub Class_Terminate()

    If mHandle = 0 Then _
        Exit Sub

    SHChangeNotifyDeregister mHandle
    Debug.Print "TFolderWatch: de-registered event 0x" & g_HexStr(mHandle)
    mHandle = 0

    If mPidl = 0 Then _
        Exit Sub

    CoTaskMemFree mPidl
    Debug.Print "TFolderWatch: released pidl 0x" & g_HexStr(mPidl)
    mPidl = 0

End Sub

Public Function Path() As String

    Path = mPath

End Function

Public Function Flags() As E_FOLDER_WATCH_FLAGS

    Flags = mFlags

End Function

Public Function FlagsAsString() As String
Dim sz As String

    sz = "Folders: "

    If (mFlags And FWF_FOLDER_CREATE) Then
        sz = sz & "C"

    Else
        sz = sz & "-"

    End If

    If (mFlags And FWF_FOLDER_DELETE) Then
        sz = sz & "D"

    Else
        sz = sz & "-"

    End If

    If (mFlags And FWF_FOLDER_RENAME) Then
        sz = sz & "R"

    Else
        sz = sz & "-"

    End If

    If (mFlags And FWF_FOLDER_CHANGE) Then
        sz = sz & "A"

    Else
        sz = sz & "-"

    End If



    sz = sz & "  Files: "

    If (mFlags And FWF_FILE_CREATE) Then
        sz = sz & "C"

    Else
        sz = sz & "-"

    End If

    If (mFlags And FWF_FILE_DELETE) Then
        sz = sz & "D"

    Else
        sz = sz & "-"

    End If

    If (mFlags And FWF_FILE_RENAME) Then
        sz = sz & "R"

    Else
        sz = sz & "-"

    End If

    If (mFlags And FWF_FILE_CHANGE) Then
        sz = sz & "A"

    Else
        sz = sz & "-"

    End If


    sz = sz & "  Recurse: " & IIf(mRecurse, "yes", "no")

    FlagsAsString = sz

End Function

Public Function Guid() As String

    Guid = mGuid

End Function

Public Function RecurseAsString() As String

    RecurseAsString = IIf(mRecurse, "1", "0")

End Function



