VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mInfo As GMPROCESSINFO
Dim mDesc As String
Dim mWindows As BTagList

Implements BTagItem

Private Function BTagItem_Name() As String

    If NOTNULL(mWindows) Then
        ' /* is an app */
        BTagItem_Name = mInfo.FullPath

    Else
        BTagItem_Name = CStr(mInfo.ProcessId)

    End If

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mInfo.ProcessName

End Function

Friend Sub bInit(ByRef Info As GMPROCESSINFO)

    LSet mInfo = Info
    mDesc = g_GetExeDescription(Info.FullPath)
'    If mDesc = "" Then _
'        mDesc = "<no description>"

'    Debug.Print "!" & mDesc & "!"

End Sub

Friend Sub bInitAsApp(ByVal Path As String, ByVal hWnd As Long)

    With mInfo
        .FullPath = Path
        .ProcessName = g_FilenameFromPath(Path)

    End With

    mDesc = g_GetExeDescription(Path)
    Set mWindows = new_BTagList()
    AddWindow hWnd

End Sub

Public Function Pid() As Long

    Pid = mInfo.ProcessId

End Function

Public Function Name() As String

    Name = mInfo.ProcessName

End Function

Public Function Description() As String

    Description = mDesc

End Function

Public Function Path() As String

    Path = mInfo.FullPath

End Function

Public Sub AddWindow(ByVal hWnd As Long)

    mWindows.Add new_BTagItem(CStr(hWnd), "")

End Sub

Public Sub CheckWindows()
Dim i As Long

    With mWindows
        If .CountItems Then
            For i = .CountItems To 1 Step -1
                If IsWindow(Val(.TagAt(i).Name)) = 0 Then _
                    .Remove i

            Next i
        End If
    End With

End Sub

Public Function WindowCount() As Long

    WindowCount = mWindows.CountItems

End Function
