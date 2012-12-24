Attribute VB_Name = "mSnarlExtensions_i"
Option Explicit

Public Enum SNARL_EXT_CMDS
    SNARL_EXT_INIT = &H241
    SNARL_EXT_QUIT
    SNARL_EXT_START
    SNARL_EXT_STOP
    SNARL_EXT_PREFS
    SNARL_EXT_STATUS                ' // wParam = Snarl status: 0=stopped, 1=running
       ' /* V40 onwards */
'    SNARL_EXT_GET_VERSION
    SNARL_EXT_GET_FLAGS

End Enum

Public Enum SNARL_EXT_FLAGS
    SNARL_EXT_IS_CONFIGURABLE = 1

End Enum

'Public Const SNARL_EXT_CURRENT_VERSION = &H45585402     ' // "EXT",0x02

Private Const CSIDL_APPDATA = &H1A
Private Const CSIDL_COMMONAPPDATA = &H23
Public Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWndOwner As Long, ByVal lpszPath As String, ByVal nFolder As Long, ByVal fCreate As Boolean) As Long

Public Function snExt_GetUserPath(ByRef Path As String, Optional ByVal AllUsers As Boolean) As Boolean
Dim dwFlags As Long
Dim sz As String

    dwFlags = IIf(AllUsers, CSIDL_COMMONAPPDATA, CSIDL_APPDATA)

    sz = String$(4096, 0)
    If SHGetSpecialFolderPath(0, sz, dwFlags, False) Then
        Path = uMakePath(uTrimStr(sz)) & "full phat\snarl\"
        snExt_GetUserPath = True

    End If

End Function

Private Function uTrimStr(ByVal sz As String) As String
Dim i As Long

    i = InStr(sz, Chr$(0))
    If i Then
        uTrimStr = Left$(sz, i - 1)

    Else
        uTrimStr = sz

    End If

End Function

Private Function uMakePath(ByVal Path As String) As String

    If (Path = "") Then _
        Exit Function

    If Right$(Path, 1) <> "\" Then
        uMakePath = Path & "\"

    Else
        uMakePath = Path

    End If

End Function

Public Function sx_IsValidImage(ByRef Image As MImage) As Boolean

    If (Image Is Nothing) Then _
        Exit Function

    sx_IsValidImage = ((Image.Width > 0) And (Image.Height > 0))

End Function
