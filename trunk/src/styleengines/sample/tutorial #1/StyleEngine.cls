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

Implements IStyleEngine

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = 1

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance

    Set IStyleEngine_CreateInstance = New StyleInstance

End Function

Private Function IStyleEngine_Date() As String
End Function

Private Function IStyleEngine_Description() As String
End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long
End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    ' /* return M_OK to indicate success or some other value to indicate an error
    '    if anything other than M_OK is returned, Snarl will abandon loading
    '    the Style Engine */

    IStyleEngine_Initialize = M_OK

End Function

Private Function IStyleEngine_LastError() As String
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

    If (Index < 1) Or (Index > 1) Then _
        Exit Sub

    With Style
'        .Copyright = App.LegalCopyright
'        .Date = LIB_DATE
'        .Description = App.FileDescription
        .Flags = S_STYLE_IS_WINDOWLESS
        .IconPath = App.Path & "\icon.png"
        .Major = 1
        .Minor = 0
        .Name = "My Style"
        .Path = App.Path & "\"
        .Schemes = "Standard"
'        .SupportEmail = "snarl@fullphat.net"
'        .URL = "www.fullphat.net"

    End With

End Sub

Private Sub IStyleEngine_TidyUp()
End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function



