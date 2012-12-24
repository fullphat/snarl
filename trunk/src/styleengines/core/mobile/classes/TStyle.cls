VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mInfo As style_info
Dim mDefaults As BPackedData
Dim mSettings As ConfigSection
Dim mConfig As ConfigFile

Dim mPage As KPrefsPage

Public Sub Init(ByVal StyleName As String, ByVal Description As String, ByVal Flags As S_STYLE_FLAGS, ByVal Schemes As String, _
                ByVal Major As Long, ByVal Minor As Long, ByRef Defaults As BPackedData, ByVal LibDate As String) ', _
                ByRef Page As KPrefsPage)

Dim ps As ConfigSection

    With mInfo
        .Copyright = App.LegalCopyright
        .Date = LibDate
        .Description = Description
        .Flags = Flags
        .Major = Major
        .Minor = Minor
        .Name = StyleName
        .Path = g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "")
        .Schemes = Schemes

'        If (IntFlags And E_METER_STYLE_LOAD_SCHEMES_FROM_FILE) Then
'
''            If (IntFlags And E_METER_STYLE_HAS_CUSTOM_SCHEME) Then _
'                .Schemes = "Custom"
'
'            Set mSchemes = New ConfigFile
'            With mSchemes
'                .File = g_MakePath(mInfo.Path) & "schemes.conf"
'                If .Load() Then
'                    .Rewind
'                    Do While .GetNextSection(ps)
'                        If mInfo.Schemes <> "" Then _
'                            mInfo.Schemes = mInfo.Schemes & "|"
'
'                        mInfo.Schemes = mInfo.Schemes & ps.GetValueWithDefault("name", "???")
'
'                    Loop
'                End If
'
'            End With
'
'        Else
''            .Schemes = Schemes
'
'        End If

        .SupportEmail = "snarl@fullphat.net"
        .URL = "www.fullphat.net"
        .IconPath = g_MakePath(.Path) & StyleName & ".png"

    End With

    Set mDefaults = New BPackedData
    Set mConfig = New ConfigFile

Dim i As Long

    With mConfig
        .File = g_MakePath(mInfo.Path) & StyleName & ".conf"
        .Load

        i = .FindSection("general")
        If i = 0 Then
            Set mSettings = .AddSectionObj("general")
            .Save

        Else
            Set mSettings = .SectionAt(i)

        End If

    End With

    Set mPage = Page

End Sub

Public Sub GetInfo(ByRef Info As style_info)

    LSet Info = mInfo

End Sub

Public Function Name() As String

    Name = mInfo.Name

End Function

Public Function GetDefault(ByVal Name As String) As String

    If Not (mDefaults Is Nothing) Then _
        GetDefault = mDefaults.ValueOf(Name)

End Function

Public Function Defaults() As BPackedData

    Set Defaults = mDefaults

End Function

Public Function ReadSetting(ByVal Name As String) As String

    If (mSettings Is Nothing) Or (mDefaults Is Nothing) Then _
        Exit Function

Dim sz As String

    If mSettings.Find(Name, sz) Then
        ReadSetting = sz

    Else
        ReadSetting = mDefaults.ValueOf(Name)

    End If

End Function

Public Sub WriteSetting(ByVal Name As String, ByVal Value As String)

    mSettings.Update Name, Value
    mConfig.Save

End Sub

Public Function Page() As KPrefsPage

    Set Page = mPage

End Function
