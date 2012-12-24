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

Public Enum E_METER_STYLE_FLAGS
    E_METER_STYLE_CAN_CHANGE_FONT = 1
    E_METER_STYLE_VARIABLE_METERS = 2
    E_METER_STYLE_LOAD_SCHEMES_FROM_FILE = 4
    E_METER_STYLE_HAS_CUSTOM_SCHEME = 8
    E_METER_STYLE_CAN_CHANGE_BKGND = &H10
    E_METER_STYLE_CAN_CHANGE_FGND = &H20
    E_METER_STYLE_CAN_CHANGE_BORDER = &H40
        ' // reserved for other colour option
    E_METER_STYLE_CAN_HIDE_SHOW_ICON = &H100
    E_METER_STYLE_CAN_SET_ICON_POS = &H200
    E_METER_STYLE_CAN_SET_FG_ALPHA = &H400
    E_METER_STYLE_HAS_SPECTRUMS = &H800
    E_METER_STYLE_NO_SHOW_PERCENT = &H1000
    E_METER_STYLE_CAN_SET_BG_ALPHA = &H2000
    E_METER_STYLE_SINGLE_PREVIEW = &H4000
    E_METER_STYLE_HAS_OWN_PAGE = &H8000&
    E_METER_STYLE_NO_LEGACY_METER = &H10000

End Enum

Dim mFlags As E_METER_STYLE_FLAGS
Dim mInfo As style_info
Dim mDefaults As BPackedData
Dim mSettings As ConfigSection
Dim mConfig As ConfigFile

Dim mSchemes As ConfigFile
Dim mMeters As String
Dim mSpectrums As String

Dim mPage As KPrefsPage

Public Sub Init(ByVal StyleName As String, ByVal Description As String, ByVal Flags As S_STYLE_FLAGS, ByVal Schemes As String, _
                ByVal Major As Long, ByVal Minor As Long, ByRef Defaults As BPackedData, ByVal LibDate As String, _
                ByVal IntFlags As E_METER_STYLE_FLAGS, ByVal MeterTypes As String, ByVal SpectrumTypes As String, ByRef Page As KPrefsPage)

Dim ps As ConfigSection

    With mInfo
        .Copyright = App.LegalCopyright
        .Date = LibDate
        .Description = Description
        .Flags = Flags
        .Major = Major
        .Minor = Minor
        .Name = StyleName
        .Path = g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & StyleName

        .Schemes = Schemes

        If (IntFlags And E_METER_STYLE_LOAD_SCHEMES_FROM_FILE) Then
            
'            If (IntFlags And E_METER_STYLE_HAS_CUSTOM_SCHEME) Then _
                .Schemes = "Custom"

            Set mSchemes = New ConfigFile
            With mSchemes
                .File = g_MakePath(mInfo.Path) & "schemes.conf"
                If .Load() Then
                    .Rewind
                    Do While .GetNextSection(ps)
                        If mInfo.Schemes <> "" Then _
                            mInfo.Schemes = mInfo.Schemes & "|"

                        mInfo.Schemes = mInfo.Schemes & ps.GetValueWithDefault("name", "???")

                    Loop
                End If

            End With

        Else
'            .Schemes = Schemes

        End If

        .SupportEmail = "snarl@fullphat.net"
        .URL = "www.fullphat.net"
        .IconPath = g_MakePath(.Path) & "icon.png"

    End With

    mFlags = IntFlags
    mMeters = MeterTypes
    mSpectrums = SpectrumTypes

    Set mDefaults = Defaults
    Set mConfig = New ConfigFile

Dim i As Long

    With mConfig
        .File = g_MakePath(mInfo.Path) & "settings"
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

Public Function Flags() As E_METER_STYLE_FLAGS

    Flags = mFlags

End Function

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

Public Sub DoPreview()
Dim sz() As String

    On Error Resume Next

    sz = Split(mInfo.Schemes, "|")

    If (mFlags And E_METER_STYLE_SINGLE_PREVIEW) = 0 Then _
        snPrivatePreviewScheme mInfo.Name, sz(0), , 20

    snPrivatePreviewScheme mInfo.Name, sz(0), , 20, 75

End Sub

Public Function GetSchemeColour(ByVal Scheme As String, ByVal Colour As String) As Long

    On Error Resume Next

    ' /* pre-set with default in case we can't find the required scheme */

    GetSchemeColour = Me.ReadSetting(Colour)

    If (mSchemes Is Nothing) Then _
        Exit Function

Dim i As Long

    i = mSchemes.FindSection(Scheme)

Debug.Print "scheme count: " & mSchemes.CountSections & " > " & i & " " & Scheme

    If i = 0 Then _
        Exit Function

Dim sz As String

    sz = mSchemes.SectionAt(i).GetValueWithDefault(Colour, "")
    If sz = "" Then _
        Exit Function

Dim S() As String

    S = Split(sz, ",")
    If UBound(S) <> 2 Then _
        Exit Function

    GetSchemeColour = rgba(Val(S(0)), Val(S(1)), Val(S(2)))

End Function

Public Function MeterTypes() As String

    MeterTypes = mMeters

End Function

Public Function MeterType(ByVal Index As Long) As String
Dim sz() As String

    sz = Split(mMeters, "|")
    If (Index > 0) And (Index <= UBound(sz) + 1) Then _
        MeterType = sz(Index - 1)

End Function

Public Function Spectrums() As String

    Spectrums = mSpectrums

End Function

Public Function GetSchemeSetting(ByVal Scheme As String, ByVal Setting As String, Optional ByVal DefaultValue As String) As String

    On Error Resume Next

    ' /* pre-set with default in case we can't find the required scheme */

    GetSchemeSetting = DefaultValue

    If (mSchemes Is Nothing) Then _
        Exit Function

Dim i As Long

    i = mSchemes.FindSection(Scheme)
    If i Then _
        GetSchemeSetting = mSchemes.SectionAt(i).GetValueWithDefault(Setting, DefaultValue)

End Function

Public Function Page() As KPrefsPage

    Set Page = mPage

End Function
