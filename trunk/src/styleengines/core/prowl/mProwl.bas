Attribute VB_Name = "mProwl"
Option Explicit

Public Type T_CONFIG
    UserKey As String
    OnlyShowPriorityNotifications As Boolean
    ReplaceCRLFs As Boolean

    UseProxyServer As Boolean
    ProxyServer As String
    ProxyServerPort As Long
    ProxyUsername As String
    ProxyPassword As String
    Timeout As Long

    AppText As String
    RedactSensitive As Boolean

End Type

Public gConfig As T_CONFIG

Public Sub g_WriteConfig()

    With New ConfigFile
        .File = style_GetSnarlConfigPath("prowl")
        With .AddSectionObj("general")
            .Add "UserKey", gConfig.UserKey
            .Add "OnlyShowPriorityNotifications", IIf(gConfig.OnlyShowPriorityNotifications, "1", "0")
            .Add "ReplaceCRLFs", IIf(gConfig.ReplaceCRLFs, "1", "0")
            .Add "AppText", gConfig.AppText
            .Add "RedactSensitive", IIf(gConfig.RedactSensitive, "1", "0")

        End With

        With .AddSectionObj("network")
            .Add "UseProxyServer", IIf(gConfig.UseProxyServer, "1", "0")
            .Add "ProxyServer", gConfig.ProxyServer
            .Add "ProxyServerPort", CStr(gConfig.ProxyServerPort)
            .Add "ProxyUsername", gConfig.ProxyUsername
            .Add "ProxyPassword", gConfig.ProxyPassword
            .Add "Timeout", gConfig.Timeout

        End With

        .save

    End With

End Sub

