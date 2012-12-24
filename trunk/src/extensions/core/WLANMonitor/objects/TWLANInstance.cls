VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWLANInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum E_WLAN_STRENGTH
    E_WLAN_STRENGTH_UNAVAILABLE
    E_WLAN_STRENGTH_1_BAR
    E_WLAN_STRENGTH_2_BARS
    E_WLAN_STRENGTH_3_BARS
    E_WLAN_STRENGTH_4_BARS
    E_WLAN_STRENGTH_5_BARS

End Enum

Dim mName As String
Dim mStrength As E_WLAN_STRENGTH
Dim mNotifyToken As Long

Public Function SetTo(ByRef Instance As SWbemObject) As Boolean
Dim pProp As SWbemProperty

    mName = Instance.Properties_.Item("InstanceName").Value
    mStrength = uGetStrength(Instance)

'    sn41EZNotify mToken, "", "", CStr(mInstance(mCount).Strength * 20), , uIcon(i)
'    mNotifyToken = sn41EZNotify(AppToken, "", mName, CStr(mStrength * 20), , uIcon(mStrength))

End Function

Public Function Name() As String

    Name = mName

End Function

Public Function Strength() As E_WLAN_STRENGTH

    Strength = mStrength

End Function

Public Sub Check(ByRef Instance As SWbemObject)
Dim n As E_WLAN_STRENGTH
Dim i As Long

    n = uGetStrength(Instance)
    If n <> mStrength Then
        i = Abs(mStrength - n)
        g_Debug "TWLANInstance.Check(): '" & mName & "' strength changed from " & mStrength & " to " & n & " (delta=" & i & ")"
        mStrength = n

        ' /* is the delta change outside our tolerance setting?  if so, then display a notification */

        If i >= Val(gConfig.ValueOf("sensitivity")) Then
            If sn41IsVisible(mNotifyToken) Then
                sn41EZUpdate mNotifyToken, , CStr(mStrength * 20), , uIcon(mStrength)

            Else
                mNotifyToken = sn41EZNotify(gAppToken, "", mName, CStr(mStrength * 20), , uIcon(mStrength))

            End If

        Else
            g_Debug "TWLANInstance.Check(): delta is within tolerance"

        End If
    End If

End Sub

Private Function uGetStrength(ByRef Instance As SWbemObject) As E_WLAN_STRENGTH

    On Error Resume Next

Dim pProp As SWbemProperty
Dim dw As Long

    Err.Clear
    Set pProp = Instance.Properties_.Item("Ndis80211ReceivedSignalStrength")
    If (Err.Number <> 0) Or (pProp Is Nothing) Then _
        Exit Function

    dw = Val(pProp.Value)

    Select Case dw
    Case Is > -57
        uGetStrength = E_WLAN_STRENGTH_5_BARS

    Case Is > -68
        uGetStrength = E_WLAN_STRENGTH_4_BARS

    Case Is > -72
        uGetStrength = E_WLAN_STRENGTH_3_BARS

    Case Is > -80
        uGetStrength = E_WLAN_STRENGTH_2_BARS

    Case Is > -90
        uGetStrength = E_WLAN_STRENGTH_1_BAR

    End Select

'If intStrength > -57 Then
'    strBars = "5 Bars"
'ElseIf intStrength > -68 Then
'    strBars = "4 Bars"
'ElseIf intStrength > -72 Then
'    strBars = "3 Bars"
'ElseIf intStrength > -80 Then
'    strBars = "2 Bars"
'ElseIf intStrength > -90 Then
'    strBars = "1 Bar"
'Else
'    strBars = "Strength cannot be determined"
'End If



End Function

Private Function uIcon(ByVal Strength As E_WLAN_STRENGTH) As String

    Select Case Strength
    Case E_WLAN_STRENGTH_5_BARS
        uIcon = "!network-wifi-full"

    Case E_WLAN_STRENGTH_4_BARS
        uIcon = "!network-wifi-high"

    Case E_WLAN_STRENGTH_3_BARS
        uIcon = "!network-wifi-mid"

    Case E_WLAN_STRENGTH_2_BARS
        uIcon = "!network-wifi-low"

    Case E_WLAN_STRENGTH_1_BAR
        uIcon = "!network-wifi-none"

    Case Else
        uIcon = "!network-wifi-off"

    End Select

End Function

