VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum NOTIFICATION_PRIORITY
    LOW_PRIORITY = -1
    NORMAL_PRIORITY = 0
    HIGH_PRIORITY = 1

End Enum

Private Type T_INFO
    Class As String
    Title As String
    Text As String
    Icon As String
    Callback As String
    Duration As Long
    Priority As NOTIFICATION_PRIORITY
    UID As String
    Actions As Actions
    CallbackScript As String
    CallbackScriptType As String
    ReplaceUID As String
    MergeUID As String

End Type

Dim mInfo As T_INFO

Private Type T_ARG
    Name As String
    Value As String
    
End Type

Dim mArg() As T_ARG
Dim mArgs As Long

Private Sub Class_Initialize()

    mInfo.Duration = -1
    ReDim mArg(0)

End Sub

Friend Function bAsRequest(ByVal AppSignature As String, ByVal Password As String) As String

    bAsRequest = "notify?app-sig=" & AppSignature & "&password=" & Password & _
                 "&class=" & mInfo.Class & "&uid=" & mInfo.UID

    If mInfo.Title <> "" Then _
        bAsRequest = bAsRequest & "&title=" & uFixUp(mInfo.Title)

    If mInfo.Text <> "" Then _
        bAsRequest = bAsRequest & "&text=" & uFixUp(mInfo.Text)

    If mInfo.Priority <> NORMAL_PRIORITY Then _
        bAsRequest = bAsRequest & "&priority=" & CStr(mInfo.Priority)
        
    If mInfo.Duration <> -1 Then _
        bAsRequest = bAsRequest & "&timeout=" & CStr(mInfo.Duration)

    If mInfo.Icon <> "" Then _
        bAsRequest = bAsRequest & "&icon=" & mInfo.Icon

    If mInfo.Callback <> "" Then _
        bAsRequest = bAsRequest & "&callback=" & mInfo.Callback

    If mInfo.ReplaceUID <> "" Then
        bAsRequest = bAsRequest & "&replace-uid=" & mInfo.ReplaceUID

    ElseIf mInfo.MergeUID <> "" Then
        bAsRequest = bAsRequest & "&merge-uid=" & mInfo.MergeUID

    End If

    If (mInfo.CallbackScript <> "") And (mInfo.CallbackScriptType <> "") Then _
        bAsRequest = bAsRequest & "&callback-script=" & mInfo.CallbackScript & "&callback-script-type=" & mInfo.CallbackScriptType

    ' /* add custom arguments */

Dim i As Long

    If mArgs > 0 Then
        For i = 1 To mArgs
            bAsRequest = bAsRequest & "&" & mArg(i).Name & "=" & mArg(i).Value

        Next i
    End If

    ' /* add the actions */

    If Not (mInfo.Actions Is Nothing) Then _
        bAsRequest = bAsRequest & mInfo.Actions.bAsWin32Request

    Debug.Print bAsRequest

End Function

Public Property Get Class() As String

    Class = mInfo.Class

End Property

Public Property Let Class(ByVal vNewValue As String)

    mInfo.Class = vNewValue

End Property

Public Property Get UID() As String

    UID = mInfo.UID

End Property

Public Property Let UID(ByVal vNewValue As String)

    mInfo.UID = vNewValue

End Property

Public Property Get Title() As String

    Title = mInfo.Title

End Property

Public Property Let Title(ByVal vNewValue As String)

    mInfo.Title = vNewValue

End Property

Public Property Get Text() As String

    Text = mInfo.Text

End Property

Public Property Let Text(ByVal vNewValue As String)

    mInfo.Text = vNewValue

End Property

Public Property Get Icon() As String

    Icon = mInfo.Icon

End Property

Public Property Let Icon(ByVal vNewValue As String)

    mInfo.Icon = vNewValue

End Property

Public Property Get DefaultCallback() As String

    DefaultCallback = mInfo.Callback

End Property

Public Property Let DefaultCallback(ByVal vNewValue As String)

    mInfo.Callback = vNewValue

End Property

Public Property Get Priority() As NOTIFICATION_PRIORITY

    Priority = mInfo.Priority

End Property

Public Property Let Priority(ByVal vNewValue As NOTIFICATION_PRIORITY)

    mInfo.Priority = vNewValue

End Property

Public Property Get Duration() As Long

    Duration = mInfo.Duration

End Property

Public Property Let Duration(ByVal vNewValue As Long)

    mInfo.Duration = vNewValue

End Property

Public Property Get Actions() As Actions

    Set Actions = mInfo.Actions

End Property

Public Property Let Actions(ByVal vNewValue As Actions)

    Set mInfo.Actions = vNewValue

End Property

Public Property Get CallbackScript() As String

    DefaultCallback = mInfo.CallbackScript

End Property

Public Property Let CallbackScript(ByVal vNewValue As String)

    mInfo.CallbackScript = vNewValue

End Property

Public Property Get CallbackScriptType() As String

    CallbackScriptType = mInfo.CallbackScriptType

End Property

Public Property Let CallbackScriptType(ByVal vNewValue As String)

    mInfo.CallbackScriptType = vNewValue

End Property

Public Property Get ReplaceUID() As String

    ReplaceUID = mInfo.ReplaceUID

End Property

Public Property Let ReplaceUID(ByVal vNewValue As String)

    mInfo.ReplaceUID = vNewValue

End Property

Public Property Get MergeUID() As String

    MergeUID = mInfo.MergeUID

End Property

Public Property Let MergeUID(ByVal vNewValue As String)

    mInfo.MergeUID = vNewValue

End Property

Public Sub Add(ByVal Name As String, ByVal Value As String, Optional ByVal UpdateIfExists As Boolean = False)
Dim i As Long

    i = uFindArg(Name)
    If (i) And (UpdateIfExists) Then
        mArg(i).Value = Value

    Else
        mArgs = mArgs + 1
        ReDim Preserve mArg(mArgs)
        mArg(mArgs).Name = Name
        mArg(mArgs).Value = Value

    End If

End Sub

Public Function Remove(ByVal Name As String) As Boolean
Dim i As Long
Dim j As Long

    i = uFindArg(Name)
    If i = 0 Then _
        Exit Function

    If i < mArgs Then
        For j = i To (mArgs - 1)
            LSet mArg(j) = mArg(j + 1)

        Next j
    End If

    mArgs = mArgs - 1
    ReDim Preserve mArg(mArgs)
    Remove = True

End Function

Private Function uFindArg(ByVal Name As String) As Long
Static i As Long

    If mArgs Then
        For i = 1 To mArgs
            If mArg(i).Name = Name Then
                uFindArg = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Function uFixUp(ByVal str As String) As String

    str = Replace$(str, "&", "&&")
    str = Replace$(str, "=", "==")

    uFixUp = str

End Function
