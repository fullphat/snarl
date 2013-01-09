VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TScriptCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theScript As ScriptControl
Attribute theScript.VB_VarHelpID = -1

Public Function InitFrom(ByVal FileName As String, ByVal Language As String, ByRef Error As String) As Boolean

    On Error GoTo er

Dim szCode As String
Dim sz As String
Dim i As Integer

    Set theScript = New ScriptControl
    theScript.Language = Language

    i = FreeFile()
    Open FileName For Input As #i
    Do While Not EOF(i)
        Line Input #i, sz
        szCode = szCode & sz & vbCrLf

    Loop

    Close #i

    If Not uSafeAddCode(szCode) Then
        Error = "Syntax error in script code"
        Exit Function

    End If

    InitFrom = True
    Exit Function

er:
    Error = err.Description
    InitFrom = False

End Function

Private Function uSafeAddCode(ByVal Code As String) As Boolean

    On Error Resume Next

    err.Clear
    theScript.AddCode Code
    uSafeAddCode = (err.Number = 0)

End Function

Private Function uHasProc(ByVal Name As String, Optional ByRef Index As Long) As Boolean

    If (theScript Is Nothing) Then
        Debug.Print "TScriptCallback.uHasProc(): script not loaded"
        Exit Function

    End If

    If (Name = "") Or (theScript.Procedures.Count = 0) Then
        Debug.Print "TScriptCallback.uHasProc(): bad arg"
        Exit Function

    End If

    Name = LCase$(Name)

Dim n As Long

    With theScript.Procedures
        If .Count > 0 Then
            For n = 1 To .Count
                If LCase$(.Item(n).Name) = Name Then
                    Index = n
                    uHasProc = True
                    Exit Function

                End If
            Next n
        End If
    End With

    Debug.Print "TScriptCallback.uHasProc(): '" & Name & "' not found"

End Function

Private Function uHasFunc(ByVal Name As String) As Boolean
Dim n As Long

    If Not uHasProc(Name, n) Then _
        Exit Function

    uHasFunc = theScript.Procedures.Item(n).HasReturnValue

End Function

Private Function uCallProc(ByVal Name As String, ParamArray Args() As Variant) As Boolean

    On Error Resume Next

    If Not uHasProc(Name) Then
        Debug.Print "TScriptCallback.uCallProc(): '" & Name & "' not found"
        Exit Function

    End If

    err.Clear
    If IsMissing(Args) Then
        theScript.Run Name

    Else
        theScript.Run Name, Args

    End If

    If err.Number <> 0 Then _
        Debug.Print "TScriptCallback.uCallProc(): '" & Name & "' failed: " & err.Description

    uCallProc = (err.Number = 0)

End Function

'Public Function Procedures() As Procedures
'
'    If Not (theScript Is Nothing) Then _
'        Set Procedures = theScript.Procedures
'
'End Function

Public Sub DoClickedCallback(ByVal UID As String)

    uCallProc "OnClicked", UID

End Sub

Public Sub DoTimedOutCallback(ByVal UID As String)

    uCallProc "OnTimedOut", UID

End Sub

Public Sub DoActionCallback(ByVal UID As String, ByVal action As String)

    uCallProc "OnActionSelected", UID, action

End Sub
