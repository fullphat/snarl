VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRunFileScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRunFileScheme.cls
    '/
    '/  Description:    Represents an individual Run File scheme
    '/
    '/  Created:        1-Jun-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* versions */

    ' /* version 1 = filename only, static template
    '    version 2 = filename with variable template
    '    version 3 = unabridged content in &/= format (usable by HeySnarl for example)
    ' */

Dim mFile As CConfFile
Const MAX_SUPPORTED_VERSION = 3

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = Me.Name()

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = CStr(Version)

End Function

Public Function CreateNew() As Boolean
Dim sz As String

    sz = style_GetSnarlStylesPath()
    If sz = "" Then _
        Exit Function

    sz = g_MakePath(sz) & "runfile"
    If Not g_Exists(sz) Then _
        Exit Function

    Set mFile = New CConfFile
    With mFile
        .SetTo g_MakePath(sz) & g_CreateGUID(True) & ".runfile"
        .Add "name", "New scheme"
        .Add "version", "1"
        .Save

    End With

    CreateNew = True

End Function

Public Function SetTo(ByVal Filename As String) As Boolean

    Set mFile = New CConfFile
    If mFile.SetTo(Filename, True) Then _
        SetTo = Me.IsValid()

End Function

Public Function Version() As Long

    Version = g_SafeLong(mFile.ValueOf("version"))

End Function

Public Sub SetVersion(ByVal Version As Long)

    mFile.Update "version", CStr(Version)
    mFile.Save

End Sub

Public Function File() As String

    File = mFile.Filename

End Function

Public Function Target() As String

    Target = mFile.ValueOf("target")

End Function

Public Sub SetTarget(ByVal Target As String)

    mFile.Update "target", Target
    mFile.Save

End Sub

Public Function IsValid() As Boolean

    If (mFile Is Nothing) Then _
        Exit Function

    IsValid = (Version > 0) And (Version <= MAX_SUPPORTED_VERSION) And (Name <> "")

End Function

Public Function Template() As String

    Template = mFile.ValueOf("template")

End Function

Public Sub SetTemplate(ByVal Template As String)

    mFile.Update "template", Template
    mFile.Save

End Sub

Public Function Name() As String

    Name = mFile.ValueOf("name")

End Function

Public Sub SetName(ByVal Name As String)

    mFile.Update "name", Name
    mFile.Save

End Sub
