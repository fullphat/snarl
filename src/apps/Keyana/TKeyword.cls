VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim mKeyword As String
Dim mGuid As String
Dim mIcon As String
Dim mDesc As String
Dim mUseRegEx As Boolean

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mGuid

End Function

Private Function BTagItem_Value() As String

End Function

Public Function SetFromExisting(ByRef Existing As CConfSection) As Boolean

    If (Existing Is Nothing) Then _
        Exit Function

    With Existing
        If (.Name = "watch") And (.IndexOf("keyword") <> 0) Then
            SetTo .GetValueWithDefault("keyword"), .GetValueWithDefault("guid"), .GetValueWithDefault("icon"), .GetValueWithDefault("description"), (.GetValueWithDefault("regex", "0") = "1")
            SetFromExisting = True

        End If

    End With

End Function

Public Sub SetTo(ByVal Keyword As String, ByVal Guid As String, Optional ByVal Icon As String, Optional ByVal Description As String, Optional ByVal UseRegEx As Boolean)

    mKeyword = Keyword
    If Guid = "" Then _
        Guid = g_CreateGUID()

    mGuid = Guid
    mDesc = Description
    mIcon = Icon
    mUseRegEx = UseRegEx

    Debug.Print "Keyword '" & Keyword & "': " & Guid

    snDoRequest "addclass?app-sig=" & App.ProductName & _
                "&id=" & mGuid & _
                "&name=" & Me.Description

End Sub

Public Sub Scan(ByVal Text As String)
Dim bMatched As Boolean

    If mUseRegEx Then
        bMatched = (Text Like mKeyword)
        Debug.Print "Scan(): regex '" & mKeyword & "' on '" & Text & "' -> " & CStr(bMatched)

    Else
        bMatched = (InStr(LCase$(Text), LCase$(mKeyword)) <> 0)

    End If

    If bMatched Then _
        Debug.Print "Scan(): " & Text & " matched keyword '" & mKeyword & "'"

    If bMatched Then _
        snDoRequest "notify?app-sig=" & App.ProductName & _
                    "&class=" & mGuid & _
                    "&title=Keyword detected" & _
                    "&text=" & mDesc & _
                    "&icon=" & IIf(mIcon = "", "!system-warning", mIcon)

End Sub

Public Function CreateSection() As CConfSection

    Set CreateSection = New CConfSection
    With CreateSection
        .SetName "watch"
        .Add "guid", mGuid
        .Add "keyword", mKeyword
        .Add "icon", mIcon
        .Add "description", mDesc
        .Add "regex", IIf(mUseRegEx, "1", "0")

    End With

End Function

Public Function Description() As String

    If mDesc = "" Then
        Description = Chr(34) & mKeyword & Chr(34) & " (" & IIf(mUseRegEx, "RegEx", "Keyword") & ")"

    Else
        Description = mDesc

    End If

End Function

Public Function Keyword() As String

    If mUseRegEx Then
        Keyword = "RegEx: " & mKeyword
    
    Else
        Keyword = "Keyword: " & mKeyword

    End If

End Function
