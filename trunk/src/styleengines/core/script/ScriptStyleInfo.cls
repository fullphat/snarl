VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptStyleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mCustomValues As BTagList

Public Name As String               ' // signature created from this
Public Format As String             ' // "display" or "redirect"
Public Description As String
Public Icon As String
Public Version As Long
Public Revision As Long

Public Schemes As String

Public Copyright As String
Public SupportURL As String
Public Error As Long
Public Reason As String

Private Sub Class_Initialize()

    Set mCustomValues = new_BTagList()

End Sub

Public Sub SetValue(ByVal Name As String, ByVal Value As String)

    If Value = "" Then
        mCustomValues.Remove mCustomValues.IndexOf(Name, False)

    Else
        mCustomValues.Update Name, Value, False

    End If

End Sub


