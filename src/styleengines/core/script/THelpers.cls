VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "THelpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mView As mfxView

Friend Sub bSet(ByRef aView As mfxView)

    Set mView = aView

End Sub

Public Function MAX(ByVal a As Long, ByVal b As Long) As Long

    MAX = mBasic_i.MAX(a, b)

End Function

Public Function MIN(ByVal a As Long, ByVal b As Long) As Long

    MIN = mBasic_i.MIN(a, b)

End Function

Public Function string_as_long(ByVal str As String) As Long

    string_as_long = g_SafeLong(str)

End Function
