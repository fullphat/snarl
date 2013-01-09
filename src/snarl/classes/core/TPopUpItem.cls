VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TPopUpItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TPopUpItem.cls
    '/
    '/  Description:    Represents a selectable item in a pop-up window
    '/
    '/  Created:        19 Feb 2012
    '/
    '/  © 2012 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Guid As String
Public Label As String
Public Data As String
Dim mIcon As MImage
Public Frame As BRect
Public Token As Long

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = Me.Guid

End Function

Private Function BTagItem_Value() As String
End Function

Public Sub SetTo(ByVal Label As String, ByVal Data As String, ByRef Icon As MImage)

    Me.Guid = g_CreateGUID(True)
    Me.Label = Label
    Me.Data = Data
    Set mIcon = Icon

End Sub

Public Function Icon() As MImage

    Set Icon = mIcon

End Function
