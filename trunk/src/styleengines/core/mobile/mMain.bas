Attribute VB_Name = "mMain"
Option Explicit

Public gID As Long

Dim mList As BTagList

Public Sub g_AddThis(ByRef This As BTagItem)

    If (mList Is Nothing) Then _
        Set mList = new_BTagList()

    mList.Add This

End Sub

Public Sub g_RemoveThis(ByRef This As BTagItem)

    mList.Remove mList.IndexOf(This.Name)

End Sub
