VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mStyle As TScriptStyle          ' // the scripted style which does the work - shared across instances
Dim mContent As mfxBitmap           ' // the latest content for this instance

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mContent

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    g_Debug "StyleInstance.UpdateContent()", LEMON_LEVEL_PROC_ENTER

    Set mContent = mStyle.UpdateContent(NotificationInfo, NotificationInfo.Icon)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Sub SetTo(ByRef Style As TScriptStyle)

    Set mStyle = Style

End Sub
