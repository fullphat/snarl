VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mx As Long
Dim my As Long

Implements MPoint

Private Property Get MPoint_x() As Long

    MPoint_x = mx
    
End Property

Private Property Get MPoint_y() As Long

    MPoint_y = my

End Property

Public Sub SetFrom(ByRef Point As MPoint)

    mx = Point.x
    my = Point.y

End Sub

Public Sub SetTo(ByVal x As Long, ByVal y As Long)

    mx = x
    my = y

End Sub

Public Sub OffsetBy(ByVal x As Long)

    mx = mx + x

End Sub
