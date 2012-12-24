Attribute VB_Name = "mMain"
Option Explicit

Public Const SCHEME_1 = "Normal"
Public Const SCHEME_2 = "Compact"

'Public Const SM_CXSCREEN = 0
'Public Const SM_CYSCREEN = 1
'Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Public gSettings As CConfFile
'Public gTinyDefaults As BPackedData

Public Function g_CreateMarker(ByVal Colour As Long, Optional ByVal Height As Long = 24) As MImage
Dim pp(3) As BPoint

    With New mfxView
        .SizeTo 20, Height + 2
        .EnableSmoothing True
        .SetHighColour Colour

        Set pp(0) = new_BPoint(0, 0)
        Set pp(1) = new_BPoint(0, Height - 1)
        Set pp(2) = new_BPoint(0 + 9, Height - 5 - 1)
        Set pp(3) = new_BPoint(0 + 9, 0)
        .FillShape pp(), True

        Set pp(0) = new_BPoint(0 + 9, 0)
        Set pp(1) = new_BPoint(0 + 9, Height - 5 - 1)
        Set pp(2) = new_BPoint(0 + 9 + 9, Height - 1)
        Set pp(3) = new_BPoint(0 + 9 + 9, 0)
        .FillShape pp(), True

        .EnableSmoothing False
        .StrokeLine new_BRect(0 + 9, 0, 0 + 9, Height - 5 - 1)

        Set g_CreateMarker = .ConvertToBitmap()

    End With

End Function

