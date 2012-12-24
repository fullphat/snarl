Attribute VB_Name = "mMain"
Option Explicit

    ' /*
    '
    '
    ' */

Public Type T_CONFIG
    RefreshInterval As Long
    UseDefaultCallback As Boolean
    SuperSensitive As Boolean
    HeadlineLength As Long          ' // R2.0: 1=short, 2=medium, 3=long

'    UseFeedIcon As Boolean
'    FeedRefresh As Long

End Type

Public gConfig As T_CONFIG


