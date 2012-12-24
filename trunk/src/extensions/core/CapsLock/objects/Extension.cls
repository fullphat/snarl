VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Extension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           extension.cls
    '/
    '/  Description:    Extension main object
    '/
    '/  Created:        2-Feb-2010
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /*
    '
    '   This class provides information about the Extension itself.  It _must_ implement
    '   both MWndProcSink and MVersionInfo otherwise it won't get loaded by Snarl.  It's
    '   worth noting the the MWndProcSink interface here is not used as a 'normal' window
    '   procedure; it's simply a useful interface that Snarl can send messages to.
    '
    '   See the individual methods and properties for more detailed information.
    '
    '   New for Snarl V40 you can also implement MSimpleEnum which can be used to retrieve more
    '   information about the extension.  Only MSimpleEnum.FindItem() is currently used - all
    '   other functions should return zero or False for now - with the following queries:
    '
    '       'path' - return current extension path (usually App.Path)
    '       'copyright' - return copyright information
    '       'url' - return fully formatted URL to website (e.g. 'http://wwww.fullphat.net/')
    '
    '
    '   IMPORTANT: This is a fully functional Snarl extension - if you're planning on
    '   writing your own extension you should use the template extension in the Snarl SDK.
    '
    ' */

Private Const LIB_DATE = "15-Jun-2010"
Private Const LIB_RELEASE = "RC"
Private Const EXT_NAME = "KeyLock"
Private Const EXT_COMMENT = "Notifies when either the caps or num lock keys are pressed"

Dim mLastErr As String
Dim mhWndOwner As Long              ' // handle of the window which launched our config (typically Snarl's Prefs panel)

Dim mToken As Long
Dim mCaps As Boolean
Dim mNum As Boolean

Dim mCapsToken As Long
Dim mNumToken As Long

'Implements MVersionInfo             ' // optional: allows the SDOS version command to query the object
Implements ISnarlExtension

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Function ISnarlExtension_GetConfigWindow() As Long

End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "fullphat products"
        .Copyright = App.LegalCopyright
        .Date = LIB_DATE
        .Description = EXT_COMMENT
        .Flags = SNARL_EXTN_WANTS_PULSE
        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Name = EXT_NAME
        .Path = App.Path
        .Release = LIB_RELEASE
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With

End Sub

Private Function ISnarlExtension_Initialize() As Long

    l3OpenLog "%APPDATA%\" & EXT_NAME & ".log"
    g_Debug "pwd: " & g_MakePath(App.Path)

    If sn41GetVersion = 0 Then
        g_Debug "Needs Snarl V41 or better", LEMON_LEVEL_CRITICAL
        mLastErr = "Requires Snarl V41 or greater"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
Static b As Boolean

    b = uGetCapsState()
    If b <> mCaps Then
        ' /* state has changed */
        
        sn41Hide mCapsToken
        
        mCaps = b
        If mCaps Then _
            mCapsToken = sn41EZNotify(mToken, "cap", "Caps Lock on", "", -1, "!system-warning")

'        Else
'            mCapsToken = sn41EZNotify(mToken, "cap", "Key Lock", "Caps Lock off", -1, g_MakePath(App.Path) & "icon.png")
'
'        End If
    End If

    b = uGetNumState()
    If b <> mNum Then
        ' /* state has changed */

        sn41Hide mNumToken

        mNum = b
        If mNum Then _
            mNumToken = sn41EZNotify(mToken, "num", "Num Lock on", "", -1, "!system-warning")

'        Else
'            mNumToken = sn41EZNotify(mToken, "num", "Key Lock", "Num Lock off", -1, g_MakePath(App.Path) & "icon.png")
'
'        End If
    End If

End Sub

Private Sub ISnarlExtension_Start()

    ' /* register */
    mToken = sn41RegisterApp("application/vnd.k23-" & EXT_NAME, EXT_NAME, g_MakePath(App.Path) & "icon.png")

    If mToken = 0 Then _
        Exit Sub

    sn41AddClass mToken, "cap", "Caps Lock pressed"
    sn41AddClass mToken, "num", "Num Lock pressed"

    mCaps = uGetCapsState()                                                 '// get current caps lock state
    mNum = uGetNumState()

End Sub

Private Sub ISnarlExtension_Stop()

    sn41UnregisterApp mToken
    mToken = 0

End Sub

Private Sub ISnarlExtension_TidyUp()
End Sub

Private Function uGetCapsState() As Boolean
Dim b As Integer

    b = GetKeyState(vbKeyCapital)
    uGetCapsState = (b And 1)

End Function

Private Function uGetNumState() As Boolean
Dim b As Integer

    b = GetKeyState(vbKeyNumlock)
    uGetNumState = (b And 1)

End Function

