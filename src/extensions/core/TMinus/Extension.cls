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
    '/  Created:        3-Apr-2009
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

Private Const CLASS_COUNTDOWN_STARTED = "Countdown started"
Private Const CLASS_COUNTDOWN_COMPLETE = "Countdown complete"

Dim mLastErr As String
Dim mhWnd As Long

Dim mhWndOwner As Long              ' // handle of the window which launched our config (typically Snarl's Prefs panel)

Dim WithEvents mHandler As TWindow
Attribute mHandler.VB_VarHelpID = -1
Dim mNotificationId As Long

    ' /* look at all these interfaces! :) */

Implements MVersionInfo             ' // mandatory: provides static info about the extension
Implements MWndProcSink             ' // mandatory: provides a way for Snarl to talk to the extension
Implements MSimpleEnum              ' // optional: provides dynamic info about the extension

Private Sub Class_Initialize()

    l3OpenLog "%APPDATA%\" & App.Title & ".log"
    g_Debug App.Title & " running from '" & g_MakePath(App.Path) & "'"
    g_Debug ""

End Sub

Private Sub mHandler_CountdownComplete(ByVal Duration As String)

    If snIsMessageVisible(mNotificationId) Then _
        snHideMessage mNotificationId

    mNotificationId = snShowMessageEx(CLASS_COUNTDOWN_COMPLETE, _
                                      "Countdown Complete!", _
                                      Duration & IIf(Right$(Duration, 1) = "s", " have", " has") & " gone by.\nIt's now " & CStr(Now()), _
                                      -1, _
                                      g_MakePath(App.Path) & "icon.png")

End Sub

Private Sub mHandler_CountdownStarted(ByVal Duration As String, ByVal DueDate As Date)

    If snIsMessageVisible(mNotificationId) Then _
        snHideMessage mNotificationId

    mNotificationId = snShowMessageEx(CLASS_COUNTDOWN_STARTED, _
                                      "Countdown started", _
                                      "Due " & CStr(DueDate) & vbCrLf & Duration & " remaining", _
                                      -1, _
                                      g_MakePath(App.Path) & "icon.png")

End Sub

Private Function MSimpleEnum_Contains(ByVal Name As String) As Boolean
End Function

Private Property Get MSimpleEnum_Count() As Long
End Property

Private Function MSimpleEnum_FindItem(ByVal Name As String, Data As String) As Boolean

    Select Case Name

    Case "last_error"
        Data = mLastErr
        MSimpleEnum_FindItem = True

    Case "name"
        Data = App.Title
        MSimpleEnum_FindItem = True

    Case "path"
        Data = g_MakePath(App.Path)
        MSimpleEnum_FindItem = True

    Case "copyright"
        Data = App.LegalCopyright
        MSimpleEnum_FindItem = True
        
    Case "url"
        Data = "http://www.fullphat.net/"
        MSimpleEnum_FindItem = True

    End Select

End Function

Private Function MSimpleEnum_GetItem(Name As String, Data As String) As Boolean
End Function

Private Sub MSimpleEnum_Rewind()
End Sub

Private Property Get MVersionInfo_Date() As String

    MVersionInfo_Date = "18-Aug-2010"

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "A simple countdown controlled from the notification area (system tray)"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = 1

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = 1

End Property

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    ' /*
    '
    '   Here we get notifications from Snarl.  Don't think of this as a normal
    '   WndProc() - it isn't; it's just a handy interface we can use.
    '
    '   For all messages, hWnd is a handle to Snarl's main window.
    '
    '   uMsg can be one of the following:
    '       - SNARL_EXT_INIT: Do one-off initialization
    '       - SNARL_EXT_START: Called when Snarl starts running
    '       - SNARL_EXT_STOP: Called when Snarl stops running
    '       - SNARL_EXT_QUIT: Called when Snarl is unloaded
    '
    ' */

    Select Case uMsg
    Case SNARL_EXT_INIT
        ReturnValue = uInit()
        MWndProcSink_WndProc = True

    Case SNARL_EXT_QUIT
        uQuit

    Case SNARL_EXT_START
        g_Debug "SNARL_EXT_START"
'        g_Debug g_HexStr(snShowMessage("Test", g_MakePath(App.Path)))

    Case SNARL_EXT_STOP
        g_Debug "SNARL_EXT_STOP"

'    Case SNARL_EXT_PREFS

    End Select

End Function

Private Function uInit() As M_RESULT
Dim szPath As String

    mLastErr = ""
    uInit = M_FAILED

    ' /* does snarl's /etc/ path exist? */

    If Not snext_getuserpath(szPath) Then
        mLastErr = "Couldn't get Snarl user path"
        uInit = M_NOT_FOUND
        Exit Function

    End If

    If Not EZRegisterClass("w>" & LCase$(App.Title) & "_extension") Then
        ' /* no class */
        mLastErr = "Failed to create window class"
        Exit Function

    End If

    Set mHandler = New TWindow
    mhWnd = EZAddWindow("w>" & LCase$(App.Title) & "_extension", mHandler)
    If IsWindow(mhWnd) = 0 Then
        ' /* no window */
        mLastErr = "Failed to create window"
        EZUnregisterClass "w>" & LCase$(App.Title) & "_extension"
        Set mHandler = Nothing
        Exit Function

    End If

    ' /* ok */

    snRegisterConfig2 mhWnd, App.Title, 0, g_MakePath(App.Path) & "icon.png", g_MakePath(App.Path) & "icon.png"
    snRegisterAlert App.Title, CLASS_COUNTDOWN_STARTED
    snRegisterAlert App.Title, CLASS_COUNTDOWN_COMPLETE

    uInit = M_OK

End Function

Private Sub uQuit()

    snRevokeConfig mhWnd
    EZRemoveWindow mhWnd
    EZUnregisterClass "w>" & LCase$(App.Title) & "_extension"
    Set mHandler = Nothing

End Sub
