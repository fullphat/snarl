
'LinkConsole.vbs
'
'This is a WSH script used to make it easier to edit
'a compiled VB6 EXE using LINK.EXE to create a console
'mode program.
'
'Drag the EXE's icon onto the icon for this file, or
'execute it from a command prompt as in:
'
' &nbsp; &nbsp; &nbsp; &nbsp;LinkConsole.vbs <EXEpath&file>
'
'Be sure to set up strLINK to match your VB6 installation.

Dim strLINK, strEXE, WSHShell

strLINK = Chr(34) & "C:\Program Files\Microsoft Visual Studio\VB98\LINK.EXE" & Chr(34)
strEXE = Chr(34) & WScript.Arguments(0) & Chr(34)

Set WSHShell = CreateObject("WScript.Shell")

WSHShell.Run strLINK & " /EDIT /SUBSYSTEM:CONSOLE " & strEXE

Set WSHShell = Nothing
WScript.Echo "Complete!"