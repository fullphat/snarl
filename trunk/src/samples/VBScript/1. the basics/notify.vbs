
Dim myApp

Set myApp = CreateObject("libsnarl.SnarlApp")
myApp.SetTo "", "test/123", "VBScript Test", "", (Nothing), "abcdef"
myApp.Notify "1", "VBScript Test", "Hello, World!", "!system-warning"
