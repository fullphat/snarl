
Dim myApp
dim myClasses

Set myClasses = CreateObject("libsnarl.NotificationClasses")
myClasses.Add "1", "Test Class"

Set myApp = CreateObject("libsnarl.SnarlApp")
myApp.SetTo "", "test/123", "VBScript Test", "", (myClasses), "abcdef"
myApp.Register()
