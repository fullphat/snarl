Attribute VB_Name = "mMain"
Option Explicit

Public Sub Main()

    If snDoRequest("register?app-sig=" & App.ProductName & "&title=" & App.Title) > 0 Then _
        snDoRequest "notify?app-sig=" & App.ProductName & _
                    "&title=Simplest App Ever" & _
                    "&text=Hello, world!"

End Sub
