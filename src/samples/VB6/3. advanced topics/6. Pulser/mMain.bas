Attribute VB_Name = "mMain"
Option Explicit

Public Sub Main()

    Load Form1
    
    With New BMsgLooper
        .Run

    End With

    Unload Form1

End Sub
