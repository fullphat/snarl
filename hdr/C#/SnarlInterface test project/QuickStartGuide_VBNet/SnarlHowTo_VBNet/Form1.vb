Imports Snarl


Public Class Form1

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		Dim nReturnId As Integer
		nReturnId = SnarlConnector.ShowMessage("Test Title", "Test Text with 10s timeout", 10, "", 0, 0)
	End Sub
End Class
