Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim FileToDelete As String
        FileToDelete = System.AppDomain.CurrentDomain.BaseDirectory & "\AdwordsProcessor.exe"
        If System.IO.File.Exists(FileToDelete) = True Then
            System.IO.File.Delete(FileToDelete)
        End If
        Dim FileToDelete1 As String
        FileToDelete1 = System.AppDomain.CurrentDomain.BaseDirectory & "\Adwords Processor.exe"
        If System.IO.File.Exists(FileToDelete1) = True Then
            System.IO.File.Delete(FileToDelete1)
        End If
        My.Computer.Network.DownloadFile("http://triprogrammingsolutions.com/adwordsprocessor/AdwordsProcessor.exe", System.AppDomain.CurrentDomain.BaseDirectory & "\AdwordsProcessor.exe")
        Label1.Text = "Downloaded"
        Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "\AdwordsProcessor.exe")
        End
    End Sub
End Class
