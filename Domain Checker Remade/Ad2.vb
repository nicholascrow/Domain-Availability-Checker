Public Class Ad2

    Private Sub Ad2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.CenterToScreen()
        Button1.Enabled = False
        WebBrowser1.Navigate(Form1.adnum2)

        Timer1.Interval = 15000
        Timer1.Start()
        Timer1.Enabled = True

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Button1.Enabled = True
        Timer1.Stop()
        Timer1.Enabled = False
        Form1.Timer1.Start()
        Form1.Timer1.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Ad2_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
        Me.CenterToScreen()
    End Sub
End Class