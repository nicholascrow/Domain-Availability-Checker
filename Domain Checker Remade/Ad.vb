Public Class Ad

    Private Sub Ad_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        Me.Width = Screen.PrimaryScreen.Bounds.Width
'        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.CenterToScreen()
        Button1.Enabled = False
        Timer1.Interval = 15000
        Timer1.Enabled = True
        Timer1.Start()
        If Form1.ads.Contains(",") Then
            If Form1.ads.Contains("http") Then
                PictureBox1.ImageLocation = Form1.ads.Split(",")(1)
                LinkLabel1.Text = Form1.ads.Split(",")(0)
            End If
        End If
        If Form1.ads1.Contains(",") Then
            If Form1.ads1.Contains("http") Then
                PictureBox2.ImageLocation = Form1.ads1.Split(",")(1)
                LinkLabel2.Text = Form1.ads1.Split(",")(0)
            End If
        End If
        If Form1.ads2.Contains(",") Then
            If Form1.ads2.Contains("http") Then
                PictureBox3.ImageLocation = Form1.ads2.Split(",")(1)
                LinkLabel3.Text = Form1.ads2.Split(",")(0)
            End If
        End If
        If Form1.ads3.Contains(",") Then
            If Form1.ads3.Contains("http") Then
                PictureBox4.ImageLocation = Form1.ads3.Split(",")(1)
                LinkLabel4.Text = Form1.ads3.Split(",")(0)
            End If
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start(Form1.ads.Split(",")(2))
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start(Form1.ads1.Split(",")(2))
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Process.Start(Form1.ads2.Split(",")(2))
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Process.Start(Form1.ads3.Split(",")(2))
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Button1.Enabled = True
        Form1.Timer1.Start()
        Form1.Timer1.Enabled = True
    End Sub

    Private Sub Ad_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
        Me.CenterToScreen()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Process.Start(Form1.ads.Split(",")(2))
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Process.Start(Form1.ads1.Split(",")(2))
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        Process.Start(Form1.ads2.Split(",")(2))
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        Process.Start(Form1.ads3.Split(",")(2))
    End Sub
End Class