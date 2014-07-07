Imports System.Threading
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions

Public Class Form1

#Region "Declarations"
    Public Maxthreads As Integer
    WithEvents x As New DomainChecking
    Public ads As String
    Public ads1 As String
    Public ads2 As String
    Public ads3 As String
    Public adnum2 As String
    Dim adtime As Integer = 1
    Dim n As New Http
#End Region
#Region "Events From Check"
    Private Sub x_FinishedInfo(ByVal url As String, ByVal result As String, ByVal threadnum As Integer) Handles x.FinishedInfo
        Dim l As New ListViewItem
        l.Text = "Completed checking for: " & url
        If result = "Available" Then
            l.SubItems.Add("Available")
            l.BackColor = Color.LightGreen
        Else
            l.SubItems.Add("Unavailable")
            l.BackColor = Color.PaleVioletRed
        End If
        '     l.SubItems.Add(threadnum)
        ListView2.Items.Add(l)
        ListView2.Update()
    End Sub
    Private Sub x_FinishedNet(ByVal url As String, ByVal result As String, ByVal threadnum As Integer) Handles x.FinishedNet
        Dim l As New ListViewItem
        l.Text = "Completed checking for: " & url
        If result = "Available" Then
            l.SubItems.Add("Available")
            l.BackColor = Color.LightGreen
        Else
            l.SubItems.Add("Unavailable")
            l.BackColor = Color.PaleVioletRed
        End If
        '   l.SubItems.Add(threadnum)
        ListView2.Items.Add(l)
        ListView2.Update()
    End Sub
    Private Sub x_FinishedOrg(ByVal url As String, ByVal result As String, ByVal threadnum As Integer) Handles x.FinishedOrg
        Dim l As New ListViewItem
        l.Text = "Completed checking for: " & url
        If result = "Available" Then
            l.SubItems.Add("Available")
            l.BackColor = Color.LightGreen
        Else
            l.SubItems.Add("Unavailable")
            l.BackColor = Color.PaleVioletRed
        End If
        '  l.SubItems.Add(threadnum)
        ListView2.Items.Add(l)
        ListView2.Update()
    End Sub
    Private Sub x_FinishedBiz(ByVal url As String, ByVal result As String, ByVal threadnum As Integer) Handles x.FinishedBiz
        Dim l As New ListViewItem
        l.Text = "Completed checking for: " & url
        If result = "Available" Then
            l.SubItems.Add("Available")
            l.BackColor = Color.LightGreen
        Else
            l.SubItems.Add("Unavailable")
            l.BackColor = Color.PaleVioletRed
        End If
        '  l.SubItems.Add(threadnum)
        ListView2.Items.Add(l)
        ListView2.Update()
    End Sub
    Private Sub x_FinishedCom(ByVal Url As String, ByVal result As String, ByVal threadnum As Integer) Handles x.FinishedCom
        Dim l As New ListViewItem
        l.Text = "Completed checking for: " & Url
        If result = "Available" Then
            l.SubItems.Add("Available")
            l.BackColor = Color.LightGreen
        Else
            l.SubItems.Add("Unavailable")
            l.BackColor = Color.PaleVioletRed
        End If
        '  l.SubItems.Add(threadnum)
        ListView2.Items.Add(l)
        ListView2.Update()
    End Sub
    Private Sub x_FinishedForDomain(ByVal finishedlist As ListViewItem) Handles x.FinishedForDomain
        ListView1.Items.Add(finishedlist)
    End Sub
#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ThreadM.Maxthreads = TrackBar1.Value
        For Each domain As String In TextBox1.Lines
            Dim d As String = domain
            Dim xx As String = d
            If d.Contains(".") Then
                xx = d.Split(".")(0)
            End If
            Dim t As New Thread(New ThreadStart(Sub() x.check(xx)))
            ThreadM.startThread(t)
        Next domain
    End Sub
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        ThreadM.stopThreads()

        'Dim data = n.UrlEncode("[" & My.Computer.Clock.GmtTime & "]  " & My.Computer.Name & " Username:" & Environment.UserName & "  NEWUSERLOGGEDOUT")
        'Dim request1 As HttpWebRequest = HttpWebRequest.Create("http://triprogrammingsolutions.com/Programs/DomainCheckerFree/add.php?stringData=" & data)
        'Dim response1 As System.Net.HttpWebResponse = request1.GetResponse
        'Dim sr1 As System.IO.StreamReader = New System.IO.StreamReader(response1.GetResponseStream())
        'Dim response12 As String = sr1.ReadToEnd
        End
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Interval = 600000
        ThreadM.Maxthreads = Maxthreads
        CheckForIllegalCrossThreadCalls = False
        Me.CenterToScreen()
        Maxthreads = TrackBar1.Value
        Try
            'Try
            '    Dim request As HttpWebRequest = HttpWebRequest.Create("http://triprogrammingsolutions.com/Programs/DomainCheckerFree/Settings")
            '    With request
            '        .Referer = "http://www.google.com"
            '        .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            '        .KeepAlive = True
            '        .Method = "GET"
            '        Dim response As System.Net.HttpWebResponse = .GetResponse
            '        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            '        Dim dataresponse As String = sr.ReadToEnd
            '        If Not GetBetween(dataresponse, "[VERSION]", "[VERSION]") = "1.0.1" Then
            '            MsgBox("New Version Available!")
            '            Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "\Update.exe")
            '            End
            '        End If
            '        ads = GetBetween(dataresponse, "[AD]", "[AD]")
            '        ads1 = GetBetween(dataresponse, "[AD2]", "[AD2]")
            '        ads2 = GetBetween(dataresponse, "[AD3]", "[AD3]")
            '        ads3 = GetBetween(dataresponse, "[AD4]", "[AD4]")
            '        adnum2 = GetBetween(dataresponse, "[ADNUM2]", "[ADNUM2]")

            '    End With
            '    Dim data = n.UrlEncode("[" & My.Computer.Clock.GmtTime & "]  " & My.Computer.Name & " Username:" & Environment.UserName & "  NEWUSERLOGGEDIN")
            '    Dim request1 As HttpWebRequest = HttpWebRequest.Create("http://triprogrammingsolutions.com/Programs/DomainCheckerFree/add.php?stringData=" & data)
            '    Dim response1 As System.Net.HttpWebResponse = request1.GetResponse
            '    Dim sr1 As System.IO.StreamReader = New System.IO.StreamReader(response1.GetResponseStream())
            '    Dim response12 As String = sr1.ReadToEnd
            '     Ad.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("You have an internet problem. Please Fix it and open again.")
            End
        End Try
    End Sub

    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        TextBox2.Text = SaveFileDialog1.FileName
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SaveFileDialog1.Filter = "CSV files (*.csv)|*.csv"
        SaveFileDialog1.FilterIndex = 1
        SaveFileDialog1.RestoreDirectory = True
        SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub ExportResultsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportResultsToolStripMenuItem.Click
        Dim os As New StreamWriter(TextBox2.Text)
        For i As Integer = 0 To ListView1.Columns.Count - 1
            os.Write("""" & ListView1.Columns(i).Text.Replace("""", """""") & """,")
        Next

        os.WriteLine()

        For i As Integer = 0 To ListView1.Items.Count - 1
            For j As Integer = 0 To ListView1.Columns.Count - 1
                os.Write("""" & ListView1.Items(i).SubItems(j).Text.Replace("""", """""") + """,")
            Next

            os.WriteLine()

        Next

        os.Close()
    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Maxthreads = TrackBar1.Value
    End Sub

    Function GetBetween(ByVal source As String, ByVal Get1 As String, ByVal Get2 As String)
        Dim lslasrrep1 As String = Get1.Replace("[", "\[")
        Dim rslasrrep1 As String = lslasrrep1.Replace("]", "\]")
        Dim lslasrrep2 As String = Get2.Replace("[", "\[")
        Dim rslasrrep2 As String = lslasrrep2.Replace("]", "\]")
        Dim r As New System.Text.RegularExpressions.Regex(rslasrrep1 & ".*" & rslasrrep2)
        Dim matchz As MatchCollection = r.Matches(source)
        If matchz(0) Is Nothing Then Throw New Exception("No Matches!")
        Dim realmatch As String = matchz(0).ToString.Replace(Get1, Nothing)
        Dim realmatch1 As String = matchz(0).ToString.Replace(Get2, Nothing)
        Return realmatch1
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If adtime = 1 Then
            Timer1.Stop()
            Timer1.Enabled = False
            adtime = 2
            Ad.Show()
        ElseIf adtime = 2 Then
            Timer1.Stop()
            Timer1.Enabled = False
            adtime = 1
            Ad2.Show()
        End If
    End Sub
End Class
