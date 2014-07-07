Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Net

Public Class DomainChecking

    Public Event FinishedForDomain(ByVal finishedlist As ListViewItem)
    Public Event FinishedCom(ByVal Url As String, ByVal result As String, ByVal threadnum As Integer)
    Public Event FinishedNet(ByVal url As String, ByVal result As String, ByVal threadnum As Integer)
    Public Event FinishedOrg(ByVal url As String, ByVal result As String, ByVal threadnum As Integer)
    Public Event FinishedInfo(ByVal url As String, ByVal result As String, ByVal threadnum As Integer)
    Public Event FinishedBiz(ByVal url As String, ByVal result As String, ByVal threadnum As Integer)

    Public Sub SetListview(ByVal domain As String, ByVal com As String, ByVal net As String, ByVal org As String, ByVal info As String, ByVal biz As String)
        Dim item As New ListViewItem
        If domain = "" Then
            Exit Sub
        Else
            item.Text = domain
            item.UseItemStyleForSubItems = False
        End If
        Dim subitem1 As New ListViewItem.ListViewSubItem
        subitem1.Text = com
        If com = "Available" Then
            subitem1.BackColor = Color.LightGreen
        Else
            subitem1.BackColor = Color.PaleVioletRed
        End If

        item.SubItems.Add(subitem1)
        '''''.net
        Dim subitem2 As New ListViewItem.ListViewSubItem
        subitem2.Text = net
        If net = "Available" Then
            subitem2.BackColor = Color.LightGreen
        Else
            subitem2.BackColor = Color.PaleVioletRed
        End If

        item.SubItems.Add(subitem2)
        '''''.org
        Dim subitem3 As New ListViewItem.ListViewSubItem
        subitem3.Text = org
        If org = "Available" Then
            subitem3.BackColor = Color.LightGreen
        Else
            subitem3.BackColor = Color.PaleVioletRed
        End If

        item.SubItems.Add(subitem3)
        '''''.info
        Dim subitem4 As New ListViewItem.ListViewSubItem
        subitem4.Text = info
        If info = "Available" Then
            subitem4.BackColor = Color.LightGreen
        Else
            subitem4.BackColor = Color.PaleVioletRed
        End If

        item.SubItems.Add(subitem4)
        '''''.biz
        Dim subitem5 As New ListViewItem.ListViewSubItem
        subitem5.Text = biz

        If biz = "Available" Then
            subitem5.BackColor = Color.LightGreen
        Else
            subitem5.BackColor = Color.PaleVioletRed
        End If

        item.SubItems.Add(subitem5)
        RaiseEvent FinishedForDomain(item)
        ThreadM.completedThread(Thread.CurrentThread)
       
    End Sub
    Function check(ByVal fulldomainnosuffix As String)

        Dim com As String
        Dim net As String
        Dim org As String
        Dim info As String
        Dim biz As String
        Dim request As HttpWebRequest = HttpWebRequest.Create("http://www.networksolutions.com/whois/results.jsp?domain=" & fulldomainnosuffix & ".com")
        With request
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .Method = "GET"
            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            If Not dataresponse.ToLower.Contains("is currently registered to") Then
                com = "Available"
            Else
                com = "Unavailable"
            End If
            RaiseEvent FinishedCom(fulldomainnosuffix & ".com", com, Thread.CurrentThread.ManagedThreadId)
        End With
        Dim request1 As HttpWebRequest = HttpWebRequest.Create("http://www.networksolutions.com/whois/results.jsp?domain=" & fulldomainnosuffix & ".net")
        With request1
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .Method = "GET"
            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            If Not dataresponse.ToLower.Contains("is currently registered to") Then
                net = "Available"
            Else
                net = "Unavailable"
            End If
            RaiseEvent FinishedNet(fulldomainnosuffix & ".net", net, Thread.CurrentThread.ManagedThreadId)
        End With
        Dim request2 As HttpWebRequest = HttpWebRequest.Create("http://www.networksolutions.com/whois/results.jsp?domain=" & fulldomainnosuffix & ".org")
        With request2
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .Method = "GET"
            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            If Not dataresponse.ToLower.Contains("is currently registered to") Then
                org = "Available"
            Else
                org = "Unavailable"
            End If
            RaiseEvent FinishedOrg(fulldomainnosuffix & ".org", org, Thread.CurrentThread.ManagedThreadId)
        End With
        Dim request3 As HttpWebRequest = HttpWebRequest.Create("http://www.networksolutions.com/whois/results.jsp?domain=" & fulldomainnosuffix & ".info")
        With request3
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .Method = "GET"
            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            If Not dataresponse.ToLower.Contains("is currently registered to") Then
                info = "Available"
            Else
                info = "Unavailable"
            End If
            RaiseEvent FinishedInfo(fulldomainnosuffix & ".info", info, Thread.CurrentThread.ManagedThreadId)
        End With
        Dim request4 As HttpWebRequest = HttpWebRequest.Create("http://www.networksolutions.com/whois/results.jsp?domain=" & fulldomainnosuffix & ".biz")
        With request4
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .Method = "GET"
            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            If Not dataresponse.ToLower.Contains("is currently registered to") Then
                biz = "Available"
            Else
                biz = "Unavailable"
            End If
            RaiseEvent FinishedBiz(fulldomainnosuffix & ".biz", biz, Thread.CurrentThread.ManagedThreadId)
        End With
        SetListview(fulldomainnosuffix, com, net, org, info, biz)
    End Function
    

End Class
