Imports System
Imports System.Threading
Public Class ThreadM
    Public Shared Sub completedThread(ByVal _thread As Thread)
        _threads.Remove(_thread)
        If (((_threads.Count < Maxthreads) AndAlso (_threadsQueue.Count > 0)) AndAlso (_threadsQueue.Count > 0)) Then
            Dim thread As Thread
            SyncLock _threadsQueue
                thread = _threadsQueue.List.Item(0)
                _threadsQueue.Remove(thread)
            End SyncLock
            _threads.Add(thread)
            thread.Start()
        End If
    End Sub

    Public Shared Sub startThread(ByVal _thread As Thread)
        If ((Maxthreads = 0) OrElse (_threads.Count < Maxthreads)) Then
            _threads.Add(_thread)
            _thread.Start()
        Else
            _threadsQueue.Add(_thread)
        End If
    End Sub
    Public Shared Function Threadscomplete() As Boolean
        If _threadsQueue.Count <> 0 Then
            Return False
        Else : Return True
        End If
    End Function
    Public Shared Sub stopThreads()
        _threadsQueue.Clear()
        SyncLock _threads.List
            Dim thread As Thread
            For Each thread In _threads.List
                thread.Abort()
            Next
        End SyncLock
        _threads.Clear()
    End Sub

    Public Shared Sub SuspendThreads()
        Dim thread As Thread
        For Each thread In _threads.List
            thread.Suspend()
        Next
    End Sub
    Public Shared Sub ResumeThreads()
        Dim thread As Thread
        For Each thread In _threads.List
            thread.Resume()
        Next
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return _threads.Count
        End Get
    End Property

    Public ReadOnly Property hasRunningThreads As Boolean
        Get
            Return (_threads.Count > 0)
        End Get
    End Property

    'Public Shared Property Instance As ThreadM
    '    Get
    '        If (ThreadM._instance Is Nothing) Then
    '            ThreadM._instance = New ThreadM
    '        End If
    '        Return ThreadM._instance
    '    End Get
    '    Set(ByVal value As ThreadM)

    '    End Set
    'End Property


    ' Fields
    Public Shared Maxthreads As Integer
    ' Private Shared _instance As ThreadM
    Public Shared _threads As threadslist(Of Thread) = New threadslist(Of Thread)
    Public Shared _threadsQueue As threadslist(Of Thread) = New threadslist(Of Thread)
End Class
