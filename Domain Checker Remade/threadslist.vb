Public Class threadslist(Of t)
    ' Methods
    Public Sub New()
        Me.lst = New List(Of T)
    End Sub

    Public Sub Add(ByVal item As T)
        SyncLock Me.lst
            Me.lst.Add(item)
        End SyncLock
    End Sub

    Public Sub Clear()
        SyncLock Me.lst
            Me.lst.Clear()
        End SyncLock
    End Sub

    Public Sub Remove(ByVal item As T)
        SyncLock Me.lst
            Me.lst.Remove(item)
        End SyncLock
    End Sub


    ' Properties
    Public ReadOnly Property Count As Integer
        Get
            SyncLock Me.lst
                Return Me.lst.Count
            End SyncLock
        End Get
    End Property

    Public ReadOnly Property List As List(Of T)
        Get
            Return Me.lst
        End Get
    End Property


    ' Fields
    Private lst As List(Of T)
End Class
