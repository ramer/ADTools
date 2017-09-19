Imports System.Collections.Specialized
Imports System.Threading
Imports System.Windows.Threading

Public Class clsThreadSafeObservableCollection(Of T)
    Implements IList(Of T)
    Implements INotifyCollectionChanged
    Private collection As IList(Of T) = New List(Of T)()
    Private dispatcher As Dispatcher
    Public Event CollectionChanged As NotifyCollectionChangedEventHandler Implements INotifyCollectionChanged.CollectionChanged
    Private sync As New ReaderWriterLock()

    Public Sub New()
        dispatcher = Dispatcher.CurrentDispatcher
    End Sub

    Public Sub New(collection As IList(Of T))
        dispatcher = Dispatcher.CurrentDispatcher

        sync.AcquireWriterLock(Timeout.Infinite)
        Me.collection = collection
        sync.ReleaseWriterLock()
    End Sub

    Public Sub Add(item As T) Implements IList(Of T).Add
        If Thread.CurrentThread Is dispatcher.Thread Then
            DoAdd(item)
        Else
            dispatcher.BeginInvoke(DirectCast(Sub() DoAdd(item), Action))
        End If
    End Sub

    Private Sub DoAdd(item As T)
        sync.AcquireWriterLock(Timeout.Infinite)
        collection.Add(item)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, item))
        sync.ReleaseWriterLock()
    End Sub

    Public Sub Clear() Implements IList(Of T).Clear
        If Thread.CurrentThread Is dispatcher.Thread Then
            DoClear()
        Else
            dispatcher.BeginInvoke(DirectCast(Sub() DoClear(), Action))
        End If
    End Sub

    Private Sub DoClear()
        sync.AcquireWriterLock(Timeout.Infinite)
        collection.Clear()
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        sync.ReleaseWriterLock()
    End Sub

    Public Function Contains(item As T) As Boolean Implements IList(Of T).Contains
        sync.AcquireReaderLock(Timeout.Infinite)
        Dim result = collection.Contains(item)
        sync.ReleaseReaderLock()
        Return result
    End Function

    Public Sub CopyTo(array As T(), arrayIndex As Integer) Implements IList(Of T).CopyTo
        sync.AcquireWriterLock(Timeout.Infinite)
        collection.CopyTo(array, arrayIndex)
        sync.ReleaseWriterLock()
    End Sub

    Public ReadOnly Property Count() As Integer Implements IList(Of T).Count
        Get
            sync.AcquireReaderLock(Timeout.Infinite)
            Dim result = collection.Count
            sync.ReleaseReaderLock()
            Return result
        End Get
    End Property

    Public ReadOnly Property IsReadOnly() As Boolean Implements IList(Of T).IsReadOnly
        Get
            Return collection.IsReadOnly
        End Get
    End Property

    Public Function Remove(item As T) As Boolean Implements IList(Of T).Remove
        If Thread.CurrentThread Is dispatcher.Thread Then
            Return DoRemove(item)
        Else
            Dim op = dispatcher.BeginInvoke(New Func(Of T, Boolean)(AddressOf DoRemove), item)
            If op Is Nothing OrElse op.Result Is Nothing Then
                Return False
            End If
            Return CBool(op.Result)
        End If
    End Function

    Private Function DoRemove(item As T) As Boolean
        sync.AcquireWriterLock(Timeout.Infinite)
        Dim index = collection.IndexOf(item)
        If index = -1 Then
            sync.ReleaseWriterLock()
            Return False
        End If
        Dim result = collection.Remove(item)
        If result Then 'AndAlso CollectionChanged IsNot Nothing Then
            RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        End If
        sync.ReleaseWriterLock()
        Return result
    End Function

    Public Function GetEnumerator() As IEnumerator(Of T) Implements IList(Of T).GetEnumerator
        Return collection.GetEnumerator()
    End Function

    Private Function System_Collections_IEnumerable_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return collection.GetEnumerator()
    End Function

    Public Function IndexOf(item As T) As Integer Implements IList(Of T).IndexOf
        sync.AcquireReaderLock(Timeout.Infinite)
        Dim result = collection.IndexOf(item)
        sync.ReleaseReaderLock()
        Return result
    End Function

    Public Sub Insert(index As Integer, item As T) Implements IList(Of T).Insert
        If Thread.CurrentThread Is dispatcher.Thread Then
            DoInsert(index, item)
        Else
            dispatcher.BeginInvoke(DirectCast(Sub() DoInsert(index, item), Action))
        End If
    End Sub

    Private Sub DoInsert(index As Integer, item As T)
        sync.AcquireWriterLock(Timeout.Infinite)
        collection.Insert(index, item)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, item, index))
        sync.ReleaseWriterLock()
    End Sub

    Public Sub RemoveAt(index As Integer) Implements IList(Of T).RemoveAt
        If Thread.CurrentThread Is dispatcher.Thread Then
            DoRemoveAt(index)
        Else
            dispatcher.BeginInvoke(DirectCast(Sub() DoRemoveAt(index), Action))
        End If
    End Sub

    Private Sub DoRemoveAt(index As Integer)
        sync.AcquireWriterLock(Timeout.Infinite)
        If collection.Count = 0 OrElse collection.Count <= index Then
            sync.ReleaseWriterLock()
            Return
        End If
        collection.RemoveAt(index)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        sync.ReleaseWriterLock()

    End Sub

    Default Public Property Item(index As Integer) As T Implements IList(Of T).Item
        Get
            sync.AcquireReaderLock(Timeout.Infinite)
            Dim result = collection(index)
            sync.ReleaseReaderLock()
            Return result
        End Get
        Set
            sync.AcquireWriterLock(Timeout.Infinite)
            If collection.Count = 0 OrElse collection.Count <= index Then
                sync.ReleaseWriterLock()
                Return
            End If
            collection(index) = Value
            sync.ReleaseWriterLock()
        End Set
    End Property

End Class
