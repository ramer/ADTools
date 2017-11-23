Imports System.DirectoryServices.Protocols

Public Class clsWatcher
    Implements IDisposable

    Public Event ObjectChanged As EventHandler(Of EventArgs)

    Public Property Connection As LdapConnection
    Public Property Watchers As New HashSet(Of IAsyncResult)()

    Public Sub New(connection As LdapConnection)
        Me.Connection = connection
    End Sub

    Public Function Register(dn As String, scope As SearchScope, Optional timeout As TimeSpan = Nothing) As Boolean
        Try
            Dim searchRequest As New SearchRequest(dn, "(objectClass=*)", scope, Nothing)
            searchRequest.Controls.Add(New DirectoryNotificationControl())
            Watchers.Add(Connection.BeginSendRequest(searchRequest, If(timeout <> Nothing, timeout, TimeSpan.FromDays(365)), PartialResultProcessing.ReturnPartialResultsAndNotifyCallback, AddressOf Notify, searchRequest))
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub Notify(result As IAsyncResult)
        Dim prc As PartialResultsCollection = Connection.GetPartialResults(result)

        For Each entry As SearchResultEntry In prc
            OnObjectChanged(New ObjectChangedEventArgs(entry.DistinguishedName, entry))
        Next
    End Sub

    Private Sub OnObjectChanged(args As ObjectChangedEventArgs)
        RaiseEvent ObjectChanged(Me, args)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        For Each watcher In Watchers
            Connection.Abort(watcher)
        Next
    End Sub

End Class

Public Class ObjectChangedEventArgs
    Inherits EventArgs

    Public Property DistinguishedName As String
    Public Property Entry As SearchResultEntry

    Public Sub New(distinguishedName As String, entry As SearchResultEntry)
        Me.DistinguishedName = distinguishedName
        Me.Entry = entry
    End Sub
End Class