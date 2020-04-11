Imports System.DirectoryServices.Protocols

Public Class clsWatcher
    Implements IDisposable

    ''' <summary>
    ''' Fires when LDAP-object has changes
    ''' </summary>
    Public Event ObjectChanged As EventHandler(Of ObjectChangedEventArgs)
    Private Property connection As LdapConnection
    Private Property asyncResult As IAsyncResult

    Public Sub New(connection As LdapConnection)
        Me.connection = connection
    End Sub

    ''' <summary>
    ''' Registers watcher on current distinguished name by scope with optional timeout
    ''' </summary>
    ''' <param name="dn">LDAP-object distinguished name</param>
    ''' <param name="scope">Watcher scope</param>
    ''' <param name="timeout">Timeout (default is 1 year)</param>
    ''' <returns>True on success, False otherwise</returns>
    Public Function Register(dn As String, scope As SearchScope, Optional timeout As TimeSpan = Nothing) As Boolean
        Try
            Dim searchRequest As New SearchRequest(dn, "(objectClass=*)", scope, Nothing)
            searchRequest.Controls.Add(New DirectoryNotificationControl())
            asyncResult = connection.BeginSendRequest(searchRequest, If(timeout <> Nothing, timeout, TimeSpan.FromDays(365)), PartialResultProcessing.ReturnPartialResultsAndNotifyCallback, AddressOf Notify, searchRequest)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub Notify(result As IAsyncResult)
        Dim prc As PartialResultsCollection = connection.GetPartialResults(result)

        For Each entry As SearchResultEntry In prc
            OnObjectChanged(New ObjectChangedEventArgs(entry.DistinguishedName, entry))
        Next

        If result.IsCompleted Then MsgBox("")
    End Sub

    Private Sub OnObjectChanged(args As ObjectChangedEventArgs)
        RaiseEvent ObjectChanged(Me, args)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        connection.Abort(asyncResult)
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