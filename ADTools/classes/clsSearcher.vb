Imports System.Collections.ObjectModel
Imports System.DirectoryServices.Protocols
Imports System.Security.Principal
Imports System.Threading
Imports System.Windows.Threading

Public Class clsSearcher
    Private dispatcher As Dispatcher

    Public Event SearchAsyncDataRecieved()
    Public Event SearchAsyncCompleted()

    Private asyncResultCollection As New List(Of IAsyncResult)

    Sub New()
        dispatcher = Dispatcher.CurrentDispatcher
    End Sub

    Private Sub NotifySearchAsyncDataRecieved()
        If Thread.CurrentThread Is dispatcher.Thread Then
            RaiseEvent SearchAsyncDataRecieved()
        Else
            dispatcher.BeginInvoke(DirectCast(Sub() RaiseEvent SearchAsyncDataRecieved(), Action))
        End If
    End Sub

    Private Sub NotifySearchAsyncCompleted()
        If Thread.CurrentThread Is dispatcher.Thread Then
            RaiseEvent SearchAsyncCompleted()
        Else
            dispatcher.BeginInvoke(DirectCast(Sub() RaiseEvent SearchAsyncCompleted(), Action))
        End If
    End Sub

    Public Function SearchSync(Optional root As clsDirectoryObject = Nothing,
                                    Optional filter As clsFilter = Nothing,
                                    Optional searchscope As SearchScope = SearchScope.Subtree,
                                    Optional ct As CancellationToken = Nothing) As ObservableCollection(Of clsDirectoryObject)

        If root Is Nothing Then Return New ObservableCollection(Of clsDirectoryObject)

        Dim results As New ObservableCollection(Of clsDirectoryObject)

        Try
            Dim searchRequest As New SearchRequest()
            searchRequest.DistinguishedName = root.distinguishedName

            If searchscope = SearchScope.Subtree Then
                If filter IsNot Nothing AndAlso Not String.IsNullOrEmpty(filter.Filter) Then searchRequest.Filter = filter.Filter
            End If

            searchRequest.Attributes.AddRange({"objectguid"})
            searchRequest.Scope = searchscope

            Dim sortRequestControl As New SortRequestControl("name", False)
            searchRequest.Controls.Add(sortRequestControl)
            Dim pageRequestControl As New PageResultRequestControl(1000)
            searchRequest.Controls.Add(pageRequestControl)
            Dim searchOptionsControl As New SearchOptionsControl(SearchOption.DomainScope)
            searchRequest.Controls.Add(searchOptionsControl)

            Dim searchResponse As SearchResponse
            Dim pageResponseControl As PageResultResponseControl
            Do
                searchResponse = root.Connection.SendRequest(searchRequest)
                pageResponseControl = searchResponse.Controls.Where(Function(rc) TypeOf rc Is PageResultResponseControl).First

                For Each entry As SearchResultEntry In searchResponse.Entries
                    If Not ct = Nothing AndAlso ct.IsCancellationRequested Then Return New ObservableCollection(Of clsDirectoryObject)
                    results.Add(New clsDirectoryObject(entry, root.Domain))
                Next

                pageRequestControl.Cookie = pageResponseControl.Cookie
            Loop While pageResponseControl IsNot Nothing AndAlso pageResponseControl.Cookie.Length > 0

            If results.Count > 0 AndAlso results(0).distinguishedName = root.distinguishedName Then results.RemoveAt(0)

        Catch ex As Exception
            ThrowException(ex, "BasicSearchSync")
        End Try

        Return results

    End Function

    Public Sub SearchAsync(returnCollection As clsThreadSafeObservableCollection(Of clsDirectoryObject),
                            Optional parentObject As clsDirectoryObject = Nothing,
                            Optional filter As clsFilter = Nothing,
                            Optional specificDomains As ObservableCollection(Of clsDomain) = Nothing)

        StopAllSearchAsync()

        returnCollection.Clear()

        Dim roots As List(Of clsDirectoryObject)
        Dim searchscope As SearchScope

        If parentObject IsNot Nothing Then
            roots = New List(Of clsDirectoryObject) From {parentObject}
            searchscope = SearchScope.OneLevel
        Else
            roots = If(specificDomains Is Nothing, domains.Select(Function(d As clsDomain) New clsDirectoryObject(d.DefaultNamingContext, d)).ToList, specificDomains.Select(Function(d As clsDomain) New clsDirectoryObject(d.DefaultNamingContext, d)).ToList)
            searchscope = SearchScope.Subtree
        End If

        For Each root In roots
            Try
                Dim searchRequest As New SearchRequest()
                searchRequest.DistinguishedName = root.distinguishedName

                If searchscope = SearchScope.Subtree Then
                    If filter IsNot Nothing AndAlso Not String.IsNullOrEmpty(filter.Filter) Then searchRequest.Filter = filter.Filter
                End If

                searchRequest.Attributes.AddRange(attributesToLoadDefault)
                searchRequest.Scope = searchscope

                Dim sortRequestControl As New SortRequestControl("name", False)
                searchRequest.Controls.Add(sortRequestControl)
                Dim pageRequestControl As New PageResultRequestControl(100)
                searchRequest.Controls.Add(pageRequestControl)
                Dim searchOptionsControl As New SearchOptionsControl(SearchOption.DomainScope)
                searchRequest.Controls.Add(searchOptionsControl)

                Dim helper As New clsSearcherHelper
                helper.root = root
                helper.attributes = attributesToLoadDefault
                helper.returnCollection = returnCollection
                helper.pageRequestControl = pageRequestControl
                helper.searchRequest = searchRequest

                asyncResultCollection.Add(root.Connection.BeginSendRequest(searchRequest, PartialResultProcessing.NoPartialResultSupport, AddressOf SearchAsyncResult, helper))

            Catch ex As Exception
                ThrowException(ex, "SearchAsync")
            End Try
        Next

    End Sub

    Private Sub SearchAsyncResult(asyncResult As IAsyncResult)
        asyncResultCollection.Remove(asyncResult)
        If asyncResultCollection.Count = 0 Then NotifySearchAsyncCompleted()

        Try
            Dim helper As clsSearcherHelper = asyncResult.AsyncState
            If helper.aborted Then Exit Sub

            Dim searchResponse As SearchResponse = DirectCast(helper.root.Connection.EndSendRequest(asyncResult), SearchResponse)
            Dim pageResponseControl As PageResultResponseControl = searchResponse.Controls.Where(Function(rc) TypeOf rc Is PageResultResponseControl).First

            Dim results As New List(Of clsDirectoryObject)
            For Each entry As SearchResultEntry In searchResponse.Entries

                Dim ma As New List(Of String)
                If helper.attributes IsNot Nothing Then
                    For Each attr As String In helper.attributes
                        If entry.Attributes(attr) Is Nothing Then ma.Add(attr)
                    Next
                End If

                results.Add(New clsDirectoryObject(entry, helper.root.Domain) With {.MissedAttributes = ma})
            Next

            If results.Count > 0 Then NotifySearchAsyncDataRecieved()

            helper.returnCollection.AddRange(results)

            helper.pageRequestControl.Cookie = pageResponseControl.Cookie
            If pageResponseControl.Cookie.Length > 0 Then asyncResultCollection.Add(helper.root.Connection.BeginSendRequest(helper.searchRequest, PartialResultProcessing.NoPartialResultSupport, AddressOf SearchAsyncResult, helper))

        Catch ex As Exception
            ThrowException(ex, "SearchAsync Callback")
        End Try
    End Sub

    Public Sub StopAllSearchAsync()
        For Each asyncResult In asyncResultCollection
            CType(asyncResult.AsyncState, clsSearcherHelper).aborted = True
        Next
    End Sub

End Class

Public Class clsSearcherHelper
    Public Property aborted As Boolean
    Public Property root As clsDirectoryObject
    Public Property attributes As String()
    Public Property returnCollection As clsThreadSafeObservableCollection(Of clsDirectoryObject)
    Public Property pageRequestControl As PageResultRequestControl
    Public Property searchRequest As New SearchRequest
End Class