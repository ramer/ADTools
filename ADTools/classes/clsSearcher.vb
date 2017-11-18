Imports System.Collections.ObjectModel
Imports System.DirectoryServices.Protocols
Imports System.Security.Principal
Imports System.Threading

Public Class clsSearcher
    Private basicsearchtasks As New List(Of Task)
    Private basicsearchtaskscts As New CancellationTokenSource
    Public Event BasicSearchAsyncDataRecieved()

    Private tombstonesearchtasks As New List(Of Task)
    Private tombstonesearchtaskscts As New CancellationTokenSource

    Sub New()

    End Sub

    Public Async Function BasicSearchStopAsync() As Task
        basicsearchtaskscts.Cancel()
        If basicsearchtasks.Count > 0 Then
            Try
                Await Task.WhenAll(basicsearchtasks.ToArray) 'magic
            Catch
            End Try

            basicsearchtasks.Clear()
            Log("Task list cleared")
        End If
    End Function

    Public Async Function BasicSearchAsync(returncollection As clsThreadSafeObservableCollection(Of clsDirectoryObject),
                                           Optional parentobject As clsDirectoryObject = Nothing,
                                           Optional filter As clsFilter = Nothing,
                                           Optional specificdomains As ObservableCollection(Of clsDomain) = Nothing,
                                           Optional attributenames As String() = Nothing) As Task
        Await BasicSearchStopAsync()

        returncollection.Clear()

        basicsearchtaskscts = New CancellationTokenSource

        Dim roots As List(Of clsDirectoryObject)
        Dim searchscope As SearchScope

        If parentobject IsNot Nothing Then
            roots = New List(Of clsDirectoryObject) From {parentobject}
            searchscope = SearchScope.OneLevel
        Else
            roots = If(specificdomains Is Nothing, domains.Select(Function(d As clsDomain) New clsDirectoryObject(d.DefaultNamingContext, d)).ToList, specificdomains.Select(Function(d As clsDomain) New clsDirectoryObject(d.DefaultNamingContext, d)).ToList)
            searchscope = SearchScope.Subtree
        End If

        For Each root In roots
            Dim mt = Task.Factory.StartNew(
                Function()
                    Return BasicSearchSync(root, filter, searchscope, basicsearchtaskscts.Token, attributenames)
                End Function, basicsearchtaskscts.Token)
            basicsearchtasks.Add(mt)
            Log(String.Format("Task created: domain search ""{0}""", root.name))

            Dim lt = mt.ContinueWith(
                Function(parenttask As Task(Of ObservableCollection(Of clsDirectoryObject)))
                    basicsearchtasks.Remove(mt)
                    Log(String.Format("Task created: collect domain {1} search results ({0})", parenttask.Result.Count, If(parenttask.Result.Count > 0, parenttask.Result(0).Domain.Name, "(unknown)")))
                    For Each result In parenttask.Result
                        If basicsearchtaskscts.Token.IsCancellationRequested Then Return False
                        returncollection.Add(result)
                        'If parenttask.Result.Count > 50 Then Thread.Sleep(50)
                    Next
                    Log("Task completed: domain search")
                    Return True
                End Function)
            basicsearchtasks.Add(lt)
            Log(String.Format("Task created: display domain ""{0}"" search results", root.name))

            Dim ft = lt.ContinueWith(
                Function(parenttask As Task(Of Boolean))
                    Try
                        basicsearchtasks.Remove(lt)
                    Catch
                        basicsearchtasks.Clear()
                    End Try
                    Log("Task completed: display search results")
                    Return True
                End Function)
        Next

        If basicsearchtasks.Count > 0 Then Await Task.WhenAny(basicsearchtasks.ToArray)
        RaiseEvent BasicSearchAsyncDataRecieved()
        If basicsearchtasks.Count > 0 Then Await Task.WhenAll(basicsearchtasks.ToArray)
    End Function

    Public Function BasicSearchSync(Optional root As clsDirectoryObject = Nothing,
                                    Optional filter As clsFilter = Nothing,
                                    Optional searchscope As SearchScope = SearchScope.Subtree,
                                    Optional ct As CancellationToken = Nothing,
                                    Optional attributenames As String() = Nothing) As ObservableCollection(Of clsDirectoryObject)

        If root Is Nothing Then Return New ObservableCollection(Of clsDirectoryObject)

        Dim results As New ObservableCollection(Of clsDirectoryObject)

        Try
            Dim searchRequest As New SearchRequest()
            searchRequest.DistinguishedName = root.distinguishedName

            If searchscope = SearchScope.Subtree Then
                If filter IsNot Nothing AndAlso Not String.IsNullOrEmpty(filter.Filter) Then searchRequest.Filter = filter.Filter
            End If

            searchRequest.Scope = searchscope
            If attributenames IsNot Nothing Then searchRequest.Attributes.AddRange(attributenames)

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

                    Dim obj As New clsDirectoryObject(entry, root.Domain)
                    Dim ma As New List(Of String)
                    If attributenames IsNot Nothing Then
                        For Each attr As String In attributenames
                            If entry.Attributes(attr) Is Nothing Then ma.Add(attr)
                        Next
                        obj.MissedAttributes = ma
                    End If

                    results.Add(obj)
                Next

                pageRequestControl.Cookie = pageResponseControl.Cookie
            Loop While pageResponseControl IsNot Nothing AndAlso pageResponseControl.Cookie.Length > 0

            If results.Count > 0 AndAlso results(0).distinguishedName = root.distinguishedName Then results.RemoveAt(0)

        Catch ex As Exception
            ThrowException(ex, "BasicSearchSync")
        End Try

        Return results

    End Function


End Class
