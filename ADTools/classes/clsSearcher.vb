Imports System.Collections.ObjectModel
Imports System.DirectoryServices
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

    Public Async Function BasicSearchAsync(ByVal returncollection As ObservableCollection(Of clsDirectoryObject),
                                           Optional parentobject As clsDirectoryObject = Nothing,
                                           Optional filter As clsFilter = Nothing,
                                           Optional specificdomains As ObservableCollection(Of clsDomain) = Nothing
                                           ) As Task
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

        Dim results As New List(Of clsDirectoryObject)

        For Each root In roots
            Dim mt = Task.Factory.StartNew(
                Sub()
                    results.AddRange(BasicSearchSync(root, filter, searchscope, basicsearchtaskscts.Token))
                End Sub, basicsearchtaskscts.Token)
            basicsearchtasks.Add(mt)
            Log(String.Format("Task created: domain {0} search", root.name))

            Dim ft = mt.ContinueWith(
                Function(parenttask As Task(Of Boolean))
                    Try
                        basicsearchtasks.Remove(mt)
                    Catch
                        basicsearchtasks.Clear()
                    End Try
                    Log("Task completed: display search results")
                    Return True
                End Function)
        Next

        'If basicsearchtasks.Count > 0 Then Await Task.WhenAny(basicsearchtasks.ToArray)
        If basicsearchtasks.Count > 0 Then Await Task.WhenAll(basicsearchtasks.ToArray)

        RaiseEvent BasicSearchAsyncDataRecieved()

        For Each obj In results
            returncollection.Add(obj)
        Next
    End Function

    Public Function BasicSearchSync(Optional root As clsDirectoryObject = Nothing,
                                    Optional filter As clsFilter = Nothing,
                                    Optional searchscope As SearchScope = SearchScope.Subtree,
                                    Optional ct As CancellationToken = Nothing) As ObservableCollection(Of clsDirectoryObject)

        If root Is Nothing Then Return New ObservableCollection(Of clsDirectoryObject)

        Dim results As New ObservableCollection(Of clsDirectoryObject)

        Try
            Dim ldapsearcher As New DirectorySearcher(root.Entry)
            ldapsearcher.SearchScope = searchscope
            ldapsearcher.PropertiesToLoad.Add("objectGuid")
            ldapsearcher.PageSize = 1000

            If searchscope = SearchScope.Subtree Then
                If filter IsNot Nothing AndAlso Not String.IsNullOrEmpty(filter.Filter) Then ldapsearcher.Filter = filter.Filter
            End If

            Dim ldapresults As SearchResultCollection = ldapsearcher.FindAll()

            For Each ldapresult As SearchResult In ldapresults
                If Not ct = Nothing AndAlso ct.IsCancellationRequested Then Return New ObservableCollection(Of clsDirectoryObject)
                results.Add(New clsDirectoryObject(ldapresult, root.Domain))
            Next
            ldapresults.Dispose()

        Catch ex As Exception
            ThrowException(ex, "BasicSearchSync")
        End Try

        Return results
    End Function

    'Public Async Function TombstoneSearchAsync(returncollection As clsThreadSafeObservableCollection(Of clsDeletedDirectoryObject), pattern As String,
    '                                  Optional specificdomains As ObservableCollection(Of clsDomain) = Nothing,
    '                                  Optional freesearch As Boolean = False) As Task

    '    tombstonesearchtaskscts.Cancel()
    '    If tombstonesearchtasks.Count > 0 Then Exit Function
    '    returncollection.Clear()
    '    tombstonesearchtaskscts = New CancellationTokenSource

    '    Dim domainlist As ObservableCollection(Of clsDomain) = If(specificdomains Is Nothing, domains, specificdomains)

    '    For Each dmn In domainlist

    '        Dim mt = Task.Factory.StartNew(
    '            Function()
    '                Return TombstoneSearchSync(pattern, dmn, freesearch, tombstonesearchtaskscts.Token)
    '            End Function, tombstonesearchtaskscts.Token)
    '        tombstonesearchtasks.Add(mt)

    '        Dim lt = mt.ContinueWith(
    '            Function(parenttask As Task(Of ObservableCollection(Of clsDeletedDirectoryObject)))
    '                tombstonesearchtasks.Remove(mt)
    '                For Each result In parenttask.Result
    '                    If tombstonesearchtaskscts.Token.IsCancellationRequested Then Return False
    '                    returncollection.Add(result)
    '                    Thread.Sleep(30)
    '                Next
    '                Return True
    '            End Function)
    '        tombstonesearchtasks.Add(lt)

    '        Dim ft = lt.ContinueWith(
    '            Function(parenttask As Task(Of Boolean))
    '                tombstonesearchtasks.Remove(lt)
    '                Return True
    '            End Function)

    '    Next

    '    Await Task.WhenAny(tombstonesearchtasks.ToArray)
    'End Function

    'Public Function TombstoneSearchSync(pattern As String,
    '                                  domain As clsDomain,
    '                                  Optional freesearch As Boolean = False,
    '                                  Optional ct As CancellationToken = Nothing) As ObservableCollection(Of clsDeletedDirectoryObject)

    '    Dim properties As String() = {"cn",
    '                                  "distinguishedName",
    '                                  "lastKnownParent",
    '                                  "name",
    '                                  "objectClass",
    '                                  "objectGUID",
    '                                  "objectSID",
    '                                  "sAMAccountName",
    '                                  "userAccountControl",
    '                                  "whenChanged",
    '                                  "whenCreated"}

    '    If domain.Validated = False Then Return New ObservableCollection(Of clsDeletedDirectoryObject)

    '    Dim results As New ObservableCollection(Of clsDeletedDirectoryObject)

    '    Try
    '        Dim patterns() As String = pattern.Split({"/", vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

    '        For Each singlepattern As String In patterns
    '            If ct.IsCancellationRequested Then Return New ObservableCollection(Of clsDeletedDirectoryObject)

    '            singlepattern = Trim(singlepattern)
    '            If String.IsNullOrEmpty(singlepattern) Then Continue For

    '            If singlepattern.StartsWith("""") And singlepattern.EndsWith("""") And Len(singlepattern) > 2 Then
    '                singlepattern = Mid(singlepattern, 2, Len(singlepattern) - 2)
    '            Else
    '                singlepattern = If(freesearch, "*" & singlepattern & "*", singlepattern & "*")
    '            End If

    '            Dim filter As String = "(&" +
    '                                        "(|" +
    '                                            "(&(objectClass=person)(!(objectClass=inetOrgPerson)))" +
    '                                            "(objectClass=computer)" +
    '                                            "(objectClass=group)" +
    '                                            "(objectClass=organizationalUnit)" +
    '                                        ")" +
    '                                        "(|" +
    '                                            "(name=" & singlepattern & ")" +
    '                                        ")" +
    '                                        "(isDeleted=TRUE)" +
    '                                     ")"

    '            Dim ldapsearcher As New DirectorySearcher(domain.SearchRoot)
    '            ldapsearcher.PropertiesToLoad.AddRange(properties)
    '            ldapsearcher.Filter = filter
    '            ldapsearcher.PageSize = 1000
    '            ldapsearcher.Tombstone = True

    '            Dim ldapresults As SearchResultCollection = ldapsearcher.FindAll()
    '            For Each ldapresult As SearchResult In ldapresults
    '                If ct.IsCancellationRequested Then Return New ObservableCollection(Of clsDeletedDirectoryObject)
    '                results.Add(New clsDeletedDirectoryObject(ldapresult, domain))
    '            Next
    '            ldapresults.Dispose()
    '        Next

    '    Catch ex As Exception
    '        ThrowException(ex, domain.Name)
    '    End Try

    '    Return results
    'End Function

End Class
