Imports System.Collections.ObjectModel
Imports System.DirectoryServices
Imports System.Security.Principal
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks

Public Class clsSearcher
    Private basicsearchtasks As New List(Of Task)
    Private basicsearchtaskscts As New CancellationTokenSource
    Public Event BasicSearchAsyncDataRecieved()

    Private tombstonesearchtasks As New List(Of Task)
    Private tombstonesearchtaskscts As New CancellationTokenSource

    Sub New()

    End Sub

    Public Async Function BasicSearchAsync(returncollection As clsThreadSafeObservableCollection(Of clsDirectoryObject), pattern As String,
                                      Optional specificdomains As ObservableCollection(Of clsDomain) = Nothing,
                                      Optional attributes As ObservableCollection(Of clsAttribute) = Nothing,
                                      Optional searchobjectclasses As clsSearchObjectClasses = Nothing,
                                      Optional freesearch As Boolean = False) As Task

        basicsearchtaskscts.Cancel()
        If basicsearchtasks.Count > 0 Then
            Try
                Await Task.WhenAll(basicsearchtasks.ToArray) 'magic
            Catch
            End Try

            basicsearchtasks.Clear()
            Log("Список задач очищен")
        End If
        returncollection.Clear()

        basicsearchtaskscts = New CancellationTokenSource

        Dim domainlist As ObservableCollection(Of clsDomain) = If(specificdomains Is Nothing, domains, specificdomains)

        For Each dmn In domainlist

            Dim mt = Task.Factory.StartNew(
                Function()
                    Return BasicSearchSync(pattern, dmn, attributes, searchobjectclasses, freesearch, basicsearchtaskscts.Token)
                End Function, basicsearchtaskscts.Token)
            basicsearchtasks.Add(mt)
            Log(String.Format("Задача на поиск в домене ""{0}"" создана", dmn.Name))

            Dim lt = mt.ContinueWith(
                Function(parenttask As Task(Of ObservableCollection(Of clsDirectoryObject)))
                    basicsearchtasks.Remove(mt)
                    Log(String.Format("Вывод {0} результатов в домене ""{1}""", parenttask.Result.Count, If(parenttask.Result.Count > 0, parenttask.Result(0).Domain.Name, "null")))
                    For Each result In parenttask.Result
                        If basicsearchtaskscts.Token.IsCancellationRequested Then Return False
                        returncollection.Add(result)
                        Thread.Sleep(30)
                    Next
                    Log("Задача на поиск закрыта")
                    Return True
                End Function)
            basicsearchtasks.Add(lt)
            Log(String.Format("Задача на вывод результатов домена ""{0}"" создана", dmn.Name))

            Dim ft = lt.ContinueWith(
                Function(parenttask As Task(Of Boolean))
                    Try
                        basicsearchtasks.Remove(lt)
                    Catch
                        basicsearchtasks.Clear()
                    End Try
                    Log("Задача на вывод результатов закрыта")
                    Return True
                End Function)

        Next

        Await Task.WhenAny(basicsearchtasks.ToArray)
        RaiseEvent BasicSearchAsyncDataRecieved()
        Await Task.WhenAll(basicsearchtasks.ToArray)
    End Function

    Public Function BasicSearchSync(pattern As String,
                                      domain As clsDomain,
                                      Optional attributes As ObservableCollection(Of clsAttribute) = Nothing,
                                      Optional searchobjectclasses As clsSearchObjectClasses = Nothing,
                                      Optional freesearch As Boolean = False,
                                      Optional ct As CancellationToken = Nothing) As ObservableCollection(Of clsDirectoryObject)

        Log(String.Format("Поиск ""{0}"" в домене ""{1}""", pattern, domain.Name))

        Dim properties As String() = {"objectGUID",
                                        "userAccountControl",
                                        "accountExpires",
                                        "name",
                                        "description",
                                        "userPrincipalName",
                                        "distinguishedName",
                                        "telephoneNumber",
                                        "physicalDeliveryOfficeName",
                                        "title",
                                        "department",
                                        "company",
                                        "mail",
                                        "whenCreated",
                                        "lastLogon",
                                        "pwdLastSet",
                                        "thumbnailPhoto",
                                        "memberOf",
                                        "givenName",
                                        "sn",
                                        "initials",
                                        "displayName",
                                        "manager",
                                        "sAMAccountName",
                                        "groupType",
                                        "dNSHostName",
                                        "location",
                                        "operatingSystem",
                                        "operatingSystemVersion"}

        If domain.Validated = True Then 'check domain validation
            Log(String.Format("Проверка валидации к домена ""{0}"" завершена", domain.Name))
        Else
            Log(String.Format("Ошибка валидации к домена ""{0}""", domain.Name))
            Return New ObservableCollection(Of clsDirectoryObject)
        End If

        Try ' check LDAP connection
            Dim o As Object = domain.SearchRoot.NativeObject
            Log(String.Format("Проверка подключения к домену ""{0}"" завершена", domain.Name))
        Catch ex As Exception
            Log(String.Format("Отсутствует подключение к домену ""{0}""", domain.Name))
            Return New ObservableCollection(Of clsDirectoryObject)
        End Try

        Dim results As New ObservableCollection(Of clsDirectoryObject)

        Try
            Dim patterns() As String = pattern.Split({"/", vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each singlepattern As String In patterns
                If ct.IsCancellationRequested Then Return New ObservableCollection(Of clsDirectoryObject)

                singlepattern = Trim(singlepattern)
                If String.IsNullOrEmpty(singlepattern) Then Continue For

                Dim _sid As SecurityIdentifier = Nothing
                Dim IsSid As Boolean = False
                Try
                    If Regex.IsMatch(singlepattern, "^S-\d-\d+-(\d+-){1,14}\d+$") Then
                        _sid = New SecurityIdentifier(singlepattern)
                        IsSid = _sid.IsAccountSid
                    End If
                Catch
                End Try

                Dim _guid As Guid = Nothing
                Dim IsGuid As Boolean = Guid.TryParse(singlepattern, _guid)

                If IsSid Then
                    Try
                        Dim de As DirectoryEntry = New DirectoryEntry("LDAP://" & domain.Name & "/<SID=" & _sid.Value & ">", domain.Username, domain.Password)
                        de.RefreshCache(properties)
                        results.Add(New clsDirectoryObject(de, domain))
                        Continue For
                    Catch ex As Exception
                        Continue For
                    End Try
                End If

                If IsGuid Then
                    Try
                        Dim de As New DirectoryEntry("LDAP://" & domain.Name & "/<GUID=" & _guid.ToString & ">", domain.Username, domain.Password)
                        de.RefreshCache(properties)
                        results.Add(New clsDirectoryObject(de, domain))
                        Continue For
                    Catch ex As DirectoryServicesCOMException
                        Continue For
                    End Try
                End If

                attributes = If(attributes, attributesForSearchDefault)
                searchobjectclasses = If(searchobjectclasses, New clsSearchObjectClasses)

                If singlepattern.StartsWith("""") And singlepattern.EndsWith("""") And Len(singlepattern) > 2 Then
                    singlepattern = Mid(singlepattern, 2, Len(singlepattern) - 2)
                Else
                    singlepattern = If(freesearch, "*" & singlepattern & "*", singlepattern & "*")
                End If

                Dim filter As String = "(&" +
                                            "(|" +
                                                If(searchobjectclasses.User, "(&(objectCategory=person)(!(objectClass=inetOrgPerson)))", "") +
                                                If(searchobjectclasses.Computer, "(objectCategory=computer)", "") +
                                                If(searchobjectclasses.Group, "(objectCategory=group)", "") +
                                                If(searchobjectclasses.Container, "(objectCategory=container)", "") +
                                            ")" +
                                            "(|" +
                                                "(" & Join(attributes.Select(Function(x As clsAttribute) x.Name).ToArray, "=" & singlepattern & ")(") & "=" & singlepattern & ")" +
                                            ")" +
                                         ")"

                Dim ldapsearcher As New DirectorySearcher(domain.SearchRoot)
                ldapsearcher.PropertiesToLoad.Add("objectGuid")
                ldapsearcher.Filter = filter
                ldapsearcher.PageSize = 1000

                Dim ldapresults As SearchResultCollection = ldapsearcher.FindAll()

                Log(String.Format("Поиск в домене ""{0}"" запущен. Найдено {1} объектов", domain.Name, ldapresults.Count))

                For Each ldapresult As SearchResult In ldapresults
                    If ct.IsCancellationRequested Then Return New ObservableCollection(Of clsDirectoryObject)

                    Dim de As DirectoryEntry = ldapresult.GetDirectoryEntry
                    de.RefreshCache(properties)
                    results.Add(New clsDirectoryObject(de, domain))
                Next
                ldapresults.Dispose()
            Next

        Catch ex As Exception
            ThrowException(ex, domain.Name)
        End Try

        Log(String.Format("Поиск в домене ""{0}"" завершен. Обработано {1} объектов", domain.Name, results.Count))

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
