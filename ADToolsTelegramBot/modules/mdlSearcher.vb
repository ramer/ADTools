'Imports System.DirectoryServices

'Module mdlSearcher

'    Public Function Search(filter As clsFilter, Optional domain As clsDomain = Nothing) As List(Of clsDirectoryObject)
'        Dim results As New List(Of clsDirectoryObject)

'        If domain IsNot Nothing AndAlso domain.Validated Then

'            For Each obj In BasicSearchSync(New clsDirectoryObject(domain.DefaultNamingContext, domain), filter)
'                results.Add(obj)
'            Next

'        Else

'            For Each dmn In domains
'                If Not dmn.Validated Then Continue For

'                For Each obj In BasicSearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), filter)
'                    results.Add(obj)
'                Next
'            Next

'        End If

'        Return results
'    End Function

'    Public Function BasicSearchSync(Optional root As clsDirectoryObject = Nothing,
'                                    Optional filter As clsFilter = Nothing,
'                                    Optional searchscope As SearchScope = SearchScope.Subtree) As List(Of clsDirectoryObject)

'        If root Is Nothing Then Return New List(Of clsDirectoryObject)

'        Dim results As New List(Of clsDirectoryObject)

'        Try
'            Dim ldapsearcher As New DirectorySearcher(root.Entry)
'            ldapsearcher.SearchScope = searchscope
'            ldapsearcher.PropertiesToLoad.Add("objectGuid")
'            ldapsearcher.SizeLimit = 10

'            If searchscope = SearchScope.Subtree Then
'                If filter IsNot Nothing AndAlso Not String.IsNullOrEmpty(filter.Filter) Then ldapsearcher.Filter = filter.Filter
'            End If

'            Dim ldapresults As SearchResultCollection = ldapsearcher.FindAll()

'            For Each ldapresult As SearchResult In ldapresults
'                Dim de As DirectoryEntry = ldapresult.GetDirectoryEntry
'                de.RefreshCache(propertiesToLoadDefault)
'                results.Add(New clsDirectoryObject(de, root.Domain))
'            Next
'            ldapresults.Dispose()

'        Catch ex As Exception
'            ThrowException(ex, "BasicSearchSync")
'        End Try

'        Return results
'    End Function

'    Public Function SearchGUID(guid As Guid) As clsDirectoryObject
'        For Each dmn In domains
'            If Not dmn.Validated Then Continue For

'            Try
'                Dim de As New DirectoryEntry("LDAP://" & dmn.Name & "/<GUID=" & guid.ToString & ">", dmn.Username, dmn.Password)
'                de.RefreshCache(propertiesToLoadDefault)
'                Return New clsDirectoryObject(de, dmn)
'            Catch ex As DirectoryServicesCOMException
'                Continue For
'            End Try
'        Next

'        Return Nothing
'    End Function

'End Module
