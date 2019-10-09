Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.DirectoryServices
Imports System.DirectoryServices.Protocols
Imports System.Security.Principal
Imports IRegisty

<DebuggerDisplay("clsDirectoryObject={name}")>
Public Class clsDirectoryObject
    Inherits Dynamic.DynamicObject
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Private cache As New Dictionary(Of String, DirectoryAttribute)
    Private WithEvents searcher As New clsSearcher

    Sub New()

    End Sub

    Sub New(DistinguishedName As String, ByRef Domain As clsDomain, Optional AttributeNames As String() = Nothing)
        _distinguishedName = DistinguishedName
        _domain = Domain

        If AttributeNames Is Nothing Then AttributeNames = attributesToLoadDefault
        Dim searchRequest = New SearchRequest(DistinguishedName, "(objectClass=*)", Protocols.SearchScope.Base, AttributeNames)
        searchRequest.Controls.Add(New ShowDeletedControl())

        Dim response As SearchResponse
        Try
            response = Connection.SendRequest(searchRequest)
        Catch ex As Exception
            Exit Sub
        End Try

        cache.Clear()

        If response.Entries.Count = 1 Then
            For Each a As DirectoryAttribute In response.Entries(0).Attributes.Values
                cache.Add(a.Name, a)
            Next
        End If
    End Sub

    Sub New(Entry As SearchResultEntry, ByRef Domain As clsDomain, Optional InitialCache As Dictionary(Of String, DirectoryAttribute) = Nothing)
        _distinguishedName = Entry.DistinguishedName
        _domain = Domain

        cache.Clear()

        For Each a As DirectoryAttribute In Entry.Attributes.Values
            cache.Add(a.Name, a)
        Next

        If InitialCache IsNot Nothing Then
            For Each a As String In InitialCache.Keys
                If Not cache.ContainsKey(a) Then cache.Add(a, InitialCache(a))
            Next
        End If
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse TypeOf obj IsNot clsDirectoryObject Then Return False
        Dim target = CType(obj, clsDirectoryObject)
        Return objectGUID.Equals(target.objectGUID)
    End Function

    <RegistrySerializerAfterDeserialize(True)>
    Public Sub AfterDeserialize()
        For Each d In domains
            If d.Name = _domainname Then
                _domain = d

                Dim searchRequest = New SearchRequest(distinguishedName, "(objectClass=*)", Protocols.SearchScope.Base, attributesToLoadDefault)
                searchRequest.Controls.Add(New ShowDeletedControl())

                Dim response As SearchResponse
                Try
                    response = Connection.SendRequest(searchRequest)
                Catch ex As Exception
                    Exit Sub
                End Try

                cache.Clear()

                If response.Entries.Count = 1 Then
                    For Each a As DirectoryAttribute In response.Entries(0).Attributes.Values
                        cache.Add(a.Name, a)
                    Next
                End If

                Exit For
            End If
        Next
    End Sub

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property Connection() As LdapConnection
        Get
            Return If(Domain IsNot Nothing, Domain.Connection, Nothing)
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property Exist() As Boolean
        Get
            Try
                Dim searchRequest = New SearchRequest(distinguishedName, "(objectClass=*)", Protocols.SearchScope.Base, {"Objectclass"})
                searchRequest.Controls.Add(New ShowDeletedControl())
                Connection.SendRequest(searchRequest)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property

    Private _distinguishedName As String
    <RegistrySerializerAlias("DistinguishedName")>
    Public Property distinguishedName() As String
        Get
            Return _distinguishedName
        End Get
        Set(value As String)
            _distinguishedName = value
            NotifyPropertyChanged("distinguishedName")
        End Set
    End Property

    <ExtendedProperty>
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property distinguishedNameFormated() As String
        Get
            Try
                Return Join(distinguishedName.Split({",", "CN=", "OU="}, StringSplitOptions.RemoveEmptyEntries).Reverse.Where(Function(x) Not x.StartsWith("DC=")).ToArray, " \ ")
            Catch
                Return ""
            End Try
        End Get
    End Property

    Private WithEvents _domain As clsDomain
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property Domain() As clsDomain
        Get
            Return _domain
        End Get
    End Property

    Private _domainname As String
    <ExtendedProperty>
    Public Property DomainName() As String
        Get
            Return If(_domain IsNot Nothing, _domain.Name, _domainname)
        End Get
        Set(value As String)
            _domainname = value
            NotifyPropertyChanged("DomainName")
        End Set
    End Property

    Public Sub Refresh(Optional attributenames As String() = Nothing)

        If attributenames Is Nothing Then
            Debug.Print("Attributes refresh requested")
        ElseIf attributenames.Count = 1 Then
            Debug.Print("Attribute ""{0}"" refresh requested", attributenames(0))
        ElseIf attributenames.Count > 1 Then
            Debug.Print("Bulk attributes refresh requested")
        End If

        Dim searchRequest = New SearchRequest(distinguishedName, "(objectClass=*)", Protocols.SearchScope.Base, attributenames)
        searchRequest.Controls.Add(New ShowDeletedControl())

        Dim response As SearchResponse
        Try
            response = Connection.SendRequest(searchRequest)
        Catch ex As Exception
            Exit Sub
        End Try

        ' only one object expercted
        If response.Entries.Count = 1 Then
            For Each a As DirectoryAttribute In response.Entries(0).Attributes.Values
                If a.Name.Contains(";range=") Then
                    Dim rangeparams() = a.Name.Split({";range="}, StringSplitOptions.RemoveEmptyEntries)
                    If rangeparams.Count = 2 Then
                        Dim rangelimits() = rangeparams(1).Split({"-"}, StringSplitOptions.RemoveEmptyEntries)
                        If rangelimits.Count = 2 Then
                            RefreshRanged(rangeparams(0), 1 + Integer.Parse(rangelimits(1)))
                        End If
                    End If
                    Continue For
                End If

                If cache.ContainsKey(a.Name) Then
                    cache(a.Name) = a
                Else
                    cache.Add(a.Name, a)
                End If
            Next
        End If
    End Sub

    Private Sub RefreshRanged(attributename As String, pagesize As Integer)
        Dim page As Integer = 0
        Do

            Dim searchRequest = New SearchRequest(distinguishedName, "(objectClass=*)", Protocols.SearchScope.Base, {attributename & ";range=" & page * pagesize & "-" & (page + 1) * pagesize - 1})
            searchRequest.Controls.Add(New ShowDeletedControl())

            Dim response As SearchResponse
            Try
                response = Connection.SendRequest(searchRequest)
            Catch ex As Exception
                Exit Sub
            End Try

            ' only one object expercted
            If response.Entries.Count = 1 Then
                For Each a As DirectoryAttribute In response.Entries(0).Attributes.Values
                    If cache.ContainsKey(a.Name) Then
                        cache(a.Name) = a
                    Else
                        cache.Add(a.Name, a)
                    End If

                    If a.Count = 0 Or a.Count < pagesize Then Exit Do
                Next
            End If

            page += 1
        Loop

    End Sub

    Public Sub RefreshAllAllowedAttributes()
        If AllowedAttributes Is Nothing Then Exit Sub

        Debug.Print("Bulk allowed attributes refresh requested")

        Refresh(AllowedAttributes.ToArray)
    End Sub

    Public Function GetValue(attributename As String) As Object
        Dim attr As clsAttribute = GetAttribute(attributename)
        If attr Is Nothing Then
            Return Nothing
        Else
            Return attr.Value
        End If
    End Function


    Public Function GetAttribute(attributename As String) As clsAttribute
        If Not cache.ContainsKey(attributename) Then Refresh({attributename}) ' refresh entry if requested attribute not found
        If Not cache.ContainsKey(attributename) Then cache.Add(attributename, Nothing) 'if entry hasn't value store nothing 

        Dim baseattr As clsAttribute = If(Domain.AttributesSchema.ContainsKey(attributename), Domain.AttributesSchema(attributename), Nothing)
        If baseattr Is Nothing Then Return Nothing

        Dim attr As clsAttribute
        Select Case baseattr.attributeSyntax
            Case "2.5.5.8" 'Boolean. For queries that include attributes of Boolean syntax in a filter, specify TRUE or FALSE (for example, myboolattr=TRUE).
                attr = New clsAttributeBoolean(baseattr, cache(attributename))
            Case "2.5.5.1" 'String that contains a DN. For attributes with this syntax, Active Directory handles attribute values as references to the object identified by the DN and automatically updates the value if the object is moved or renamed. For queries that include attributes of DN syntax in a filter, specify full distinguished names—wildcards (for example, cn=John*) are not supported.
                attr = New clsAttributeDistinguishedName(baseattr, cache(attributename))
            Case "2.5.5.2" 'An OID string, which is a string that contains digits (0-9) and decimal points (.).
                attr = New clsAttributeObjectIdentifier(baseattr, cache(attributename))
            Case "2.5.5.4" 'A case-insensitive string that contains characters from the teletex character set.
                attr = New clsAttributeCaseInsensitiveString(baseattr, cache(attributename))
            Case "2.5.5.5" 'A case-sensitive string that contains characters from the printable character set.
                attr = New clsAttributeIA5String(baseattr, cache(attributename))
            Case "2.5.5.6" 'String that contains digits.
                attr = New clsAttributeNumericString(baseattr, cache(attributename))
            Case "2.5.5.11" 'A case-sensitive character string, with UTC-Time.
                attr = New clsAttributeUTCTime(baseattr, cache(attributename))
            Case "2.5.5.12" 'A case-insensitive Unicode string.
                attr = New clsAttributeUnicodeString(baseattr, cache(attributename))
            Case "2.5.5.13" 'A string that contains OSI presentation addresses.
                attr = New clsAttributeOSIString(baseattr, cache(attributename))
            Case "2.5.5.14" 'From X400.
                attr = New clsAttributeDistinguishedNameWithString(baseattr, cache(attributename))
            Case "2.5.5.15" 'Octet string that contains a Windows NT/Windows 2000 security descriptor.
                attr = New clsAttributeNTSecurityDescriptor(baseattr, cache(attributename))
            Case "2.5.5.17" 'Octet string that contains a security identifier (SID). Use this syntax to store SID values only.
                attr = New clsAttributeSID(baseattr, cache(attributename))
            Case "2.5.5.7" 'An OctetString that contains a binary value and a DN. A value with this syntax has the following format: B : CharCount : binaryvalue : ObjectDN
                attr = New clsAttributeDNBinary(baseattr, cache(attributename))
            Case "2.5.5.10" 'Array of bytes. Use OctetString to store binary data.
                attr = New clsAttributeOctetString(baseattr, cache(attributename))
            Case "2.5.5.9" '32-bit integer.
                attr = New clsAttributeInteger(baseattr, cache(attributename))
            Case "2.5.5.16" 'Large integer. Use for 64-bit values.
                attr = New clsAttributeLong(baseattr, cache(attributename))
            Case Else
                Throw New ArgumentException
        End Select

        Return attr
    End Function

    Public Sub SetAttribute(attributename As String, value As Object)
        Exit Sub
        If value Is Nothing OrElse String.IsNullOrEmpty(value.ToString) Then

            Try

                Dim modifyRequest As New ModifyRequest(distinguishedName, DirectoryAttributeOperation.Delete, attributename, Nothing)
                Dim response As ModifyResponse = Connection.SendRequest(modifyRequest)
                Refresh({attributename})
                If cache.ContainsKey(attributename) Then cache.Remove(attributename)
            Catch e As DirectoryOperationException

            End Try

        Else

            Try

                If TypeOf value IsNot Byte() Then value = value.ToString
                Dim modifyRequest As New ModifyRequest(distinguishedName, DirectoryAttributeOperation.Replace, attributename, value)
                Dim response As ModifyResponse = Connection.SendRequest(modifyRequest)
                Refresh({attributename})

            Catch generatedExceptionName As DirectoryOperationException

                Try

                    If TypeOf value IsNot Byte() Then value = value.ToString
                    Dim modifyRequest As New ModifyRequest(distinguishedName, DirectoryAttributeOperation.Add, attributename, value)
                    Dim response As ModifyResponse = Connection.SendRequest(modifyRequest)
                    Refresh({attributename})

                Catch e As DirectoryOperationException

                End Try

            Catch e As Exception

            End Try

        End If
    End Sub

    Public Sub UpdateAttribute(operation As DirectoryAttributeOperation, attributename As String, value As Object)
        Try

            Dim modifyRequest As New ModifyRequest(distinguishedName, operation, attributename, value)
            Dim response As ModifyResponse = Connection.SendRequest(modifyRequest)
            Refresh({attributename})

        Catch e As Exception
            Throw e
        End Try
    End Sub

    Public Sub MoveTo(destinationDN As String)
        Try
            Dim newName As String = distinguishedName.Split({","}, StringSplitOptions.RemoveEmptyEntries).First
            Dim modifyDnRequest As New ModifyDNRequest(distinguishedName, destinationDN, newName)
            Dim modifyDnResponse As ModifyDNResponse = DirectCast(Connection.SendRequest(modifyDnRequest), ModifyDNResponse)
            _distinguishedName = newName & "," & destinationDN
            NotifyPropertyChanged("distinguishedName")
        Catch e As Exception
            Throw e
        End Try
    End Sub

    Public Sub Rename(newName As String)
        Try
            Dim eDN As List(Of String) = distinguishedName.Split({","}, StringSplitOptions.RemoveEmptyEntries).ToList
            If eDN.Count <= 1 Then Exit Sub
            eDN.RemoveAt(0)
            Dim newParent As String = Join(eDN.ToArray, ",")

            Dim modifyDnRequest As New ModifyDNRequest(distinguishedName, newParent, newName)
            Dim modifyDnResponse As ModifyDNResponse = DirectCast(Connection.SendRequest(modifyDnRequest), ModifyDNResponse)
            _distinguishedName = newName & "," & newParent
            NotifyPropertyChanged("distinguishedName")
        Catch e As Exception
            Throw e
        End Try
    End Sub

    Public Sub DeleteTree()
        Try
            Dim deleteRequest As New DeleteRequest(distinguishedName)
            Dim treeDeleteControl As New TreeDeleteControl()
            deleteRequest.Controls.Add(treeDeleteControl)
            Dim deleteResponse As DeleteResponse = DirectCast(Connection.SendRequest(deleteRequest), DeleteResponse)
        Catch e As Exception
            Throw e
        End Try
    End Sub

    Public Sub Delete()
        Try
            Dim deleteRequest As New DeleteRequest(distinguishedName)
            Dim deleteResponse As DeleteResponse = DirectCast(Connection.SendRequest(deleteRequest), DeleteResponse)
        Catch e As Exception
            Throw e
        End Try
    End Sub

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property Parent As clsDirectoryObject
        Get
            Dim eDN As List(Of String) = distinguishedName.Split({","}, StringSplitOptions.RemoveEmptyEntries).ToList
            If eDN.Count <= 1 Then Return Nothing
            eDN.RemoveAt(0)
            Return New clsDirectoryObject(Join(eDN.ToArray, ","), Domain, {"name", "objectClass", "objectCategory"})
        End Get
    End Property

    Private _childcontainers As ObservableCollection(Of clsDirectoryObject)
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property ChildContainers() As ObservableCollection(Of clsDirectoryObject)
        Get
            If IsDeleted Then Return New ObservableCollection(Of clsDirectoryObject)
            If _childcontainers IsNot Nothing Then
                Return _childcontainers
            Else
                _childcontainers = New ObservableCollection(Of clsDirectoryObject)
                GetChildContainersAsync()
            End If

            Return _childcontainers
        End Get
    End Property

    Private Async Sub GetChildContainersAsync()
        _childcontainers = Await Task.Run(
            Function()
                Return searcher.SearchChildContainersSync(
            Me,
            New clsFilter("(|(objectClass=organizationalUnit)(objectClass=container)(objectClass=builtindomain)(objectClass=domaindns)(objectClass=lostandfound))"),, True)
            End Function)
        NotifyPropertyChanged("ChildContainers")
    End Sub

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property CanWrite(name As String) As Boolean
        Get
            Return AllowedAttributesEffective.Contains(name, StringComparer.OrdinalIgnoreCase)
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsReadOnly(name As String) As Boolean
        Get
            Return Not AllowedAttributesEffective.Contains(name, StringComparer.OrdinalIgnoreCase)
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property HasValue(name As String) As Boolean
        Get
            Return GetValue(name) IsNot Nothing
        End Get
    End Property

    Private _allowedattributes As List(Of String)
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property AllowedAttributes() As List(Of String)
        Get
            If _allowedattributes IsNot Nothing Then Return _allowedattributes

            Dim a As Object = GetValue("allowedAttributes")

            If a Is Nothing Then
                _allowedattributes = New List(Of String)
            Else
                _allowedattributes = New List(Of String)(CType(a, String()))
            End If

            Return _allowedattributes
        End Get
    End Property

    Private _allowedattributeseffective As List(Of String)
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property AllowedAttributesEffective() As List(Of String)
        Get
            If _allowedattributeseffective IsNot Nothing Then Return _allowedattributeseffective

            Dim a As Object = GetValue("allowedAttributesEffective")

            If a Is Nothing Then
                _allowedattributeseffective = New List(Of String)
            Else
                _allowedattributeseffective = New List(Of String)(CType(a, String()))
            End If

            Return _allowedattributeseffective
        End Get
    End Property

    '<RegistrySerializerIgnorable(True)>
    'Public ReadOnly Property AllAttributes As ObservableCollection(Of clsAttributeSchema)
    '    Get
    '        Dim aa As New ObservableCollection(Of clsAttributeSchema)
    '        For Each attr As String In AllowedAttributes
    '            If Domain.AttributesSchema.ContainsKey(attr) Then aa.Add(New clsAttributeSchema(Domain.AttributesSchema(attr)))
    '        Next
    '        Return aa
    '    End Get
    'End Property

    Public Overrides Function TryGetMember(ByVal binder As Dynamic.GetMemberBinder, ByRef result As Object) As Boolean
        result = GetValue(binder.Name)
        Return True
    End Function

    Public Overrides Function TrySetMember(ByVal binder As Dynamic.SetMemberBinder, ByVal value As Object) As Boolean
        SetAttribute(binder.Name, value)
        NotifyPropertyChanged(binder.Name)
        Return True
    End Function

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property SchemaClass() As enmDirectoryObjectSchemaClass
        Get
            Dim cat = objectCategory
            Dim cls = objectClass
            If cat Is Nothing OrElse cls Is Nothing Then Return Nothing

            If cat = "person" And cls.Contains("user") Then
                Return enmDirectoryObjectSchemaClass.User
            ElseIf cat = "person" And cls.Contains("contact", StringComparer.OrdinalIgnoreCase) Then
                Return enmDirectoryObjectSchemaClass.Contact
            ElseIf cls.Contains("computer", StringComparer.OrdinalIgnoreCase) Then
                Return enmDirectoryObjectSchemaClass.Computer
            ElseIf cls.Contains("group", StringComparer.OrdinalIgnoreCase) Then
                Return enmDirectoryObjectSchemaClass.Group
            ElseIf cls.Contains("organizationalunit", StringComparer.OrdinalIgnoreCase) Then
                Return enmDirectoryObjectSchemaClass.OrganizationalUnit
            ElseIf cls.Contains("container", StringComparer.OrdinalIgnoreCase) Or cls.Contains("builtindomain", StringComparer.OrdinalIgnoreCase) Then
                Return enmDirectoryObjectSchemaClass.Container
            ElseIf cls.Contains("domaindns", StringComparer.OrdinalIgnoreCase) Then
                Return enmDirectoryObjectSchemaClass.DomainDNS
            ElseIf cls.Contains("lostandfound", StringComparer.OrdinalIgnoreCase) Then
                Return enmDirectoryObjectSchemaClass.UnknownContainer
            Else
                Return enmDirectoryObjectSchemaClass.Unknown
            End If
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsUser
        Get
            Return SchemaClass = enmDirectoryObjectSchemaClass.User
        End Get
    End Property
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsContact
        Get
            Return SchemaClass = enmDirectoryObjectSchemaClass.Contact
        End Get
    End Property
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsComputer
        Get
            Return SchemaClass = enmDirectoryObjectSchemaClass.Computer
        End Get
    End Property
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsGroup
        Get
            Return SchemaClass = enmDirectoryObjectSchemaClass.Group
        End Get
    End Property
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsDeleted As Boolean?
        Get
            Return GetValue("isDeleted")
        End Get
    End Property
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsRecycled As Boolean?
        Get
            Return GetValue("isRecycled")
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property Image As Grid
        Get
            Dim grd As New Grid With {.ToolTip = If(String.IsNullOrEmpty(StatusFormatted), Nothing, StatusFormatted), .SnapsToDevicePixels = True}
            Dim imggrd As New Grid With {.Margin = New Thickness(2)}
            grd.Children.Add(imggrd)

            Dim img As New Image With {.Stretch = Stretch.Uniform, .StretchDirection = StretchDirection.Both, .HorizontalAlignment = HorizontalAlignment.Center, .VerticalAlignment = VerticalAlignment.Center}

            If SchemaClass = enmDirectoryObjectSchemaClass.User AndAlso thumbnailPhoto IsNot Nothing Then
                Dim clipbrdp As New Border With {.Background = Brushes.White, .CornerRadius = New CornerRadius(999)}
                img.Source = thumbnailPhoto
                imggrd.Children.Add(clipbrdp)
                imggrd.Children.Add(img)
                imggrd.OpacityMask = New VisualBrush() With {.Visual = clipbrdp}

                Dim strokebrdr As New Border With {.BorderThickness = New Thickness(3), .CornerRadius = New CornerRadius(999), .BorderBrush = If(Status = enmDirectoryObjectStatus.Normal, New SolidColorBrush(Color.FromRgb(74, 217, 65)), If(Status = enmDirectoryObjectStatus.Expired, New SolidColorBrush(Color.FromRgb(246, 204, 33)), New SolidColorBrush(Color.FromRgb(228, 71, 71))))}
                grd.Children.Add(strokebrdr)
            Else
                img.Source = StatusImage
                imggrd.Children.Add(img)
            End If

            Return grd
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property Status() As enmDirectoryObjectStatus
        Get
            Dim _status As enmDirectoryObjectStatus = enmDirectoryObjectStatus.Normal

            If SchemaClass = enmDirectoryObjectSchemaClass.User Or SchemaClass = enmDirectoryObjectSchemaClass.Computer Then
                If passwordNeverExpires Is Nothing Then
                    _status = enmDirectoryObjectStatus.Expired
                ElseIf passwordNeverExpires = False Then
                    If passwordExpiresDate() = Nothing Then
                        _status = enmDirectoryObjectStatus.Expired
                    ElseIf passwordExpiresDate() <= Now Then
                        _status = enmDirectoryObjectStatus.Expired
                    End If
                End If

                If accountNeverExpires Is Nothing Then
                    _status = enmDirectoryObjectStatus.Expired
                ElseIf accountNeverExpires = False AndAlso accountExpiresDate <= Now Then
                    _status = enmDirectoryObjectStatus.Expired
                End If

                If disabled Is Nothing Then
                    _status = enmDirectoryObjectStatus.Expired
                ElseIf disabled Then
                    _status = enmDirectoryObjectStatus.Blocked
                End If
            End If

            Return _status
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property StatusFormatted() As String
        Get
            Dim _status As String = ""

            If SchemaClass = enmDirectoryObjectSchemaClass.User OrElse SchemaClass = enmDirectoryObjectSchemaClass.Computer Then
                If passwordNeverExpires Is Nothing Then
                    _status &= My.Resources.str_PasswordExpirationUnknown & vbCr
                ElseIf passwordNeverExpires = False Then
                    If passwordExpiresDate() = Nothing Then
                        _status &= My.Resources.str_PasswordExpirationUnknown & vbCr
                    ElseIf passwordExpiresDate() <= Now Then
                        _status &= My.Resources.str_PasswordExpired & vbCr
                    End If
                End If

                If accountNeverExpires Is Nothing Then
                    _status &= My.Resources.str_ObjectExpirationUnknown & vbCr
                ElseIf accountNeverExpires = False AndAlso accountExpiresDate <= Now Then
                    _status &= My.Resources.str_ObjectExpired & vbCr
                End If

                If disabled Is Nothing Then
                    _status &= My.Resources.str_ObjectStatusUnknown & vbCr
                ElseIf disabled Then
                    _status &= My.Resources.str_ObjectDisabled & vbCr
                End If
                If _status.Length > 1 Then _status = _status.Remove(_status.Length - 1)
            End If

            Return _status
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property StatusImage() As BitmapImage
        Get
            Dim _image As String = ""

            If SchemaClass = enmDirectoryObjectSchemaClass.User Then
                _image = "user.png"
                If passwordNeverExpires Is Nothing Then
                    _image = "user_unknown.png"
                ElseIf passwordNeverExpires = False Then
                    If passwordExpiresDate = Nothing Then
                        _image = "user_expired.png"
                    ElseIf passwordExpiresDate() <= Now Then
                        _image = "user_expired.png"
                    End If
                End If

                If accountNeverExpires Is Nothing Then
                    _image = "user_unknown.png"
                ElseIf accountNeverExpires = False AndAlso accountExpiresDate <= Now Then
                    _image = "user_expired.png"
                End If

                If disabled Is Nothing Then
                    _image = "user_unknown.png"
                ElseIf disabled Then
                    _image = "user_blocked.png"
                End If
            ElseIf SchemaClass = enmDirectoryObjectSchemaClass.Computer Then
                _image = "computer.png"
                If passwordNeverExpires Is Nothing Then
                    _image = "computer_unknown.png"
                ElseIf passwordNeverExpires = False Then
                    If passwordExpiresDate = Nothing Then
                        _image = "computer_expired.png"
                    ElseIf passwordExpiresDate() <= Now Then
                        _image = "computer_expired.png"
                    End If
                End If

                If accountNeverExpires Is Nothing Then
                    _image = "computer_unknown.png"
                ElseIf accountNeverExpires = False AndAlso accountExpiresDate <= Now Then
                    _image = "computer_expired.png"
                End If

                If disabled Is Nothing Then
                    _image = "computer_unknown.png"
                ElseIf disabled Then
                    _image = "computer_blocked.png"
                End If
            ElseIf SchemaClass = enmDirectoryObjectSchemaClass.Group Then
                _image = "group.png"
                If groupTypeSecurity Then
                    _image = "group.png"
                ElseIf groupTypeDistribution Then
                    _image = "group_distribution.png"
                Else
                    _image = "puzzle.png"
                End If
            Else
                Return ClassImage
            End If

            Dim bi As New BitmapImage(New Uri("pack://application:,,,/images/" & _image))

            Return bi
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property ClassImage() As BitmapImage
        Get
            Dim _image As String = ""

            Select Case SchemaClass
                Case enmDirectoryObjectSchemaClass.User
                    _image = "user_unknown.png"
                Case enmDirectoryObjectSchemaClass.Contact
                    _image = "contact.png"
                Case enmDirectoryObjectSchemaClass.Computer
                    _image = "computer_unknown.png"
                Case enmDirectoryObjectSchemaClass.Group
                    _image = "group.png"
                Case enmDirectoryObjectSchemaClass.OrganizationalUnit
                    _image = "folder.png"
                Case enmDirectoryObjectSchemaClass.Container
                    _image = "folder_locked.png"
                Case enmDirectoryObjectSchemaClass.DomainDNS
                    _image = "domain.png"
                Case enmDirectoryObjectSchemaClass.UnknownContainer
                    _image = "folder_unknown.png"
                Case enmDirectoryObjectSchemaClass.Unknown
                    _image = "puzzle.png"
                Case Else
                    _image = "puzzle.png"
            End Select

            Dim bi As New BitmapImage(New Uri("pack://application:,,,/images/" & _image))

            Return bi
        End Get
    End Property

    Public Sub ResetPassword()
        If String.IsNullOrEmpty(Domain.DefaultPassword) Then Throw New Exception(My.Resources.str_DefaultPasswordIsNotSet)

        Dim path = "LDAP://" & If(String.IsNullOrEmpty(Domain.Server), Domain.Name, Domain.Server) & "/" & distinguishedName
        Using de As New DirectoryEntry(path, Domain.Username, Domain.Password)
            de.Invoke("SetPassword", Domain.DefaultPassword)
            de.CommitChanges()
        End Using

        pwdLastSet = 0

        If Domain.LogActions = True AndAlso Not String.IsNullOrEmpty(Domain.LogActionsAttribute) Then
            SetAttribute(Domain.LogActionsAttribute, String.Format("{0} {1} ({2})", My.Resources.str_PasswordChanged, Domain.Username, Now.ToShortTimeString & " " & Now.ToShortDateString))
        End If
    End Sub

    Public Sub SetPassword(password As String)
        If String.IsNullOrEmpty(password) Then Exit Sub

        Dim path = "LDAP://" & If(String.IsNullOrEmpty(Domain.Server), Domain.Name, Domain.Server) & "/" & distinguishedName
        Using de As New DirectoryEntry(path, Domain.Username, Domain.Password)
            de.Invoke("SetPassword", password)
            de.CommitChanges()
        End Using

        pwdLastSet = -1

        If Domain.LogActions = True AndAlso Not String.IsNullOrEmpty(Domain.LogActionsAttribute) Then
            SetAttribute(Domain.LogActionsAttribute, String.Format("{0} {1} ({2})", My.Resources.str_PasswordChanged, Domain.Username, Now.ToShortTimeString & " " & Now.ToShortDateString))
        End If
    End Sub


#Region "User attributes"

    <RegistrySerializerIgnorable(True)>
    Public Property sn() As String
        Get
            Return GetValue("sn")
        End Get
        Set(ByVal value As String)
            SetAttribute("sn", value)

            NotifyPropertyChanged("sn")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property initials() As String
        Get
            Return GetValue("initials")
        End Get
        Set(ByVal value As String)
            SetAttribute("initials", value)

            NotifyPropertyChanged("initials")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property givenName() As String
        Get
            Return GetValue("givenName")
        End Get
        Set(ByVal value As String)
            SetAttribute("givenName", value)

            NotifyPropertyChanged("givenName")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property displayName() As String
        Get
            Return GetValue("displayName")
        End Get
        Set(ByVal value As String)
            SetAttribute("displayName", value)

            NotifyPropertyChanged("displayName")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property employeeID() As String
        Get
            Return GetValue("employeeID")
        End Get
        Set(ByVal value As String)
            SetAttribute("employeeID", value)

            NotifyPropertyChanged("employeeID")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property employeeNumber() As String
        Get
            Return GetValue("employeeNumber")
        End Get
        Set(ByVal value As String)
            SetAttribute("employeeNumber", value)

            NotifyPropertyChanged("employeeNumber")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property employeeType() As String
        Get
            Return GetValue("employeeType")
        End Get
        Set(ByVal value As String)
            SetAttribute("employeeType", value)

            NotifyPropertyChanged("employeeType")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property physicalDeliveryOfficeName() As String
        Get
            Return GetValue("physicalDeliveryOfficeName")
        End Get
        Set(ByVal value As String)
            SetAttribute("physicalDeliveryOfficeName", value)

            NotifyPropertyChanged("physicalDeliveryOfficeName")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property company() As String
        Get
            Return GetValue("company")
        End Get
        Set(ByVal value As String)
            SetAttribute("company", value)

            NotifyPropertyChanged("company")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property telephoneNumber() As String
        Get
            Return GetValue("telephoneNumber")
        End Get
        Set(ByVal value As String)
            SetAttribute("telephoneNumber", value)

            NotifyPropertyChanged("telephoneNumber")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property homePhone() As String
        Get
            Return GetValue("homePhone")
        End Get
        Set(ByVal value As String)
            SetAttribute("homePhone", value)

            NotifyPropertyChanged("homePhone")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ipPhone() As String
        Get
            Return GetValue("ipPhone")
        End Get
        Set(ByVal value As String)
            SetAttribute("ipPhone", value)

            NotifyPropertyChanged("ipPhone")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property mobile() As String
        Get
            Return GetValue("mobile")
        End Get
        Set(ByVal value As String)
            SetAttribute("mobile", value)

            NotifyPropertyChanged("mobile")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property streetAddress() As String
        Get
            Return GetValue("streetAddress")
        End Get
        Set(ByVal value As String)
            SetAttribute("streetAddress", value)

            NotifyPropertyChanged("streetAddress")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property mail() As String
        Get
            Return GetValue("mail")
        End Get
        Set(ByVal value As String)
            SetAttribute("mail", value)

            NotifyPropertyChanged("mail")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property title() As String
        Get
            Return GetValue("title")
        End Get
        Set(ByVal value As String)
            SetAttribute("title", value)

            NotifyPropertyChanged("title")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property department() As String
        Get
            Return GetValue("department")
        End Get
        Set(ByVal value As String)
            SetAttribute("department", value)

            NotifyPropertyChanged("department")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property userPrincipalName() As String
        Get
            Return GetValue("userPrincipalName")
        End Get
        Set(ByVal value As String)
            SetAttribute("userPrincipalName", value)

            NotifyPropertyChanged("userPrincipalName")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public Property userPrincipalNameName() As String
        Get
            If userPrincipalName Is Nothing Then Return Nothing
            Dim s As String
            If userPrincipalName.Contains("@") Then
                s = Split(userPrincipalName, "@")(0)
            ElseIf Len(userPrincipalName) > 0 Then
                s = userPrincipalName
            Else
                s = ""
            End If
            Return s
        End Get
        Set(ByVal value As String)
            userPrincipalName = value & "@" & userPrincipalNameDomain
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public Property userPrincipalNameDomain() As String
        Get
            If userPrincipalName Is Nothing Then Return Nothing
            Dim s As String
            If userPrincipalName.Contains("@") Then
                s = Split(userPrincipalName, "@")(1)
            Else
                s = ""
            End If
            Return s
        End Get
        Set(ByVal value As String)
            userPrincipalName = userPrincipalNameName & "@" & value
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property thumbnailPhoto() As BitmapImage
        Get
            Dim photo = GetValue("thumbnailPhoto")

            If photo IsNot Nothing Then
                Using ms = New System.IO.MemoryStream(CType(photo, Byte()))
                    Dim bi = New BitmapImage()
                    bi.BeginInit()
                    bi.CacheOption = BitmapCacheOption.OnLoad
                    bi.StreamSource = ms
                    bi.EndInit()
                    Return bi
                End Using
            Else
                Return Nothing
            End If

        End Get
        Set(value As BitmapImage)

            If value Is Nothing Then
                Try
                    SetAttribute("thumbnailPhoto", Nothing)
                Catch ex As Exception
                    Throw ex
                End Try
            Else
                Try
                    Dim bytes() As Byte = Nothing
                    value.CopyPixels(bytes, 0, 0)
                    SetAttribute("thumbnailPhoto", bytes)
                Catch ex As Exception
                    Throw ex
                End Try
            End If

            NotifyPropertyChanged("thumbnailPhoto")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property userWorkstations() As String()
        Get
            Dim w As String = GetValue("userWorkstations")
            Return If(Not String.IsNullOrEmpty(w), w.Split({","}, StringSplitOptions.RemoveEmptyEntries), New String() {})
        End Get
        Set(value As String())
            SetAttribute("userWorkstations", Join(value, ","))

            NotifyPropertyChanged("userWorkstations")
        End Set
    End Property

#End Region

#Region "Computer attributes"

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property dNSHostName As String
        Get
            Return GetValue("dNSHostName")
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property location() As String
        Get
            Return GetValue("location")
        End Get
        Set(ByVal value As String)
            SetAttribute("location", value)

            NotifyPropertyChanged("location")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property operatingSystem() As String
        Get
            Return GetValue("operatingSystem")
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property operatingSystemVersion() As String
        Get
            Return GetValue("operatingSystemVersion")
        End Get
    End Property

#End Region

#Region "Group attributes"

    <RegistrySerializerIgnorable(True)>
    Public Property groupType() As Long
        Get
            Return GetValue("groupType")
        End Get
        Set(ByVal value As Long)
            Try
                SetAttribute("groupType", value)
            Catch ex As Exception
                ShowWrongMemberMessage()
            End Try

            NotifyPropertyChanged("groupType")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property groupTypeScopeDomainLocal() As Boolean
        Get
            Return groupType And mdlVariables.ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
        End Get
        Set(ByVal value As Boolean)
            If value Then
                Dim gt As Long = groupType
                If Not (gt And mdlVariables.ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP) Then
                    gt = gt + mdlVariables.ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
                End If
                If (gt And mdlVariables.ADS_GROUP_TYPE_GLOBAL_GROUP) Then
                    gt = gt - mdlVariables.ADS_GROUP_TYPE_GLOBAL_GROUP
                End If
                If (gt And mdlVariables.ADS_GROUP_TYPE_UNIVERSAL_GROUP) Then
                    gt = gt - mdlVariables.ADS_GROUP_TYPE_UNIVERSAL_GROUP
                End If
                groupType = gt
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property groupTypeScopeGlobal() As Boolean
        Get
            Return groupType And mdlVariables.ADS_GROUP_TYPE_GLOBAL_GROUP
        End Get
        Set(ByVal value As Boolean)
            If value Then
                Dim gt As Long = groupType
                If (gt And mdlVariables.ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP) Then
                    gt = gt - mdlVariables.ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
                End If
                If Not (gt And mdlVariables.ADS_GROUP_TYPE_GLOBAL_GROUP) Then
                    gt = gt + mdlVariables.ADS_GROUP_TYPE_GLOBAL_GROUP
                End If
                If (gt And mdlVariables.ADS_GROUP_TYPE_UNIVERSAL_GROUP) Then
                    gt = gt - mdlVariables.ADS_GROUP_TYPE_UNIVERSAL_GROUP
                End If
                groupType = gt
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property groupTypeScopeUniversal() As Boolean
        Get
            Return groupType And mdlVariables.ADS_GROUP_TYPE_UNIVERSAL_GROUP
        End Get
        Set(ByVal value As Boolean)
            If value Then
                Dim gt As Long = groupType
                If (gt And mdlVariables.ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP) Then
                    gt = gt - mdlVariables.ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
                End If
                If (gt And mdlVariables.ADS_GROUP_TYPE_GLOBAL_GROUP) Then
                    gt = gt - mdlVariables.ADS_GROUP_TYPE_GLOBAL_GROUP
                End If
                If Not (gt And mdlVariables.ADS_GROUP_TYPE_UNIVERSAL_GROUP) Then
                    gt = gt + mdlVariables.ADS_GROUP_TYPE_UNIVERSAL_GROUP
                End If
                groupType = gt
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property groupTypeSecurity() As Boolean
        Get
            Return groupType And mdlVariables.ADS_GROUP_TYPE_SECURITY_ENABLED
        End Get
        Set(ByVal value As Boolean)
            If value Then
                If Not (groupType And mdlVariables.ADS_GROUP_TYPE_SECURITY_ENABLED) Then
                    groupType = groupType + mdlVariables.ADS_GROUP_TYPE_SECURITY_ENABLED
                End If
            Else
                If (groupType And mdlVariables.ADS_GROUP_TYPE_SECURITY_ENABLED) Then
                    groupType = groupType - mdlVariables.ADS_GROUP_TYPE_SECURITY_ENABLED
                End If
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property groupTypeDistribution() As Boolean
        Get
            Return Not groupTypeSecurity
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property info() As String
        Get
            Return GetValue("info")
        End Get
        Set(ByVal value As String)
            SetAttribute("info", value)

            NotifyPropertyChanged("info")
        End Set
    End Property

#End Region

#Region "Shared attributes"

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property objectClass() As String()
        Get
            Try
                Return GetValue("objectClass")
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property objectCategory() As String
        Get
            Try
                Dim oc = GetValue("objectCategory")
                If oc IsNot Nothing Then
                    Dim ocarr = LCase(oc.ToString).Split(New String() {"=", ","}, StringSplitOptions.RemoveEmptyEntries)
                    Return If(ocarr.Length >= 2, ocarr(1), Nothing)
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property accountExpires() As Long?
        Get
            Return GetValue("accountExpires")
        End Get
        Set(ByVal value As Long?)
            If value IsNot Nothing Then SetAttribute("accountExpires", value)

            NotifyPropertyChanged("accountExpires")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public Property accountExpiresDate() As Date
        Get
            If accountExpires IsNot Nothing Then
                If accountExpires = 0 Or accountExpires = 9223372036854775807 Then
                    Return Nothing
                Else
                    Return Date.FromFileTime(accountExpires)
                End If
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Date)
            accountExpires = value.ToFileTime
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property accountExpiresFormated() As String
        Get
            If accountExpires IsNot Nothing Then
                If accountExpires = 0 Or accountExpires = 9223372036854775807 Then
                    Return My.Resources.str_Never
                Else
                    Return Date.FromFileTime(accountExpires).ToString
                End If
            Else
                Return My.Resources.str_Unknown
            End If
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public Property accountNeverExpires() As Boolean?
        Get
            If accountExpires IsNot Nothing Then
                Return accountExpires = 0 Or accountExpires = 9223372036854775807
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Boolean?)
            If value IsNot Nothing Then
                accountExpires = If(value, 0, Now.ToFileTime)
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property badPwdCount() As Integer?
        Get
            Return GetValue("badPwdCount")
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property description() As String
        Get
            Dim val = GetValue("description")
            Return If(TypeOf val Is String(), val(0), val)
        End Get
        Set(ByVal value As String)
            SetAttribute("description", value)

            NotifyPropertyChanged("description")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property lastLogon() As Long?
        Get
            Return GetValue("lastLogon")
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property lastLogonDate() As Date
        Get
            If lastLogon IsNot Nothing Then
                Return If(lastLogon <= 0, Nothing, Date.FromFileTime(lastLogon))
            Else
                Return Nothing
            End If
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property lastLogonFormated() As String
        Get
            If lastLogon IsNot Nothing Then
                Return If(lastLogon <= 0, My.Resources.str_Never, Date.FromFileTime(lastLogon).ToString)
            Else
                Return My.Resources.str_Unknown
            End If
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property logonCount() As Integer
        Get
            Return GetValue("logonCount")
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property name() As String
        Get
            Return If(GetValue("name"), "Null object")
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property objectGUID() As Guid
        Get
            Dim a = GetValue("objectGUID")
            If a Is Nothing Then Return Nothing
            Return New Guid(TryCast(a, Byte()))
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property objectGUIDFormated() As String
        Get
            Return New Guid(TryCast(GetValue("objectGUID"), Byte())).ToString
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property parentGUID() As Guid
        Get
            Return New Guid(TryCast(GetValue("parentGUID"), Byte()))
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property parentGUIDFormated() As String
        Get
            Return New Guid(TryCast(GetValue("parentGUID"), Byte())).ToString
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property objectSIDFormatted() As String
        Get
            Try
                Dim sid As New SecurityIdentifier(TryCast(GetValue("objectSid"), Byte()), 0)
                If sid.IsAccountSid Then Return sid.ToString
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property pwdLastSet() As Long?
        Get
            Return GetValue("pwdLastSet")
        End Get
        Set(ByVal value As Long?)
            If value IsNot Nothing Then SetAttribute("pwdLastSet", value)

            NotifyPropertyChanged("pwdLastSet")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property pwdLastSetDate() As Date
        Get
            Dim pls = pwdLastSet
            Return If(pls IsNot Nothing AndAlso pls > 0, Date.FromFileTime(pls), Nothing)
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property pwdLastSetFormated() As String
        Get
            Dim pls = pwdLastSet
            Return If(pls Is Nothing, My.Resources.str_Unknown, If(pls = 0, My.Resources.str_Expired, Date.FromFileTime(pls).ToString))
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property passwordExpiresDate() As Date
        Get
            Dim pne = passwordNeverExpires
            If pne IsNot Nothing AndAlso pne Then
                Return Nothing
            Else
                Dim pls = pwdLastSet
                Return If(pls IsNot Nothing AndAlso pls > 0, Date.FromFileTime(pls).AddDays(Domain.MaxPwdAge), Nothing)
            End If
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property passwordExpiresFormated() As String
        Get
            Dim pne = passwordNeverExpires
            If pne IsNot Nothing AndAlso pne Then
                Return My.Resources.str_Never
            Else
                Dim pls = pwdLastSet
                Return If(pls Is Nothing, My.Resources.str_Unknown, If(pls = 0, My.Resources.str_Expired, Date.FromFileTime(pls).AddDays(Domain.MaxPwdAge).ToString))
            End If
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public Property userMustChangePasswordNextLogon() As Boolean?
        Get
            Dim pls = pwdLastSet
            If pls IsNot Nothing Then
                Return Not CBool(pls)
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Boolean?)
            If value IsNot Nothing Then
                If value Then
                    pwdLastSet = 0 'password expired
                Else
                    pwdLastSet = -1 'password changed today
                End If
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property sAMAccountName() As String
        Get
            Return GetValue("sAMAccountName")
        End Get
        Set(ByVal value As String)
            SetAttribute("sAMAccountName", value)

            NotifyPropertyChanged("sAMAccountName")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property userAccountControl() As Integer?
        Get
            Return GetValue("userAccountControl")
        End Get
        Set(ByVal value As Integer?)
            If value IsNot Nothing Then SetAttribute("userAccountControl", value)

            NotifyPropertyChanged("userAccountControl")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public Property normalAccount() As Boolean?
        Get
            Dim uac = userAccountControl
            Return If(uac Is Nothing, Nothing, uac And mdlVariables.ADS_UF_NORMAL_ACCOUNT)
        End Get
        Set(ByVal value As Boolean?)
            If value Is Nothing Then Exit Property

            Dim uac = userAccountControl
            If value Then
                If Not (uac And mdlVariables.ADS_UF_NORMAL_ACCOUNT) Then
                    userAccountControl = uac + mdlVariables.ADS_UF_NORMAL_ACCOUNT
                End If
            Else
                If (uac And mdlVariables.ADS_UF_NORMAL_ACCOUNT) Then
                    userAccountControl = uac - mdlVariables.ADS_UF_NORMAL_ACCOUNT
                End If
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public Property disabled() As Boolean?
        Get
            Dim uac = userAccountControl
            Return If(uac Is Nothing, Nothing, uac And mdlVariables.ADS_UF_ACCOUNTDISABLE)
        End Get
        Set(ByVal value As Boolean?)
            If value Is Nothing Then Exit Property

            Dim uac = userAccountControl
            If value Then
                If Not (uac And mdlVariables.ADS_UF_ACCOUNTDISABLE) Then
                    userAccountControl = uac + mdlVariables.ADS_UF_ACCOUNTDISABLE
                End If
            Else
                If (uac And mdlVariables.ADS_UF_ACCOUNTDISABLE) Then
                    userAccountControl = uac - mdlVariables.ADS_UF_ACCOUNTDISABLE
                End If
            End If

            If Domain.LogActions = True AndAlso Not String.IsNullOrEmpty(Domain.LogActionsAttribute) Then
                SetAttribute(Domain.LogActionsAttribute, String.Format("{0} {1} ({2})", If(value, My.Resources.str_Disabled, My.Resources.str_Enabled), Domain.Username, Now.ToShortTimeString & " " & Now.ToShortDateString))
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property disabledFormated() As String
        Get
            Dim d = disabled
            If d IsNot Nothing Then
                If d Then
                    Return "✓"
                Else
                    Return ""
                End If
            Else
                Return My.Resources.str_Unknown
            End If
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public Property passwordNeverExpires() As Boolean?
        Get
            Dim uac = userAccountControl
            Return If(uac Is Nothing, Nothing, uac And mdlVariables.ADS_UF_DONT_EXPIRE_PASSWD)
        End Get
        Set(ByVal value As Boolean?)
            If value Is Nothing Then Exit Property

            Dim uac = userAccountControl
            If value Then
                If Not (uac And mdlVariables.ADS_UF_DONT_EXPIRE_PASSWD) Then
                    userAccountControl = uac + mdlVariables.ADS_UF_DONT_EXPIRE_PASSWD
                End If
            Else
                If (uac And mdlVariables.ADS_UF_DONT_EXPIRE_PASSWD) Then
                    userAccountControl = uac - mdlVariables.ADS_UF_DONT_EXPIRE_PASSWD
                End If
            End If
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property whenCreated() As Date
        Get
            Dim val = GetValue("whenCreated")
            Return Date.ParseExact(val, "yyyyMMddHHmmss.f'Z'", Globalization.CultureInfo.InvariantCulture)
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    <ExtendedProperty>
    Public ReadOnly Property whenCreatedFormated() As String
        Get
            Return If(whenCreated = Nothing, My.Resources.str_Unknown, whenCreated.ToString)
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property whenChanged() As Date
        Get
            Dim val = GetValue("whenChanged")
            Return Date.ParseExact(val, "yyyyMMddHHmmss.f'Z'", Globalization.CultureInfo.InvariantCulture)
        End Get
    End Property

    Private _manager As clsDirectoryObject
    <RegistrySerializerIgnorable(True)>
    Public Property manager() As clsDirectoryObject
        Get
            If _manager Is Nothing Then
                Dim managerDN As String = GetValue("manager")

                If managerDN Is Nothing Then
                    _manager = Nothing
                Else
                    _manager = New clsDirectoryObject(managerDN, Domain)
                End If
            End If
            Return _manager
        End Get
        Set(value As clsDirectoryObject)
            If value Is Nothing Then
                Try
                    SetAttribute("manager", Nothing)
                Catch ex As Exception
                    Throw ex
                End Try
            Else
                SetAttribute("manager", value.distinguishedName)
                _manager = value
            End If

            NotifyPropertyChanged("manager")
        End Set
    End Property

    Private _directreports As ObservableCollection(Of clsDirectoryObject)
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property directReports() As ObservableCollection(Of clsDirectoryObject)
        Get
            If _directreports Is Nothing Then
                Dim o As String() = GetValue("directReports")

                If o Is Nothing Then
                    _directreports = New ObservableCollection(Of clsDirectoryObject)
                Else
                    _directreports = New ObservableCollection(Of clsDirectoryObject)(o.Select(Function(x) New clsDirectoryObject(x, Domain)).ToArray)
                End If
            End If
            Return _directreports
        End Get
    End Property

    Private _managedby As clsDirectoryObject
    <RegistrySerializerIgnorable(True)>
    Public Property managedBy() As clsDirectoryObject
        Get
            If _managedby Is Nothing Then
                Dim managerDN As String = GetValue("managedBy")

                If managerDN Is Nothing Then
                    _managedby = Nothing
                Else
                    _managedby = New clsDirectoryObject(managerDN, Domain)
                End If
            End If
            Return _managedby
        End Get
        Set(value As clsDirectoryObject)
            If value Is Nothing Then
                Try
                    SetAttribute("managedBy", Nothing)
                Catch ex As Exception
                    Throw ex
                End Try
            Else
                SetAttribute("managedBy", value.distinguishedName)
                _managedby = value
            End If

            NotifyPropertyChanged("managedBy")
        End Set
    End Property

    Private _managedobjects As ObservableCollection(Of clsDirectoryObject)
    <RegistrySerializerIgnorable(True)>
    Public Property managedObjects As ObservableCollection(Of clsDirectoryObject)
        Get
            If _managedobjects Is Nothing Then
                Dim o As String() = GetValue("managedObjects")

                If o Is Nothing Then
                    _managedobjects = New ObservableCollection(Of clsDirectoryObject)
                Else
                    _managedobjects = New ObservableCollection(Of clsDirectoryObject)(o.Select(Function(x) New clsDirectoryObject(x, Domain)).ToArray)
                End If
            End If
            Return _managedobjects
        End Get
        Set(value As ObservableCollection(Of clsDirectoryObject))
            _managedobjects = value

            NotifyPropertyChanged("managedObjects")
        End Set
    End Property

    Private _memberof As ObservableCollection(Of clsDirectoryObject)
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property memberOf() As ObservableCollection(Of clsDirectoryObject)
        Get
            If _memberof Is Nothing Then
                Dim o As String() = GetValue("memberOf")

                If o Is Nothing Then
                    _memberof = New ObservableCollection(Of clsDirectoryObject)
                Else
                    _memberof = New ObservableCollection(Of clsDirectoryObject)(o.Select(Function(x) New clsDirectoryObject(x, Domain)).ToArray)
                End If
            End If
            Return _memberof
        End Get
    End Property

    Private _member As ObservableCollection(Of clsDirectoryObject)
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property member() As ObservableCollection(Of clsDirectoryObject)
        Get

            'Dim searchRequest = New SearchRequest(Domain.DefaultNamingContext, "(memberof=" & distinguishedName & ")", Protocols.SearchScope.Subtree, {"distinguishedName"})
            'searchRequest.Controls.Add(New ShowDeletedControl())
            'searchRequest.Controls.Add(New SearchOptionsControl(SearchOption.DomainScope))

            'Dim response As SearchResponse
            'Try
            '    response = Connection.SendRequest(searchRequest)

            '    Debug.Print(String.Format("Total {0}", response.Entries.Count))

            '    For Each entry In response.Entries
            '        Debug.Print(DirectCast(entry, System.DirectoryServices.Protocols.SearchResultEntry).DistinguishedName)
            '    Next

            'Catch ex As Exception
            '    Return Nothing
            'End Try

            If _member Is Nothing Then
                Refresh({"member"})
                Dim members As New List(Of String)
                For Each k In cache.Keys.Where(Function(name) name = "member" Or name.StartsWith("member;"))
                    For Each m As String In GetValue(k)
                        members.Add(m)
                        Debug.Print(m)
                    Next
                Next
                Dim o As String() = members.ToArray

                If o Is Nothing Then
                    _member = New ObservableCollection(Of clsDirectoryObject)
                Else
                    _member = New ObservableCollection(Of clsDirectoryObject)(o.Select(Function(x) New clsDirectoryObject(x, Domain)).ToArray)
                End If
            End If
            Return _member

        End Get
    End Property

#End Region

#Region "Events"

    Private Sub Domain_ObjectChanged(sender As Object, e As ObjectChangedEventArgs) Handles _domain.ObjectChanged
        If e.Entry Is Nothing OrElse e.DistinguishedName <> distinguishedName Then Exit Sub

        Dim newcache As New Dictionary(Of String, DirectoryAttribute)
        For Each attribute As DirectoryAttribute In e.Entry.Attributes.Values
            newcache.Add(attribute.Name, attribute)
        Next

        Try
            For Each n In newcache.Keys
                If cache.ContainsKey(n) Then
                    If Not AttributeEquals(cache(n), newcache(n)) Then
                        Debug.WriteLine(distinguishedName & " - " & n & " attribute changed:")
                        Debug.Write("From:")
                        For Each item In cache(n).GetValues(GetType(String))
                            Debug.WriteLine(vbTab & item)
                        Next
                        Debug.Write("To:")
                        For Each item In newcache(n).GetValues(GetType(String))
                            Debug.WriteLine(vbTab & item)
                        Next
                        cache(n) = newcache(n)
                        NotifyPropertyChanged(n)
                    End If
                Else
                    Debug.WriteLine(distinguishedName & " - " & n & " attribute added:")
                    Debug.Write("Value:")
                    For Each item In newcache(n).GetValues(GetType(String))
                        Debug.WriteLine(vbTab & item)
                    Next
                    cache.Add(n, newcache(n))
                    NotifyPropertyChanged(n)
                End If

            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Function AttributeEquals(a As DirectoryAttribute, b As DirectoryAttribute) As Boolean
        Dim aval() = a.GetValues(GetType(String))
        Dim bval() = b.GetValues(GetType(String))
        If aval.Count <> bval.Count Then Return False

        For i = 0 To aval.Count - 1
            If Not aval(i) = bval(i) Then Return False
        Next

        Return True
    End Function

    Private Sub NotifyPropertyChanged(propertyName As String)
        OnPropertyChanged(New PropertyChangedEventArgs(propertyName))

        Select Case LCase(propertyName)
            Case "distinguishedname"
                NotifyPropertyChanged("distinguishedNameFormated")
                NotifyPropertyChanged("name")
            Case "distinguishedname"
                NotifyPropertyChanged("distinguishedNameFormated")
            Case "userprincipalname"
                NotifyPropertyChanged("userPrincipalNameName")
                NotifyPropertyChanged("userPrincipalNameDomain")
            Case "thumbnailphoto"
                NotifyPropertyChanged("Image")
            Case "grouptype"
                NotifyPropertyChanged("groupTypeScopeDomainLocal")
                NotifyPropertyChanged("groupTypeScopeDomainGlobal")
                NotifyPropertyChanged("groupTypeScopeDomainUniversal")
                NotifyPropertyChanged("groupTypeSecurity")
                NotifyPropertyChanged("groupTypeDistribution")
                NotifyPropertyChanged("StatusImage")
            Case "accountexpires"
                NotifyPropertyChanged("accountExpiresDate")
                NotifyPropertyChanged("accountExpiresFormated")
                NotifyPropertyChanged("accountNeverExpires")
                NotifyPropertyChanged("accountExpiresAt")
                NotifyPropertyChanged("Status")
                NotifyPropertyChanged("StatusFormatted")
                NotifyPropertyChanged("StatusImage")
                NotifyPropertyChanged("Image")
            Case "pwdlastset"
                NotifyPropertyChanged("pwdLastSetDate")
                NotifyPropertyChanged("pwdLastSetFormated")
                NotifyPropertyChanged("passwordExpiresDate")
                NotifyPropertyChanged("passwordExpiresFormated")
                NotifyPropertyChanged("userMustChangePasswordNextLogon")
                NotifyPropertyChanged("Status")
                NotifyPropertyChanged("StatusFormatted")
                NotifyPropertyChanged("StatusImage")
                NotifyPropertyChanged("Image")
            Case "useraccountcontrol"
                NotifyPropertyChanged("normalAccount")
                NotifyPropertyChanged("disabled")
                NotifyPropertyChanged("disabledFormated")
                NotifyPropertyChanged("passwordNeverExpires")
                NotifyPropertyChanged("passwordExpiresDate")
                NotifyPropertyChanged("passwordExpiresFormated")
                NotifyPropertyChanged("userMustChangePasswordNextLogon")
                NotifyPropertyChanged("Status")
                NotifyPropertyChanged("StatusFormatted")
                NotifyPropertyChanged("StatusImage")
                NotifyPropertyChanged("Image")
        End Select
    End Sub

#End Region

End Class

Public Class ExtendedPropertyAttribute
    Inherits Attribute

End Class
