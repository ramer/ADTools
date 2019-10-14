Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.DirectoryServices.Protocols
Imports System.Net
Imports CredentialManagement
Imports HandlebarsDotNet
Imports IRegisty

<DebuggerDisplay("Domain={Name}")>
Public Class clsDomain
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
    Public Event ObjectChanged(sender As Object, e As ObjectChangedEventArgs)

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Protected Overridable Sub OnObjectChanged(e As ObjectChangedEventArgs)
        RaiseEvent ObjectChanged(Me, e)
    End Sub

    Private Sub NotifyPropertyChanged(propertyName As String)
        OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Private Sub NotifyObjectChanged(ocea As ObjectChangedEventArgs)
        OnObjectChanged(ocea)
    End Sub

    Private Sub Watcher_ObjectChanged(sender As Object, e As ObjectChangedEventArgs) Handles _watcher.ObjectChanged
        NotifyObjectChanged(e)
    End Sub

    Public Sub AfterSerialize()
        SaveCredentials()
    End Sub

    Sub New()

    End Sub

#Region "Functions"

    ''' <summary>
    ''' Uses Credential Management to load credentials from Windows store
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Function LoadCredentials() As Boolean
        If String.IsNullOrEmpty(Name) Then Return False

        Try
            Dim cred As New Credential("", "", "ADTools: " & Name, CredentialType.Generic)
            cred.PersistanceType = PersistanceType.Enterprise
            cred.Load()
            Username = cred.Username
            Password = cred.Password
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Uses Credential Management to save credentials to Windows store
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Function SaveCredentials() As Boolean
        If String.IsNullOrEmpty(Name) Or String.IsNullOrEmpty(Username) Or String.IsNullOrEmpty(Password) Then Return False

        Try
            Dim cred As New Credential("", "", "ADTools: " & Name, CredentialType.Generic)
            cred.PersistanceType = PersistanceType.Enterprise
            cred.Username = _username
            cred.Password = _password
            cred.Save()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Initializing LDAP-connection, using Username as Password
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Function SetupConnection() As Boolean
        If String.IsNullOrEmpty(Name) Or String.IsNullOrEmpty(Username) Or String.IsNullOrEmpty(Password) Then Return False

        Dim endpoint As String = If(String.IsNullOrEmpty(Server), Name, Server)

        If Connection IsNot Nothing Then
            If Watcher IsNot Nothing Then StopWatcher()
            Connection.Dispose()
        End If

        Dim ldapdi As New LdapDirectoryIdentifier(endpoint)
        Dim ldapnc As New NetworkCredential(Username, Password)
        Dim ldapconnection As New LdapConnection(ldapdi, ldapnc)
        ldapconnection.SessionOptions.TcpKeepAlive = True
        ldapconnection.SessionOptions.AutoReconnect = True
        ldapconnection.Timeout = TimeSpan.FromSeconds(10)
        ldapconnection.AutoBind = True
        Try
            ldapconnection.Bind()
            Connection = ldapconnection
            Return True
        Catch ex As Exception
            Connection = Nothing
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Initializes LDAP-watcher
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Function StartWatcher() As Boolean
        If Connection Is Nothing Or String.IsNullOrEmpty(DefaultNamingContext) Then Return False
        If Watcher IsNot Nothing Then Return False
        Watcher = New clsWatcher(Connection)
        Return Watcher.Register(DefaultNamingContext, SearchScope.Subtree)
    End Function

    ''' <summary>
    ''' Stops LDAP-watcher
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Function StopWatcher() As Boolean
        If Watcher Is Nothing Then Return False
        Watcher.Dispose()
        Watcher = Nothing
        Return True
    End Function

    ''' <summary>
    ''' Requesting defaultNamingContext, configurationNamingContext, schemaNamingContext
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Async Function UpdateNamingContextsAsync() As Task(Of Boolean)
        Return Await Task.Run(
            Function()
                Try
                    Dim searchRequest = New SearchRequest(Nothing, "(objectClass=*)", SearchScope.Base, {"defaultNamingContext", "configurationNamingContext", "schemaNamingContext"})
                    Dim response As SearchResponse = Connection.SendRequest(searchRequest)

                    DefaultNamingContext = response.Entries(0).Attributes("defaultNamingContext")(0)
                    ConfigurationNamingContext = response.Entries(0).Attributes("configurationNamingContext")(0)
                    SchemaNamingContext = response.Entries(0).Attributes("schemaNamingContext")(0)

                    Return True
                Catch
                    DefaultNamingContext = Nothing
                    ConfigurationNamingContext = Nothing
                    SchemaNamingContext = Nothing

                    Return False
                End Try
            End Function)
    End Function

    ''' <summary>
    ''' Requesting primary domain properties
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Async Function UpdatePropertiesAsync() As Task(Of Boolean)
        Return Await Task.Run(
            Function()
                Try
                    Dim searchRequest = New SearchRequest(DefaultNamingContext, "(objectClass=*)", SearchScope.Base, {"lockoutThreshold", "lockoutDuration", "lockOutObservationWindow", "maxPwdAge", "minPwdAge", "minPwdLength", "pwdProperties", "pwdHistoryLength"})
                    Dim response As SearchResponse = Connection.SendRequest(searchRequest)

                    MaxPwdAge = -TimeSpan.FromTicks(Long.Parse(response.Entries(0).Attributes("maxPwdAge")(0))).Days

                    Dim p As New ObservableCollection(Of clsDomainProperty)
                    p.Add(New clsDomainProperty(My.Resources.str_LockoutThreshold, String.Format(My.Resources.str_LockoutThresholdFormat, response.Entries(0).Attributes("lockoutThreshold")(0))))
                    p.Add(New clsDomainProperty(My.Resources.str_LockoutDuration, String.Format(My.Resources.str_LockoutDurationFormat, -TimeSpan.FromTicks(Long.Parse(response.Entries(0).Attributes("lockoutDuration")(0))).Minutes)))
                    p.Add(New clsDomainProperty(My.Resources.str_LockoutObservationWindow, String.Format(My.Resources.str_LockoutObservationWindowFormat, -TimeSpan.FromTicks(Long.Parse(response.Entries(0).Attributes("lockOutObservationWindow")(0))).Minutes)))
                    p.Add(New clsDomainProperty(My.Resources.str_MaximumPasswordAge, String.Format(My.Resources.str_MaximumPasswordAgeFormat, -TimeSpan.FromTicks(Long.Parse(response.Entries(0).Attributes("maxPwdAge")(0))).Days)))
                    p.Add(New clsDomainProperty(My.Resources.str_MinimumPasswordAge, String.Format(My.Resources.str_MinimumPasswordAgeFormat, -TimeSpan.FromTicks(Long.Parse(response.Entries(0).Attributes("minPwdAge")(0))).Days)))
                    p.Add(New clsDomainProperty(My.Resources.str_MinimumPasswordLenght, String.Format(My.Resources.str_MinimumPasswordLenghtFormat, response.Entries(0).Attributes("minPwdLength")(0))))
                    p.Add(New clsDomainProperty(My.Resources.str_PasswordComplexityRequirements, String.Format(My.Resources.str_PasswordComplexityRequirementsFormat, CBool(response.Entries(0).Attributes("pwdProperties")(0)))))
                    p.Add(New clsDomainProperty(My.Resources.str_PasswordHistory, String.Format(My.Resources.str_PasswordHistoryFormat, response.Entries(0).Attributes("pwdHistoryLength")(0))))

                    Properties = p
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            End Function)
    End Function

    ''' <summary>
    ''' Requesting all object attributes from domain schema
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Async Function UpdateAttributesAsync() As Task(Of Boolean)
        Return Await Task.Run(
            Function()
                Try
                    Dim a As New Dictionary(Of String, clsAttribute)
                    Dim pageRequestControl As New PageResultRequestControl(1000)
                    Dim pageResponseControl As PageResultResponseControl
                    Dim searchRequest = New SearchRequest(SchemaNamingContext, "(objectClass=attributeSchema)", SearchScope.Subtree, {"adminDisplayName", "isSingleValued", "searchFlags", "attributeSyntax", "lDAPDisplayName", "rangeLower", "rangeUpper"})
                    searchRequest.Controls.Add(pageRequestControl)

                    Do
                        Dim response As SearchResponse = Connection.SendRequest(searchRequest)

                        For Each sre As SearchResultEntry In response.Entries
                            If sre IsNot Nothing AndAlso sre.Attributes.Contains("lDAPDisplayName") AndAlso sre.Attributes("lDAPDisplayName").Count > 0 AndAlso sre.Attributes("lDAPDisplayName")(0) IsNot Nothing Then
                                a.Add(sre.Attributes("lDAPDisplayName")(0), New clsAttribute(sre))
                            End If
                        Next

                        pageResponseControl = response.Controls(0)
                        If pageResponseControl.Cookie.Length = 0 Then Exit Do

                        pageRequestControl.Cookie = pageResponseControl.Cookie
                    Loop

                    Attributes = a
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            End Function)
    End Function

    ''' <summary>
    ''' Requesting Exchange server name from domain configuration, if exists
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Async Function UpdateExchangeServersAsync() As Task(Of Boolean)
        Return Await Task.Run(
            Function()
                Try
                    Dim e As New ObservableCollection(Of String)
                    Dim searchRequest = New SearchRequest(ConfigurationNamingContext, "(objectClass=msExchExchangeServer)", SearchScope.Subtree, {"name"})
                    Dim response As SearchResponse = Connection.SendRequest(searchRequest)

                    For Each exch As SearchResultEntry In response.Entries
                        e.Add(LCase(exch.Attributes("name")(0)))
                    Next exch

                    ExchangeServers = e

                    If Not String.IsNullOrEmpty(ExchangeServer) AndAlso Not ExchangeServers.Contains(ExchangeServer) Then
                        UseExchange = False
                        ExchangeServer = Nothing
                    End If

                    Return True
                Catch ex As Exception
                    Return False
                End Try
            End Function)
    End Function

    ''' <summary>
    ''' Requesting DNS-suffixes from domain configuration
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Async Function UpdateSuffixesAsync() As Task(Of Boolean)
        Return Await Task.Run(
            Function()
                Try
                    Dim s As New ObservableCollection(Of String)
                    Dim searchRequest = New SearchRequest(ConfigurationNamingContext, "(&(objectClass=crossRef)(systemFlags=3))", SearchScope.Subtree, {"dnsRoot"})
                    Dim response As SearchResponse = Connection.SendRequest(searchRequest)

                    For Each suffix As SearchResultEntry In response.Entries
                        s.Add(LCase(suffix.Attributes("dnsRoot")(0)))
                    Next suffix

                    Suffixes = s
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            End Function)
    End Function

    ''' <summary>
    ''' Initializing default group objects based on group DN
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Private Async Function UpdateDefaultGroupsAsync() As Task(Of Boolean)
        Return Await Task.Run(
            Function()
                Try
                    If DefaultGroupsDN.Count > 0 Then DefaultGroups = New ObservableCollection(Of clsDirectoryObject)(DefaultGroupsDN.Select(Function(x As String) New clsDirectoryObject(x, Me)).ToList)
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            End Function)
    End Function

    ''' <summary>
    ''' Initializes primary domain properties and set Validation flag on Application startup
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Public Async Function Initialize() As Task
        Validated = False

        If Not LoadCredentials() Then Return
        If Not SetupConnection() Then Return

        If String.IsNullOrEmpty(DefaultNamingContext) Or
        String.IsNullOrEmpty(ConfigurationNamingContext) Or
        String.IsNullOrEmpty(SchemaNamingContext) Then
            If Not Await UpdateNamingContextsAsync() Then Return
        End If

        Await UpdatePropertiesAsync()
        Await UpdateAttributesAsync()
        Await UpdateSuffixesAsync()
        Await UpdateExchangeServersAsync()
        Await UpdateDefaultGroupsAsync()

        If EnableWatcher Then StartWatcher()

        Validated = True
    End Function

    ''' <summary>
    ''' Initializes primary domain properties and set Validation flag on domain adding
    ''' </summary>
    ''' <returns>True on success, False otherwise</returns>
    Public Async Function ConnectAsync() As Task
        Validated = False

        If Not SetupConnection() Then Return
        If Not Await UpdateNamingContextsAsync() Then Return

        Await UpdatePropertiesAsync()
        Await UpdateAttributesAsync()
        Await UpdateSuffixesAsync()
        Await UpdateExchangeServersAsync()

        If EnableWatcher Then StartWatcher()

        Validated = True
    End Function

#End Region

#Region "Properties"

    Private _name As String
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
            NotifyPropertyChanged("Name")
        End Set
    End Property

    Private _username As String
    <RegistrySerializerIgnorable(True)>
    Public Property Username() As String
        Get
            Return _username
        End Get
        Set(value As String)
            _username = value
            NotifyPropertyChanged("Username")
        End Set
    End Property

    Private _password As String
    <RegistrySerializerIgnorable(True)>
    Public Property Password() As String
        Get
            Return _password
        End Get
        Set(value As String)
            _password = value
            NotifyPropertyChanged("Password")
        End Set
    End Property

    Private _server As String
    Public Property Server() As String
        Get
            Return _server
        End Get
        Set(value As String)
            _server = value
            NotifyPropertyChanged("Server")
        End Set
    End Property

    Private _defaultnamingcontext As String
    Public Property DefaultNamingContext() As String
        Get
            Return _defaultnamingcontext
        End Get
        Set(value As String)
            _defaultnamingcontext = value
            NotifyPropertyChanged("DefaultNamingContext")
            NotifyPropertyChanged("SearchRoot")
            NotifyPropertyChanged("UnderlyingObject")
            NotifyPropertyChanged("ChildContainers")
        End Set
    End Property

    Private _configurationnamingcontext As String
    Public Property ConfigurationNamingContext() As String
        Get
            Return _configurationnamingcontext
        End Get
        Set(value As String)
            _configurationnamingcontext = value
            NotifyPropertyChanged("ConfigurationNamingContext")
        End Set
    End Property

    Private _schemanamingcontext As String
    Public Property SchemaNamingContext() As String
        Get
            Return _schemanamingcontext
        End Get
        Set(value As String)
            _schemanamingcontext = value
            NotifyPropertyChanged("SchemaNamingContext")
        End Set
    End Property

    Private _searchroot As String
    Public Property SearchRoot() As String
        Get
            Return If(String.IsNullOrEmpty(_searchroot), _defaultnamingcontext, _searchroot)
        End Get
        Set(value As String)
            _searchroot = value
            NotifyPropertyChanged("SearchRoot")
        End Set
    End Property

    Private _properties As New ObservableCollection(Of clsDomainProperty)
    <RegistrySerializerIgnorable(True)>
    Public Property Properties() As ObservableCollection(Of clsDomainProperty)
        Get
            Return _properties
        End Get
        Set(value As ObservableCollection(Of clsDomainProperty))
            _properties = value
            NotifyPropertyChanged("Properties")
        End Set
    End Property

    Private _maxpwdage As Integer
    Public Property MaxPwdAge() As Integer
        Get
            Return _maxpwdage
        End Get
        Set(value As Integer)
            _maxpwdage = value
            NotifyPropertyChanged("MaxPwdAge")
        End Set
    End Property

    Private _suffixes As New ObservableCollection(Of String)
    Public Property Suffixes As ObservableCollection(Of String)
        Get
            Return _suffixes
        End Get
        Set(value As ObservableCollection(Of String))
            _suffixes = value
            NotifyPropertyChanged("Suffixes")
        End Set
    End Property

    Private _usernamepattern As String = ""
    Public Property UsernamePattern() As String
        Get
            Return _usernamepattern
        End Get
        Set(value As String)
            _usernamepattern = value
            _usernamepatterntemplates = Nothing
            NotifyPropertyChanged("UsernamePattern")
            NotifyPropertyChanged("UsernamePatternTemplates")
        End Set
    End Property

    Private _usernamepatterntemplates As Func(Of Object, String)() = Nothing
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property UsernamePatternTemplates As Func(Of Object, String)()
        Get
            If _usernamepatterntemplates IsNot Nothing Then Return _usernamepatterntemplates

            Dim patterns() As String = UsernamePattern.Split({",", vbCr, vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries).Select(Function(x) Trim(x)).ToArray
            _usernamepatterntemplates = patterns.Select(Function(x) Handlebars.Compile(x)).ToArray
            Return _usernamepatterntemplates
        End Get
    End Property

    Private _computerpattern As String = ""
    Public Property ComputerPattern() As String
        Get
            Return _computerpattern
        End Get
        Set(value As String)
            _computerpattern = value
            _computerpatterntemplates = Nothing
            NotifyPropertyChanged("ComputerPattern")
            NotifyPropertyChanged("ComputerPatternTemplates")
        End Set
    End Property

    Private _computerpatterntemplates As Func(Of Object, String)() = Nothing
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property ComputerPatternTemplates As Func(Of Object, String)()
        Get
            If _computerpatterntemplates IsNot Nothing Then Return _computerpatterntemplates

            Dim patterns() As String = ComputerPattern.Split({",", vbCr, vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries).Select(Function(x) Trim(x)).ToArray
            _computerpatterntemplates = patterns.Select(Function(x) Handlebars.Compile(x)).ToArray
            Return _computerpatterntemplates
        End Get
    End Property

    Private _telephonenumberpattern As New ObservableCollection(Of clsTelephoneNumberPattern)
    Public Property TelephoneNumberPattern() As ObservableCollection(Of clsTelephoneNumberPattern)
        Get
            Return _telephonenumberpattern
        End Get
        Set(value As ObservableCollection(Of clsTelephoneNumberPattern))
            _telephonenumberpattern = value
            NotifyPropertyChanged("TelephoneNumberPattern")
        End Set
    End Property

    Private _defaultpassword As String = ""
    Public Property DefaultPassword() As String
        Get
            Return _defaultpassword
        End Get
        Set(value As String)
            _defaultpassword = value
            NotifyPropertyChanged("DefaultPassword")
        End Set
    End Property

    Private _defaultgroups As New ObservableCollection(Of clsDirectoryObject)
    <RegistrySerializerIgnorable(True)>
    Public Property DefaultGroups() As ObservableCollection(Of clsDirectoryObject)
        Get
            Return _defaultgroups
        End Get
        Set(value As ObservableCollection(Of clsDirectoryObject))
            _defaultgroups = value
            _defaultgroupsdn = New ObservableCollection(Of String)(_defaultgroups.Select(Function(x) x.distinguishedName).ToList)
            NotifyPropertyChanged("DefaultGroups")
        End Set
    End Property

    Private _defaultgroupsdn As ObservableCollection(Of String)
    Public Property DefaultGroupsDN() As ObservableCollection(Of String)
        Get
            Return _defaultgroupsdn
        End Get
        Set(value As ObservableCollection(Of String))
            _defaultgroupsdn = value
            NotifyPropertyChanged("DefaultGroupsDN")
        End Set
    End Property

    Private _logactions As Boolean
    Public Property LogActions() As Boolean
        Get
            Return _logactions
        End Get
        Set(ByVal value As Boolean)
            _logactions = value
            NotifyPropertyChanged("LogActions")
        End Set
    End Property

    Private _logactionsattribute As String
    Public Property LogActionsAttribute() As String
        Get
            Return _logactionsattribute
        End Get
        Set(ByVal value As String)
            _logactionsattribute = value
            NotifyPropertyChanged("LogActionsAttribute")
        End Set
    End Property

    Private _exchangeservers As New ObservableCollection(Of String)
    Public Property ExchangeServers() As ObservableCollection(Of String)
        Get
            Return _exchangeservers
        End Get
        Set(value As ObservableCollection(Of String))
            _exchangeservers = value
            NotifyPropertyChanged("ExchangeServers")
        End Set
    End Property

    Private _useexchange As Boolean
    Public Property UseExchange() As Boolean
        Get
            Return _useexchange
        End Get
        Set(value As Boolean)
            _useexchange = value
            NotifyPropertyChanged("UseExchange")
        End Set
    End Property

    Private _exchangeserver As String
    Public Property ExchangeServer() As String
        Get
            Return _exchangeserver
        End Get
        Set(value As String)
            _exchangeserver = value
            NotifyPropertyChanged("ExchangeServer")
        End Set
    End Property

    Private _mailboxpattern As String = ""
    Public Property MailboxPattern As String
        Get
            Return _mailboxpattern
        End Get
        Set(value As String)
            _mailboxpattern = value
            NotifyPropertyChanged("MailboxPattern")
        End Set
    End Property

    Private _enablewatcher As Boolean
    Public Property EnableWatcher As Boolean
        Get
            Return _enablewatcher
        End Get
        Set(value As Boolean)
            _enablewatcher = value

            If value Then
                StartWatcher()
            Else
                StopWatcher()
            End If

            NotifyPropertyChanged("EnableWatcher")
        End Set
    End Property

    Private _connection As LdapConnection
    <RegistrySerializerIgnorable(True)>
    Public Property Connection() As LdapConnection
        Get
            Return _connection
        End Get
        Set(value As LdapConnection)
            _connection = value
            NotifyPropertyChanged("Connection")
        End Set
    End Property

    Private _attributes As New Dictionary(Of String, clsAttribute)
    <RegistrySerializerIgnorable(True)>
    Public Property Attributes() As Dictionary(Of String, clsAttribute)
        Get
            Return _attributes
        End Get
        Set(value As Dictionary(Of String, clsAttribute))
            _attributes = value
            NotifyPropertyChanged("AttributesSchema")
        End Set
    End Property

    Private _validated As Boolean
    <RegistrySerializerIgnorable(True)>
    Public Property Validated() As Boolean
        Get
            Return _validated
        End Get
        Set(value As Boolean)
            _validated = value
            NotifyPropertyChanged("Validated")
        End Set
    End Property

    Private _issearchable As Boolean = True
    <RegistrySerializerIgnorable(True)>
    Public Property IsSearchable() As Boolean
        Get
            Return _issearchable
        End Get
        Set(value As Boolean)
            _issearchable = value
            NotifyPropertyChanged("IsSearchable")
        End Set
    End Property

    Private _underlyingobject As clsDirectoryObject
    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property UnderlyingObject() As clsDirectoryObject
        Get
            If String.IsNullOrEmpty(DefaultNamingContext) Then Return Nothing
            If _underlyingobject Is Nothing Then _underlyingobject = New clsDirectoryObject(DefaultNamingContext, Me)
            Return _underlyingobject
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property ChildContainers() As ObservableCollection(Of clsDirectoryObject)
        Get
            If UnderlyingObject Is Nothing Then Return New ObservableCollection(Of clsDirectoryObject)
            AddHandler _underlyingobject.PropertyChanged, Sub(sender As Object, e As PropertyChangedEventArgs) If e.PropertyName = "ChildContainers" Then NotifyPropertyChanged("ChildContainers")
            Return _underlyingobject.ChildContainers
        End Get
    End Property

    Private WithEvents _watcher As clsWatcher
    <RegistrySerializerIgnorable(True)>
    Public Property Watcher() As clsWatcher
        Get
            Return _watcher
        End Get
        Set(value As clsWatcher)
            _watcher = value
            NotifyPropertyChanged("Watcher")
        End Set
    End Property

#End Region

End Class
