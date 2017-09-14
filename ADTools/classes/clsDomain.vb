Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.DirectoryServices
Imports System.Threading
Imports System.Threading.Tasks
Imports CredentialManagement
Imports IRegisty

Public Class clsDomain
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _rootdse As DirectoryEntry
    Private _rootdsepath As String
    Private _defaultnamingcontext As DirectoryEntry
    Private _defaultnamingcontextpath As String
    Private _configurationnamingcontext As DirectoryEntry
    Private _configurationnamingcontextpath As String
    Private _schemanamingcontext As DirectoryEntry
    Private _schemanamingcontextpath As String
    Private _searchroot As DirectoryEntry
    Private _searchrootpath As String

    Private _properties As New ObservableCollection(Of clsDomainProperty)
    Private _maxpwdage As Integer
    Private _suffixes As New ObservableCollection(Of String)

    Private _name As String
    Private _username As String
    Private _password As String

    Private _usernamepattern As String = ""
    Private _computerpattern As String = ""
    Private _telephonenumberpattern As New ObservableCollection(Of clsTelephoneNumberPattern)

    Private _defaultpassword As String = ""

    Private _exchangeservers As New ObservableCollection(Of String)
    Private _useexchange As Boolean
    Private _exchangeserver As String
    Private _mailboxpattern As String = ""

    Private _issearchable As Boolean = True
    Private _validated As Boolean

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Private Sub NotifyPropertyChanged(propertyName As String)
        OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Sub New()

    End Sub

    <RegistrySerializerAfterSerialize(True)>
    Public Sub AfterSerialize()
        Dim cred As New Credential("", "", My.Application.Info.AssemblyName & ": " & Name, CredentialType.Generic)
        cred.PersistanceType = PersistanceType.Enterprise
        cred.Username = _username
        cred.Password = _password
        cred.Save()
    End Sub

    <RegistrySerializerAfterDeserialize(True)>
    Public Sub AfterDeserialize()
        Dim cred As New Credential("", "", My.Application.Info.AssemblyName & ": " & Name, CredentialType.Generic)
        cred.PersistanceType = PersistanceType.Enterprise
        cred.Load()
        Username = cred.Username
        Password = cred.Password

        Dim connecttester As Object = Nothing
        If Not String.IsNullOrEmpty(Username) And Not String.IsNullOrEmpty(Password) Then
            Try
                RootDSE = New DirectoryEntry(RootDSEPath, Username, Password)
                connecttester = RootDSE.NativeObject
                DefaultNamingContext = New DirectoryEntry(DefaultNamingContextPath, Username, Password)
                connecttester = DefaultNamingContext.NativeObject
                ConfigurationNamingContext = New DirectoryEntry(ConfigurationNamingContextPath, Username, Password)
                connecttester = ConfigurationNamingContext.NativeObject
                SchemaNamingContext = New DirectoryEntry(SchemaNamingContextPath, Username, Password)
                connecttester = SchemaNamingContext.NativeObject
                SearchRoot = New DirectoryEntry(SearchRootPath, Username, Password)
                connecttester = SearchRoot.NativeObject

                Validated = True
            Catch
            End Try
        End If
    End Sub

    Public Async Function ConnectAsync() As Task(Of Boolean)
        Validated = False
        If Len(Name) = 0 Or Len(Username) = 0 Or Len(Password) = 0 Then Return False

        Dim connectionstring As String = "LDAP://" & Name & "/"
        Dim newDefaultNamingContext As String = ""
        Dim success As Boolean = False

        'RootDSE
        success = Await Task.Run(
            Function()
                Try
                    RootDSE = New DirectoryEntry(connectionstring & "RootDSE", Username, Password)
                    If RootDSE.Properties.Count = 0 Then Return False
                Catch
                    RootDSE = Nothing
                    Return False
                End Try
                Return True
            End Function)
        If Not success Then Return False

        'defaultNamingContext
        success = Await Task.Run(
            Function()
                Try
                    DefaultNamingContext = New DirectoryEntry(connectionstring & GetLDAPProperty(_rootdse.Properties, "defaultNamingContext"), Username, Password)
                    If DefaultNamingContext.Properties.Count = 0 Then Return False
                Catch
                    DefaultNamingContext = Nothing
                    Return False
                End Try
                Return True
            End Function)
        If Not success Then Return False

        'configurationNamingContext
        success = Await Task.Run(
            Function()
                Try
                    ConfigurationNamingContext = New DirectoryEntry(connectionstring & GetLDAPProperty(_rootdse.Properties, "configurationNamingContext"), Username, Password)
                    If ConfigurationNamingContext.Properties.Count = 0 Then Return False
                Catch
                    ConfigurationNamingContext = Nothing
                    Return False
                End Try
                Return True
            End Function)
        If Not success Then Return False

        'schemaNamingContext
        success = Await Task.Run(
            Function()
                Try
                    SchemaNamingContext = New DirectoryEntry(connectionstring & GetLDAPProperty(_rootdse.Properties, "schemaNamingContext"), Username, Password)
                    If SchemaNamingContext.Properties.Count = 0 Then Return False
                Catch ex As Exception
                    SchemaNamingContext = Nothing
                    Return False
                End Try
                Return True
            End Function)
        If Not success Then Return False

        'properties
        Properties = Await Task.Run(
            Function()
                Dim p As New ObservableCollection(Of clsDomainProperty)
                Try
                    p.Clear()
                    p.Add(New clsDomainProperty("Пороговое значение блокировки", String.Format("{0} ошибок входа", GetLDAPProperty(_defaultnamingcontext.Properties, "lockoutThreshold"))))
                    p.Add(New clsDomainProperty("Время до сброса счетчика блокировки", String.Format("{0} минут", -TimeSpan.FromTicks(LongFromLargeInteger(_defaultnamingcontext.Properties("lockoutDuration")(0))).Minutes)))
                    p.Add(New clsDomainProperty("Продолжительность блокировки учетной записи", String.Format("{0} минут", -TimeSpan.FromTicks(LongFromLargeInteger(_defaultnamingcontext.Properties("lockOutObservationWindow")(0))).Minutes)))
                    p.Add(New clsDomainProperty("Максимальный срок действия пароля", String.Format("{0} дней", -TimeSpan.FromTicks(LongFromLargeInteger(_defaultnamingcontext.Properties("maxPwdAge")(0))).Days)))
                    p.Add(New clsDomainProperty("Минимальный срок действия пароля", String.Format("{0} дней", -TimeSpan.FromTicks(LongFromLargeInteger(_defaultnamingcontext.Properties("minPwdAge")(0))).Days)))
                    p.Add(New clsDomainProperty("Минимальная длина пароля", String.Format("{0} символов", GetLDAPProperty(_defaultnamingcontext.Properties, "minPwdLength") & " симв.")))
                    p.Add(New clsDomainProperty("Пароль должен отвечать требованиям сложности", String.Format("{0}", GetLDAPProperty(_defaultnamingcontext.Properties, "pwdProperties"))))
                    p.Add(New clsDomainProperty("Вести журнал паролей", String.Format("{0} сохраненных паролей", GetLDAPProperty(_defaultnamingcontext.Properties, "pwdHistoryLength"))))
                Catch
                    p.Clear()
                End Try
                Return p
            End Function)

        'maximum password age
        MaxPwdAge = Await Task.Run(
            Function()
                Try
                    Return -TimeSpan.FromTicks(LongFromLargeInteger(GetLDAPProperty(_defaultnamingcontext.Properties, "maxPwdAge"))).Days
                Catch ex As Exception
                    Return 0
                End Try
            End Function)

        'domain suffixes
        Suffixes = Await Task.Run(
            Function()
                Dim s As New ObservableCollection(Of String)
                Try
                    Dim LDAPsearcher As New DirectorySearcher(ConfigurationNamingContext)
                    Dim LDAPresults As SearchResultCollection = Nothing

                    LDAPsearcher.Filter = "(&(objectClass=crossRef)(systemFlags=3))"
                    LDAPresults = LDAPsearcher.FindAll()
                    For Each LDAPresult As SearchResult In LDAPresults
                        s.Add(LCase(GetLDAPProperty(LDAPresult.Properties, "dnsRoot")))
                    Next LDAPresult
                Catch ex As Exception
                    s.Clear()
                End Try
                Return s
            End Function)


        'search root
        If SearchRoot Is Nothing AndAlso DefaultNamingContext IsNot Nothing Then SearchRoot = DefaultNamingContext


        'exchange servers
        ExchangeServers = Await Task.Run(
            Function()
                Dim e As New ObservableCollection(Of String)
                Try
                    Dim LDAPsearcher As New DirectorySearcher(_configurationnamingcontext)
                    Dim LDAPresults As SearchResultCollection = Nothing
                    Dim LDAPresult As SearchResult

                    LDAPsearcher.Filter = "(objectClass=msExchExchangeServer)"
                    LDAPresults = LDAPsearcher.FindAll()
                    For Each LDAPresult In LDAPresults
                        e.Add(LCase(GetLDAPProperty(LDAPresult.Properties, "name")))
                    Next LDAPresult
                Catch ex As Exception
                    e.Clear()
                End Try
                Return e
            End Function)

        If Not String.IsNullOrEmpty(ExchangeServer) AndAlso Not ExchangeServers.Contains(ExchangeServer) Then
            UseExchange = False
            ExchangeServer = Nothing
        End If

        Validated = True

        Return True
    End Function

    <RegistrySerializerIgnorable(True)>
    Public Property RootDSE() As DirectoryEntry
        Get
            Return _rootdse
        End Get
        Set(value As DirectoryEntry)
            _rootdse = value
            If value IsNot Nothing Then RootDSEPath = value.Path
            NotifyPropertyChanged("RootDSE")
        End Set
    End Property

    <RegistrySerializerAlias("RootDSE")>
    Public Property RootDSEPath() As String
        Get
            Return _rootdsepath
        End Get
        Set(value As String)
            _rootdsepath = value
            NotifyPropertyChanged("RootDSEPath")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property DefaultNamingContext() As DirectoryEntry
        Get
            Return _defaultnamingcontext
        End Get
        Set(value As DirectoryEntry)
            _defaultnamingcontext = value
            If value IsNot Nothing Then DefaultNamingContextPath = value.Path
            NotifyPropertyChanged("DefaultNamingContext")
        End Set
    End Property

    <RegistrySerializerAlias("DefaultNamingContext")>
    Public Property DefaultNamingContextPath() As String
        Get
            Return _defaultnamingcontextpath
        End Get
        Set(value As String)
            _defaultnamingcontextpath = value
            NotifyPropertyChanged("DefaultNamingContextPath")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ConfigurationNamingContext() As DirectoryEntry
        Get
            Return _configurationnamingcontext
        End Get
        Set(value As DirectoryEntry)
            _configurationnamingcontext = value
            If value IsNot Nothing Then ConfigurationNamingContextPath = value.Path
            NotifyPropertyChanged("ConfigurationNamingContext")
        End Set
    End Property

    <RegistrySerializerAlias("ConfigurationNamingContext")>
    Public Property ConfigurationNamingContextPath() As String
        Get
            Return _configurationnamingcontextpath
        End Get
        Set(value As String)
            _configurationnamingcontextpath = value
            NotifyPropertyChanged("ConfigurationNamingContextPath")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property SchemaNamingContext() As DirectoryEntry
        Get
            Return _schemanamingcontext
        End Get
        Set(value As DirectoryEntry)
            _schemanamingcontext = value
            If value IsNot Nothing Then SchemaNamingContextPath = value.Path
            NotifyPropertyChanged("SchemaNamingContext")
        End Set
    End Property

    <RegistrySerializerAlias("SchemaNamingContext")>
    Public Property SchemaNamingContextPath() As String
        Get
            Return _schemanamingcontextpath
        End Get
        Set(value As String)
            _schemanamingcontextpath = value
            NotifyPropertyChanged("SchemaNamingContextPath")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property SearchRoot() As DirectoryEntry
        Get
            Return _searchroot
        End Get
        Set(value As DirectoryEntry)
            _searchroot = value
            If value IsNot Nothing Then SearchRootPath = value.Path
            NotifyPropertyChanged("SearchRoot")
        End Set
    End Property

    <RegistrySerializerAlias("SearchRoot")>
    Public Property SearchRootPath() As String
        Get
            Return _searchrootpath
        End Get
        Set(value As String)
            _searchrootpath = value
            NotifyPropertyChanged("SearchRootPath")
        End Set
    End Property

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

    Public Property MaxPwdAge() As Integer
        Get
            Return _maxpwdage
        End Get
        Set(value As Integer)
            _maxpwdage = value
            NotifyPropertyChanged("MaxPwdAge")
        End Set
    End Property

    Public Property Suffixes As ObservableCollection(Of String)
        Get
            Return _suffixes
        End Get
        Set(value As ObservableCollection(Of String))
            _suffixes = value
            NotifyPropertyChanged("Suffixes")
        End Set
    End Property

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
            NotifyPropertyChanged("Name")
        End Set
    End Property

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

    Public Property UsernamePattern() As String
        Get
            Return _usernamepattern
        End Get
        Set(value As String)
            _usernamepattern = value
            NotifyPropertyChanged("UsernamePattern")
        End Set
    End Property

    Public Property ComputerPattern() As String
        Get
            Return _computerpattern
        End Get
        Set(value As String)
            _computerpattern = value
            NotifyPropertyChanged("ComputerPattern")
        End Set
    End Property

    Public Property MailboxPattern As String
        Get
            Return _mailboxpattern
        End Get
        Set(value As String)
            _mailboxpattern = value
            NotifyPropertyChanged("MailboxPattern")
        End Set
    End Property

    Public Property TelephoneNumberPattern() As ObservableCollection(Of clsTelephoneNumberPattern)
        Get
            Return _telephonenumberpattern
        End Get
        Set(value As ObservableCollection(Of clsTelephoneNumberPattern))
            _telephonenumberpattern = value
            NotifyPropertyChanged("TelephoneNumberPattern")
        End Set
    End Property

    Public Property DefaultPassword() As String
        Get
            Return _defaultpassword
        End Get
        Set(value As String)
            _defaultpassword = value
            NotifyPropertyChanged("DefaultPassword")
        End Set
    End Property

    Public Property ExchangeServers() As ObservableCollection(Of String)
        Get
            Return _exchangeservers
        End Get
        Set(value As ObservableCollection(Of String))
            _exchangeservers = value
            NotifyPropertyChanged("ExchangeServers")
        End Set
    End Property

    Public Property UseExchange() As Boolean
        Get
            Return _useexchange
        End Get
        Set(value As Boolean)
            _useexchange = value
            NotifyPropertyChanged("UseExchange")
        End Set
    End Property

    Public Property ExchangeServer() As String
        Get
            Return _exchangeserver
        End Get
        Set(value As String)
            _exchangeserver = value
            NotifyPropertyChanged("ExchangeServer")
        End Set
    End Property

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

    Public Property IsSearchable() As Boolean
        Get
            Return _issearchable
        End Get
        Set(value As Boolean)
            _issearchable = value
            NotifyPropertyChanged("IsSearchable")
        End Set
    End Property
End Class
