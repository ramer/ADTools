Imports System.DirectoryServices
Imports CredentialManagement
Imports IRegisty

Public Class clsDomain
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

    Private _maxpwdage As Integer

    Private _name As String
    Private _username As String
    Private _password As String

    Private _defaultpassword As String = ""

    Private _validated As Boolean

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
        Dim cred As New Credential("", "", "ADTools: " & Name, CredentialType.Generic)
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

        'maximum password age
        MaxPwdAge = Await Task.Run(
            Function()
                Try
                    Return -TimeSpan.FromTicks(LongFromLargeInteger(GetLDAPProperty(_defaultnamingcontext.Properties, "maxPwdAge"))).Days
                Catch ex As Exception
                    Return 0
                End Try
            End Function)

        'search root
        If SearchRoot Is Nothing AndAlso DefaultNamingContext IsNot Nothing Then SearchRoot = DefaultNamingContext

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
        End Set
    End Property

    <RegistrySerializerAlias("RootDSE")>
    Public Property RootDSEPath() As String
        Get
            Return _rootdsepath
        End Get
        Set(value As String)
            _rootdsepath = value
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
        End Set
    End Property

    <RegistrySerializerAlias("DefaultNamingContext")>
    Public Property DefaultNamingContextPath() As String
        Get
            Return _defaultnamingcontextpath
        End Get
        Set(value As String)
            _defaultnamingcontextpath = value
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
        End Set
    End Property

    <RegistrySerializerAlias("ConfigurationNamingContext")>
    Public Property ConfigurationNamingContextPath() As String
        Get
            Return _configurationnamingcontextpath
        End Get
        Set(value As String)
            _configurationnamingcontextpath = value
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
        End Set
    End Property

    <RegistrySerializerAlias("SchemaNamingContext")>
    Public Property SchemaNamingContextPath() As String
        Get
            Return _schemanamingcontextpath
        End Get
        Set(value As String)
            _schemanamingcontextpath = value
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
        End Set
    End Property

    <RegistrySerializerAlias("SearchRoot")>
    Public Property SearchRootPath() As String
        Get
            Return _searchrootpath
        End Get
        Set(value As String)
            _searchrootpath = value
        End Set
    End Property

    Public Property MaxPwdAge() As Integer
        Get
            Return _maxpwdage
        End Get
        Set(value As Integer)
            _maxpwdage = value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property Username() As String
        Get
            Return _username
        End Get
        Set(value As String)
            _username = value
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property Password() As String
        Get
            Return _password
        End Get
        Set(value As String)
            _password = value
        End Set
    End Property

    Public Property DefaultPassword() As String
        Get
            Return _defaultpassword
        End Get
        Set(value As String)
            _defaultpassword = value
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property Validated() As Boolean
        Get
            Return _validated
        End Get
        Set(value As Boolean)
            _validated = value
        End Set
    End Property

End Class
