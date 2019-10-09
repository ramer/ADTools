
Imports System.ComponentModel
Imports System.DirectoryServices.Protocols
Imports IRegisty

<DebuggerDisplay("clsAttributeSchema={lDAPDisplayName}")>
Public Class clsAttribute
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private Sub NotifyPropertyChanged(propertyName As String)
        OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New()

    End Sub

    Sub New(adminDisplayName As String, isSingleValued As Boolean, searchFlags As Integer, attributeSyntax As String, lDAPDisplayName As String, rangeLower As Integer, rangeUpper As Integer, Optional IsComplex As Boolean = False)
        Me.adminDisplayName = adminDisplayName
        Me.isSingleValued = isSingleValued
        Me.searchFlags = searchFlags
        Me.attributeSyntax = attributeSyntax
        Me.lDAPDisplayName = lDAPDisplayName
        Me.rangeLower = rangeLower
        Me.rangeUpper = rangeUpper

        Me.IsComplex = IsComplex
    End Sub

    Private Function GetVal(sre As SearchResultEntry, name As String) As Object
        If sre IsNot Nothing AndAlso sre.Attributes.Contains(name) AndAlso sre.Attributes(name).Count > 0 AndAlso sre.Attributes(name)(0) IsNot Nothing Then Return sre.Attributes(name)(0)
        Return Nothing
    End Function

    Sub New(sre As SearchResultEntry)
        adminDisplayName = If(GetVal(sre, "adminDisplayName"), "")
        isSingleValued = If(GetVal(sre, "isSingleValued"), True)
        searchFlags = If(GetVal(sre, "searchFlags"), 0)
        attributeSyntax = If(GetVal(sre, "attributeSyntax"), "")
        lDAPDisplayName = If(GetVal(sre, "lDAPDisplayName"), "")
        rangeLower = If(GetVal(sre, "rangeLower"), 0)
        rangeUpper = If(GetVal(sre, "rangeUpper"), 1024)

        IsComplex = False
    End Sub

    Sub New(baseattr As clsAttribute)
        adminDisplayName = baseattr.adminDisplayName
        isSingleValued = baseattr.isSingleValued
        searchFlags = baseattr.searchFlags
        attributeSyntax = baseattr.attributeSyntax
        lDAPDisplayName = baseattr.lDAPDisplayName
        rangeLower = baseattr.rangeLower
        rangeUpper = baseattr.rangeUpper

        IsComplex = baseattr.IsComplex
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse TypeOf obj IsNot clsAttribute Then Return False
        Dim target = CType(obj, clsAttribute)
        If String.IsNullOrEmpty(adminDisplayName) Then Return False
        Return adminDisplayName.Equals(target.adminDisplayName)
    End Function

    Friend attr As DirectoryAttribute

    Public Property adminDisplayName() As String
    Public Property isSingleValued As Boolean
    Public Property searchFlags As Integer
    Public Property attributeSyntax As String
    Public Property lDAPDisplayName As String
    Public Property rangeLower As Integer
    Public Property rangeUpper As Integer

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsDefault As Boolean
        Get
            Return attributesDefaultNames.Contains(lDAPDisplayName)
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property IsIndexed As Boolean
        Get
            Return searchFlags And 1
        End Get
    End Property

    Private _iscomplex As Boolean
    <RegistrySerializerIgnorable(True)>
    Public Property IsComplex() As Boolean
        Get
            Return _iscomplex
        End Get
        Set(ByVal value As Boolean)
            _iscomplex = value
        End Set
    End Property

    Private _value As Object
    <RegistrySerializerIgnorable(True)>
    Public Overridable Property Value() As Object
        Get
            Return _value
        End Get
        Set(ByVal value As Object)
            _value = value
            NotifyPropertyChanged("Value")
        End Set
    End Property

End Class

Public Class clsAttributeBoolean
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return Boolean.Parse(attr.Item(0))
                Else
                    Dim val As New List(Of Boolean)
                    For i = 0 To attr.Count - 1
                        val.Add(Boolean.Parse(attr.Item(i)))
                    Next
                    Return val.ToArray
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeCaseInsensitiveString
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeDistinguishedName
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeDistinguishedNameWithString
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeOSIString
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeNTSecurityDescriptor
    Inherits clsAttribute
    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeIA5String
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeNumericString
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeObjectIdentifier
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeUnicodeString
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeOctetString
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(Byte()))(0)
                Else
                    Return attr.GetValues(GetType(Byte()))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeDNBinary
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(Byte()))(0)
                Else
                    Return attr.GetValues(GetType(Byte()))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeInteger
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return Integer.Parse(attr.Item(0))
                Else
                    Return attr.GetValues(GetType(Integer))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeLong
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return Long.Parse(attr.Item(0))
                Else
                    Return attr.GetValues(GetType(Long))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeUTCTime
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(String))(0)
                Else
                    Return attr.GetValues(GetType(String))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class

Public Class clsAttributeSID
    Inherits clsAttribute

    Sub New(baseattr As clsAttribute, attr As DirectoryAttribute)
        MyBase.New(baseattr)
        Me.attr = attr
    End Sub

    Public Overrides Property Value As Object
        Get
            If attr Is Nothing Then Return Nothing
            Try
                If isSingleValued Then
                    Return attr.GetValues(GetType(Byte()))(0)
                Else
                    Return attr.GetValues(GetType(Byte()))
                End If
            Catch ex As Exception
                Throw New ArgumentException
            End Try
        End Get
        Set(value As Object)

        End Set
    End Property
End Class