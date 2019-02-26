
Imports System.ComponentModel
Imports IRegisty

<DebuggerDisplay("clsAttributeSchema={lDAPDisplayName}")>
Public Class clsAttributeSchema
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New()

    End Sub

    Sub New(adminDisplayName As String, isSingleValued As Boolean, searchFlags As Integer, attributeSyntax As String, lDAPDisplayName As String, Optional IsComplex As Boolean = False)
        Me.adminDisplayName = adminDisplayName
        Me.isSingleValued = isSingleValued
        Me.searchFlags = searchFlags
        Me.attributeSyntax = attributeSyntax
        Me.lDAPDisplayName = lDAPDisplayName
        Me.IsComplex = IsComplex
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse TypeOf obj IsNot clsAttributeSchema Then Return False
        Dim target = CType(obj, clsAttributeSchema)
        If String.IsNullOrEmpty(adminDisplayName) Then Return False
        Return adminDisplayName.Equals(target.adminDisplayName)
    End Function

    Public Sub SetValueFromDirectoryObject(obj As clsDirectoryObject)
        If obj Is Nothing Then Exit Sub
        Value = obj.GetAttribute(lDAPDisplayName)
    End Sub

    Public Sub SetStringValueFromDirectoryObject(obj As clsDirectoryObject)
        If obj Is Nothing Then Exit Sub
        Value = Join(CType(obj.GetAttribute(lDAPDisplayName, GetType(String())), String()), "; ")
    End Sub

    Public Property adminDisplayName() As String
    Public Property isSingleValued As Boolean
    Public Property searchFlags As Integer
    Public Property attributeSyntax As String
    Public Property lDAPDisplayName As String

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property ADSType As enmADSType
        Get
            If AttributeSchemeType.ContainsKey(attributeSyntax) Then
                Return AttributeSchemeType(attributeSyntax)
            Else
                Return enmADSType.ADSTYPE_UNKNOWN
            End If
        End Get
    End Property

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
    Public Property Value() As Object
        Get
            Return _value
        End Get
        Set(ByVal value As Object)
            _value = value
            NotifyPropertyChanged("Value")
        End Set
    End Property

End Class
