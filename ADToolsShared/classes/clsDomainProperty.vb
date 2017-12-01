
Public Class clsDomainProperty
    Dim _property As String
    Dim _value As String

    Sub New(prop As String, value As String)
        _property = prop
        _value = value
    End Sub

    Public ReadOnly Property Prop() As String
        Get
            Return _property
        End Get
    End Property

    Public ReadOnly Property Value() As String
        Get
            Return _value
        End Get
    End Property
End Class
