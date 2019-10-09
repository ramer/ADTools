
Public Class clsDomainProperty

    Sub New(prop As String, value As String)
        _property = prop
        _value = value
    End Sub

    Private _property As String
    Public ReadOnly Property Prop() As String
        Get
            Return _property
        End Get
    End Property

    Private _value As String
    Public ReadOnly Property Value() As String
        Get
            Return _value
        End Get
    End Property
End Class
