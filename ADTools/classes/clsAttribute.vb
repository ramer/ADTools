Public Class clsAttribute
    Private _name As String
    Private _label As String
    Private _value As Object

    Sub New()

    End Sub

    Sub New(Name As String,
            Label As String,
            Optional Value As Object = Nothing)

        _name = Name
        _label = Label
        _value = Value
    End Sub

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Label() As String
        Get
            Return _label
        End Get
        Set(ByVal value As String)
            _label = value
        End Set
    End Property

    Public Property Value() As Object
        Get
            Return _value
        End Get
        Set(ByVal value As Object)
            _value = value
        End Set
    End Property

End Class
