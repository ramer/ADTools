Public Class clsViewColumnInfo
    Private _header As String
    Private _attributes As New List(Of String)
    Private _width As Double

    Sub New()

    End Sub

    Sub New(Header As String, Attributes As List(Of String), Optional Width As Double = 150.0)
        _header = Header
        _attributes = Attributes
        _width = Width
    End Sub

    Public Property Header() As String
        Get
            Return _header
        End Get
        Set(ByVal value As String)
            _header = value
        End Set
    End Property

    Public Property Attributes() As List(Of String)
        Get
            Return _attributes
        End Get
        Set(ByVal value As List(Of String))
            _attributes = value
        End Set
    End Property

    Public Property Width() As Double
        Get
            Return _width
        End Get
        Set(ByVal value As Double)
            _width = value
        End Set
    End Property

End Class
