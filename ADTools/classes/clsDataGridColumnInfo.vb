Public Class clsDataGridColumnInfo
    Private _header As String
    Private _attributes As New List(Of clsAttribute)
    Private _displayindex As Integer
    Private _width As Double

    Sub New()

    End Sub

    Sub New(Header As String, Attributes As List(Of clsAttribute), Optional displayindex As Integer = 0, Optional width As Double = 150.0)
        _header = Header
        _attributes = Attributes
        _displayindex = displayindex
        _width = width
    End Sub

    Public Property Header() As String
        Get
            Return _header
        End Get
        Set(ByVal value As String)
            _header = value
        End Set
    End Property

    Public Property Attributes() As List(Of clsAttribute)
        Get
            Return _attributes
        End Get
        Set(ByVal value As List(Of clsAttribute))
            _attributes = value
        End Set
    End Property

    Public Property DisplayIndex() As Integer
        Get
            Return _displayindex
        End Get
        Set(ByVal value As Integer)
            _displayindex = value
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
