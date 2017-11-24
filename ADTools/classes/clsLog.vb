Public Class clsLog
    Private _timestamp As Date
    Private _message As String

    Sub New()
        _timestamp = Now
    End Sub

    Sub New(Message As String)
        _timestamp = Now
        _message = Message
    End Sub

    Public ReadOnly Property Image() As BitmapImage
        Get
            Return New BitmapImage(New Uri("pack://application:,,,/" & "images/log.png"))
        End Get
    End Property

    Public ReadOnly Property TimeStamp() As Date
        Get
            Return _timestamp
        End Get
    End Property

    Public Property Message() As String
        Get
            Return _message
        End Get
        Set(ByVal value As String)
            _message = value
        End Set
    End Property

End Class
