Public Class clsErrorLog
    Private _timestamp As Date
    Private _command As String
    Private _object As String
    Private _error As Exception

    Sub New()
        _timestamp = Now
    End Sub

    Sub New(Command As String,
        Optional [Object] As String = "",
        Optional [Error] As Exception = Nothing)

        _timestamp = Now
        _command = Command
        _object = [Object]
        _error = [Error]
    End Sub

    Public ReadOnly Property TimeStamp() As Date
        Get
            Return _timestamp
        End Get
    End Property

    Public ReadOnly Property Command() As String
        Get
            Return _command
        End Get
    End Property

    Public ReadOnly Property Obj() As String
        Get
            Return _object
        End Get
    End Property

    Public ReadOnly Property Message() As String
        Get
            If _error IsNot Nothing Then
                Return _error.Message
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property StackTrace As String
        Get
            If _error IsNot Nothing Then
                Return _error.StackTrace
            Else
                Return Nothing
            End If
        End Get
    End Property
End Class
