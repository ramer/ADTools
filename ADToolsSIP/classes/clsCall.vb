Public Class clsCall

    Sub New(displayName As String, telephoneNumber As String, data As String)
        _displayName = displayName
        _telephoneNumber = telephoneNumber
        _filter = String.Format("{0}* / *{1}", displayName, telephoneNumber)
        _data = data
        _timestamp = Now
    End Sub

    Private _displayName As String
    Public ReadOnly Property displayName() As String
        Get
            Return _displayName
        End Get
    End Property

    Private _telephoneNumber As String
    Public ReadOnly Property telephoneNumber() As String
        Get
            Return _telephoneNumber
        End Get
    End Property

    Private _filter As String
    Public ReadOnly Property Filter() As String
        Get
            Return _filter
        End Get
    End Property

    Private _data As String
    Public ReadOnly Property Data() As String
        Get
            Return _data
        End Get
    End Property

    Private _timestamp As Date
    Public ReadOnly Property Timestamp() As Date
        Get
            Return _timestamp
        End Get
    End Property

End Class
