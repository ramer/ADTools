Public Class clsSearchHistory
    Private _root As clsDirectoryObject
    Private _pattern As String

    Sub New(root As clsDirectoryObject, pattern As String)
        _root = root
        _pattern = pattern
    End Sub

    Public ReadOnly Property Root() As clsDirectoryObject
        Get
            Return _root
        End Get
    End Property

    Public ReadOnly Property Pattern() As String
        Get
            Return _pattern
        End Get
    End Property

End Class
