Public Class clsSearchHistory
    Private _root As clsDirectoryObject
    Private _filter As clsFilter

    Sub New(root As clsDirectoryObject, filter As clsFilter)
        _root = root
        _filter = filter
    End Sub

    Public ReadOnly Property Root() As clsDirectoryObject
        Get
            Return _root
        End Get
    End Property

    Public ReadOnly Property Filter() As clsFilter
        Get
            Return _filter
        End Get
    End Property

End Class
