Imports System.ComponentModel

Public Class clsSearchObjectClasses
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Private _user As Boolean = True
    Private _computer As Boolean = True
    Private _group As Boolean = True
    Private _container As Boolean = True

    Sub New()

    End Sub

    Sub New(user As Boolean, computer As Boolean, group As Boolean, container As Boolean)
        _user = user
        _computer = computer
        _group = group
        _container = container
    End Sub

    Public Property User As Boolean
        Get
            Return _user
        End Get
        Set(value As Boolean)
            _user = value
            NotifyPropertyChanged("User")
        End Set
    End Property

    Public Property Computer As Boolean
        Get
            Return _computer
        End Get
        Set(value As Boolean)
            _computer = value
            NotifyPropertyChanged("Computer")
        End Set
    End Property

    Public Property Group As Boolean
        Get
            Return _group
        End Get
        Set(value As Boolean)
            _group = value
            NotifyPropertyChanged("Group")
        End Set
    End Property

    Public Property Container As Boolean
        Get
            Return _container
        End Get
        Set(value As Boolean)
            _container = value
            NotifyPropertyChanged("Container")
        End Set
    End Property
End Class
