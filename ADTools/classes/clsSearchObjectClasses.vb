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
    Private _contact As Boolean = True
    Private _computer As Boolean = True
    Private _group As Boolean = True
    Private _organizationalunit As Boolean = False

    Sub New()

    End Sub

    Sub New(user As Boolean, contact As Boolean, computer As Boolean, group As Boolean, organizationalunit As Boolean)
        _user = user
        _contact = contact
        _computer = computer
        _group = group
        _organizationalunit = organizationalunit
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

    Public Property Contact As Boolean
        Get
            Return _contact
        End Get
        Set(value As Boolean)
            _contact = value
            NotifyPropertyChanged("Contact")
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

    Public Property OrganizationalUnit As Boolean
        Get
            Return _organizationalunit
        End Get
        Set(value As Boolean)
            _organizationalunit = value
            NotifyPropertyChanged("OrganizationalUnit")
        End Set
    End Property
End Class
