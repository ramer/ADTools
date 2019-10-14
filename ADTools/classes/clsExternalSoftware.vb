Imports System.ComponentModel
Imports IRegisty

Public Class clsExternalSoftware
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New()

    End Sub

    Private _label As String
    Public Property Label() As String
        Get
            Return _label
        End Get
        Set(ByVal value As String)
            _label = value

            NotifyPropertyChanged("Label")
        End Set
    End Property

    Private _path As String
    Public Property Path() As String
        Get
            Return _path
        End Get
        Set(ByVal value As String)
            _path = value

            NotifyPropertyChanged("Path")
            NotifyPropertyChanged("Image")
        End Set
    End Property

    Private _arguments As String
    Public Property Arguments() As String
        Get
            Return _arguments
        End Get
        Set(ByVal value As String)
            _arguments = value

            NotifyPropertyChanged("Arguments")
        End Set
    End Property

    Private _currentcredentials As Boolean
    Public Property CurrentCredentials() As Boolean
        Get
            Return _currentcredentials
        End Get
        Set(ByVal value As Boolean)
            _currentcredentials = value

            NotifyPropertyChanged("CurrentCredentials")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property Image() As ImageSource
        Get
            Try
                Return GetApplicationIcon(Path)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

End Class
