Imports System.ComponentModel

Public Class clsExternalSoftware
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _label As String
    Private _path As String
    Private _arguments As String
    Private _currentcredentials As Boolean
    Private _image As ImageSource

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New(Label As String, Path As String, Arguments As String, CurrentCredentials As Boolean)
        _label = Label
        _path = Path
        _arguments = Arguments
        _currentcredentials = CurrentCredentials
        _image = Image
    End Sub

    Sub New()

    End Sub

    Public Property Label() As String
        Get
            Return _label
        End Get
        Set(ByVal value As String)
            _label = value

            NotifyPropertyChanged("Label")
        End Set
    End Property

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

    Public Property Arguments() As String
        Get
            Return _arguments
        End Get
        Set(ByVal value As String)
            _arguments = value

            NotifyPropertyChanged("Arguments")
        End Set
    End Property

    Public Property CurrentCredentials() As Boolean
        Get
            Return _currentcredentials
        End Get
        Set(ByVal value As Boolean)
            _currentcredentials = value

            NotifyPropertyChanged("CurrentCredentials")
        End Set
    End Property

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
