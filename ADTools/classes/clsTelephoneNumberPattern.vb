Imports System.ComponentModel

Public Class clsTelephoneNumberPattern
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _label As String
    Private _pattern As String
    Private _range As String

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New(Label As String, Pattern As String, Range As String)
        _label = Label
        _pattern = Pattern
        _range = Range
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

    Public Property Pattern() As String
        Get
            Return _pattern
        End Get
        Set(ByVal value As String)
            _pattern = value
            NotifyPropertyChanged("Pattern")
        End Set
    End Property

    Public Property Range() As String
        Get
            Return _range
        End Get
        Set(ByVal value As String)
            _range = value
            NotifyPropertyChanged("Range")
        End Set
    End Property

End Class
