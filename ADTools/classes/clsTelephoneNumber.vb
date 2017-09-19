Imports System.ComponentModel

Public Class clsTelephoneNumber
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _label As String
    Private _telephonenumber As String

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New(Label As String, TelephoneNumber As String)
        _label = Label
        _telephonenumber = TelephoneNumber
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

    Public Property TelephoneNumber() As String
        Get
            Return _telephonenumber
        End Get
        Set(ByVal value As String)
            _telephonenumber = value

            NotifyPropertyChanged("TelephoneNumber")
        End Set
    End Property

End Class
