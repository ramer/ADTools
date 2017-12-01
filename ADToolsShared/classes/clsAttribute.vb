Imports System.ComponentModel
Imports IRegisty

Public Class clsAttribute
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _name As String
    Private _label As String
    Private _value As Object
    Private _newvalue As Object

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New()

    End Sub

    Sub New(Name As String,
            Label As String,
            Optional Value As Object = Nothing,
            Optional NewValue As Object = Nothing)

        _name = Name
        _label = Label
        _value = Value
        _newvalue = NewValue
    End Sub

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
            NotifyPropertyChanged("Name")
        End Set
    End Property

    Public Property Label() As String
        Get
            Return _label
        End Get
        Set(ByVal value As String)
            _label = value
            NotifyPropertyChanged("Label")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property Value() As Object
        Get
            Return _value
        End Get
        Set(ByVal value As Object)
            _value = value
            NotifyPropertyChanged("Value")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property NewValue() As Object
        Get
            Return _newvalue
        End Get
        Set(ByVal value As Object)
            _newvalue = value
            NotifyPropertyChanged("NewValue")
        End Set
    End Property

End Class
