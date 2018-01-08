Imports System.ComponentModel

Public Class pgContact
    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(pgContact),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomainobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgContact = CType(d, pgContact)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Sub cmboTelephoneNumber_DropDownOpened(sender As Object, e As EventArgs) Handles cmboTelephoneNumber.DropDownOpened
        cmboTelephoneNumber.ItemsSource = GetNextDomainTelephoneNumbers(currentobject.Domain)
    End Sub

    Private Sub Manager_hyperlink_click(sender As Object, e As RequestNavigateEventArgs)
        ShowDirectoryObjectProperties(CurrentObject.manager, Window.GetWindow(Me))
    End Sub

    Private Sub pgContact_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
    End Sub

End Class
