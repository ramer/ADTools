Imports System.ComponentModel

Public Class pgOrganizationalUnit
    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(pgOrganizationalUnit),
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
        Dim instance As pgOrganizationalUnit = CType(d, pgOrganizationalUnit)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Sub hlManagedBy_Click(sender As Object, e As RoutedEventArgs) Handles hlManagedBy.Click
        ShowDirectoryObjectProperties(CurrentObject.managedBy, Window.GetWindow(Me))
    End Sub

    Private Sub pgOrganizationalUnit_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
    End Sub

End Class
