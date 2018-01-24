Class pgUserBasicInformation

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                        GetType(clsDirectoryObject),
                                                        GetType(pgUserBasicInformation),
                                                        New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgUserBasicInformation = CType(d, pgUserBasicInformation)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Sub New(obj As clsDirectoryObject)
        InitializeComponent()
        CurrentObject = obj
    End Sub

    Private Sub cmboTelephoneNumber_DropDownOpened(sender As Object, e As EventArgs) Handles cmboTelephoneNumber.DropDownOpened
        cmboTelephoneNumber.ItemsSource = GetNextDomainTelephoneNumbers(CurrentObject.Domain)
    End Sub

    Private Sub cmboTelephoneNumber_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmboTelephoneNumber.SelectionChanged
        e.Handled = True
    End Sub

    Private Sub btnResolveDepartment_Click(sender As Object, e As RoutedEventArgs) Handles btnResolveDepartment.Click
        If CurrentObject Is Nothing Then Exit Sub
        CurrentObject.department = Replace(CurrentObject.Parent.name, "OU=", "")
    End Sub

End Class
