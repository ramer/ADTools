Class pgUserOrganization

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(pgUserOrganization),
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
        Dim instance As pgUserOrganization = CType(d, pgUserOrganization)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Sub New(obj As clsDirectoryObject)
        InitializeComponent()
        CurrentObject = obj
    End Sub

    Private Sub hlManager_Click(sender As Object, e As RoutedEventArgs) Handles hlManager.Click
        ShowDirectoryObjectProperties(CurrentObject.manager, Nothing)
    End Sub

End Class
