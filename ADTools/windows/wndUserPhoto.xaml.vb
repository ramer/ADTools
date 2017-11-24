Public Class wndUserPhoto
    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(wndUserPhoto),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Public Property _currentobject As clsDirectoryObject

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As wndUserPhoto = CType(d, wndUserPhoto)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Sub imgPhoto_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgPhoto.MouseDown
        Me.Close()
    End Sub

    Private Sub wndUserPhoto_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDown
        Me.Close()
    End Sub

End Class
