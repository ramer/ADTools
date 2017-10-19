
Public Class CustomTextBlock
    Inherits TextBlock

    Public Property InlineCollection() As InlineCollection
        Get
            Return DirectCast(GetValue(InlineCollectionProperty), InlineCollection)
        End Get
        Set
            SetValue(InlineCollectionProperty, Value)
        End Set
    End Property

    Public Shared ReadOnly InlineCollectionProperty As DependencyProperty = DependencyProperty.Register("InlineCollection",
                                                    GetType(InlineCollection),
                                                    GetType(CustomTextBlock),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf InlineCollectionPropertyChanged))

    Private Shared Sub InlineCollectionPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As CustomTextBlock = CType(d, CustomTextBlock)
        With instance
            .Inlines.Clear()
            Dim inlines = CType(e.NewValue, InlineCollection)
            If inlines IsNot Nothing Then .Inlines.AddRange(CType(e.NewValue, InlineCollection).ToList)
        End With
    End Sub
End Class
