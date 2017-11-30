
Public Class TabControlExtended
    Inherits TabControl

    Dim baseheight As Double

    Public Shared ReadOnly VisibleProperty As DependencyProperty = DependencyProperty.Register("Visible",
                                                    GetType(Boolean),
                                                    GetType(TabControlExtended),
                                                    New PropertyMetadata(True, AddressOf VisiblePropertyChanged))

    Public Property Visible() As Boolean
        Get
            Return GetValue(VisibleProperty)
        End Get
        Set
            SetCurrentValue(VisibleProperty, Value)
        End Set
    End Property

    Sub New()
    End Sub

    Private Shared Sub VisiblePropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As TabControlExtended = CType(d, TabControlExtended)
        With instance
            If CType(e.NewValue, Boolean) = True Then
                If .Height <= .MinHeight Then .Height = .baseheight
            Else
                If .Height > .MinHeight Then .Height = .MinHeight
            End If
        End With
    End Sub

    Private Sub TabControlExtended_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDoubleClick
        Visible = Not Visible
    End Sub

    Private Sub TabControlExtended_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        baseheight = If(Height > 0, Height, 100)
    End Sub
End Class