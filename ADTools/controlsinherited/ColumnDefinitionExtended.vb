Imports System.ComponentModel

Public Class ColumnDefinitionExtended
    Inherits ColumnDefinition

    Dim _threshold As Double = 0

    Dim widthDecription As DependencyPropertyDescriptor = Nothing

    Public Shared ReadOnly VisibleProperty As DependencyProperty = DependencyProperty.Register("Visible",
                                                    GetType(Boolean),
                                                    GetType(ColumnDefinitionExtended),
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
        widthDecription = DependencyPropertyDescriptor.FromProperty(WidthProperty, GetType(ColumnDefinition))
        widthDecription.AddValueChanged(Me, AddressOf WidthPropertyChanged)
    End Sub

    Private Sub ColumnDefinitionExtended_Unloaded(sender As Object, e As RoutedEventArgs) Handles Me.Unloaded
        widthDecription.RemoveValueChanged(Me, AddressOf WidthPropertyChanged)
    End Sub

    Private Shared Sub VisiblePropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ColumnDefinitionExtended = CType(d, ColumnDefinitionExtended)
        With instance
            If CType(e.NewValue, Boolean) = True Then
                If .Width.Value <= ._threshold Then .Width = New GridLength(250, .Width.GridUnitType)
            Else
                If .Width.Value > ._threshold Then .Width = New GridLength(0, .Width.GridUnitType)
            End If
        End With
    End Sub

    Private Sub WidthPropertyChanged(sender As Object, e As EventArgs)
        If Width.Value <= _threshold Then
            If Visible = True Then Visible = False
        Else
            If Visible = False Then Visible = True
        End If
    End Sub

End Class