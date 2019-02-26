Public Class ConverterTreeViewItemDepthToMargin
    Implements IValueConverter

    Public Property Length As Double = 19

    Public Function Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim NewLenght As Double
        Double.TryParse(parameter, NewLenght)
        If NewLenght > 0 Then Length = NewLenght

        Dim item As TreeViewItem = value
        If item Is Nothing Then Return New Thickness(0)

        Dim depth As Integer = 0
        Dim parent As TreeViewItem = FindVisualParent(Of TreeViewItem)(item)
        While parent IsNot Nothing
            depth += 1
            parent = FindVisualParent(Of TreeViewItem)(parent)
        End While

        Return New Thickness(Length * depth, 0, 0, 0)
    End Function

    Public Function ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterTreeViewItemDepthToMargin is a OneWay converter.")
    End Function
End Class
