
Public Class ConverterEnumToVisibility
    Implements IValueConverter

    Public Function Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim parameterString As String = TryCast(parameter, String)
        If parameterString Is Nothing Then Return DependencyProperty.UnsetValue
        If [Enum].IsDefined(value.[GetType](), value) = False Then Return DependencyProperty.UnsetValue
        Dim parameterValue As Object = [Enum].Parse(value.[GetType](), parameterString)
        Return If(parameterValue.Equals(value), Visibility.Visible, Visibility.Collapsed)
    End Function

    Public Function ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Dim parameterString As String = TryCast(parameter, String)
        If parameterString Is Nothing Then Return DependencyProperty.UnsetValue
        Return [Enum].Parse(targetType, parameterString)
    End Function

End Class