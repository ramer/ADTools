Imports System.Globalization

Public NotInheritable Class ConverterDesiredWidthToVisibility
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        If Not IsNumeric(parameter) Then Throw New ArgumentException
        If CType(value, Double) > parameter Then
            Return Visibility.Visible
        Else
            Return Visibility.Collapsed
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterDesiredWidthToVisibility is a OneWay converter.")
    End Function


End Class
