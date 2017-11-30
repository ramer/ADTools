Imports System.Globalization

Public NotInheritable Class ConverterIntegerToVisibility
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        If TypeOf value Is Integer AndAlso value > 0 Then
            Return Visibility.Visible
        Else
            Return Visibility.Collapsed
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterIntegerToVisibility is a OneWay converter.")
    End Function


End Class
