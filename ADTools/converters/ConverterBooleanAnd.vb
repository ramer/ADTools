Imports System.Globalization

Public Class ConverterBooleanAnd
    Implements IMultiValueConverter

    Private Function IMultiValueConverter_Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
        For Each value As Object In values
            If CBool(value) = False Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Function IMultiValueConverter_ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterBooleanAnd is a OneWay converter.")
    End Function
End Class
