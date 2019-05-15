Public Class ConverterBooleanInverted
    Implements IValueConverter

    Public Function IValueConverter_Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        'If targetType <> GetType(Boolean) Then
        '    Throw New InvalidOperationException("The target must be a boolean")
        'End If

        Return Not CBool(value)
    End Function

    Public Function IValueConverter_ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        'If targetType <> GetType(Boolean) Then
        '    Throw New InvalidOperationException("The target must be a boolean")
        'End If

        Return Not CBool(value)
    End Function
End Class
