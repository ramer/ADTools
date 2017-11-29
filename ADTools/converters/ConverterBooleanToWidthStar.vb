Imports System.Globalization

Public NotInheritable Class ConverterBooleanToWidthStar
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim flag = False
        If TypeOf value Is Boolean Then
            flag = CBool(value)
        ElseIf TypeOf value Is System.Nullable(Of Boolean) Then
            Dim nullable = DirectCast(value, System.Nullable(Of Boolean))
            flag = nullable.GetValueOrDefault()
        End If
        If parameter IsNot Nothing Then
            If Boolean.Parse(DirectCast(parameter, String)) Then
                flag = Not flag
            End If
        End If
        If flag Then
            Return New GridLength(1, GridUnitType.Star)
        Else
            Return New GridLength(0, GridUnitType.Star)
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Dim back = ((TypeOf value Is GridLength) AndAlso (DirectCast(value, GridLength).Value > 0))
        If parameter IsNot Nothing Then
            If CBool(parameter) Then
                back = Not back
            End If
        End If
        Return back
    End Function


End Class