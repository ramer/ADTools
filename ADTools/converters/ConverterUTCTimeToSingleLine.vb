
Public Class ConverterUTCTimeToSingleLine
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Try
            If value Is Nothing Then Return "<null>"
            If IsArray(value) Then
                Dim arr = CType(value, Array)
                Dim vallist As New List(Of String)
                For i = 0 To arr.Length - 1
                    vallist.Add(Date.ParseExact(arr(i), "yyyyMMddHHmmss.f'Z'", Globalization.CultureInfo.InvariantCulture))
                Next
                Return Join(vallist.ToArray, "; ")
            End If
            Return Date.ParseExact(value, "yyyyMMddHHmmss.f'Z'", Globalization.CultureInfo.InvariantCulture)
        Catch
            Return "<error>"
        End Try
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterUTCTimeToSingleLine is a OneWay converter.")
    End Function
End Class
