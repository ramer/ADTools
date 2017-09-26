
Public Class ConverterToSingleLine
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        If value Is Nothing Then Return ""
        If IsArray(value) Then
            Dim arr = CType(value, Array)
            Dim vallist As New List(Of String)
            For i = 0 To arr.Length - 1
                vallist.Add(arr(i).ToString)
            Next
            value = Join(vallist.ToArray, " ")
        End If
        Return value.ToString.Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbCrLf, " ")
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterToSingleLine is a OneWay converter.")
    End Function

End Class
