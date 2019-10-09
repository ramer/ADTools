
Public Class ConverterOctetStringToSingleLine
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Try
            If value Is Nothing Then Return "<null>"
            If TypeOf value Is Byte()() Then
                Dim vallist As New List(Of String)
                For i = 0 To CType(value, Byte()()).GetUpperBound(0)
                    If CType(value(i), Byte()).Length = 16 Then
                        vallist.Add(New Guid(CType(value(i), Byte())).ToString)
                    Else
                        vallist.Add(Text.Encoding.UTF8.GetString(CType(value(i), Byte())))
                    End If
                Next
                Return Join(vallist.ToArray, "; ")
            ElseIf TypeOf value Is Byte() Then
                If CType(value, Byte()).Length = 16 Then
                    Return New Guid(CType(value, Byte())).ToString
                ElseIf CType(value, Byte()).Length > 1024 Then
                    Return "<binary>"
                Else
                    Return Text.Encoding.UTF8.GetString(CType(value, Byte()))
                End If
            End If
            Return value.ToString
        Catch ex As Exception
            Return "<error>"
        End Try
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterOctetStringToSingleLine is a OneWay converter.")
    End Function
End Class
