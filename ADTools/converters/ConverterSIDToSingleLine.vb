
Public Class ConverterSIDToSingleLine
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Try
            If value Is Nothing Then Return "<null>"
            If TypeOf value Is Byte()() Then
                Dim vallist As New List(Of String)
                For i = 0 To CType(value, Byte()()).GetUpperBound(0)
                    Dim sid As New Security.Principal.SecurityIdentifier(CType(value(i), Byte()), 0)
                    vallist.Add(sid.ToString)
                Next
                Return Join(vallist.ToArray, "; ")
            ElseIf TypeOf value Is Byte() Then
                Dim sid As New Security.Principal.SecurityIdentifier(CType(value, Byte()), 0)
                Return sid.ToString
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
