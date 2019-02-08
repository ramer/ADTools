
Public Class ConverterDataRowViewToDataContext
    Implements IValueConverter

    Public Function IValueConverter_Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim cell As DataGridCell = value
        If cell Is Nothing Then Return Nothing

        Dim drv As System.Data.DataRowView = cell.DataContext
        If drv Is Nothing Then Return Nothing

        Return drv(cell.Column.SortMemberPath)
    End Function

    Public Function IValueConverter_ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterDataRowViewToDataContext is a OneWay converter.")
    End Function

End Class

