Imports System.Text.RegularExpressions
Imports HandlebarsDotNet

Public Class ConverterToInlinesWithHyperlink
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        If value Is Nothing Then Return Nothing

        Dim spn = New Span()

        Dim regexpatterns As New List(Of String)
        For Each domain In domains
            For Each template In domain.UsernamePatternTemplates
                Dim data = New With {.n = "\d+"}
                regexpatterns.Add(template(data))
            Next
            For Each template In domain.ComputerPatternTemplates
                Dim data = New With {.n = "\d+"}
                regexpatterns.Add(template(data))
            Next
        Next

        Dim regexMatchSearcher As New Regex("(" & Join(regexpatterns.ToArray, ")|(") & ")", RegexOptions.IgnoreCase)
        Dim matches As MatchCollection = regexMatchSearcher.Matches(value.ToString)
        Dim lastindex As Integer = 0
        For Each match As Match In matches
            spn.Inlines.Add(New Run(value.ToString.Substring(lastindex, match.Index - lastindex)))
            lastindex = match.Index + match.Length
            Dim hl = New Hyperlink(New Run(match.Value)) With {.Tag = match.Value}
            AddHandler hl.Click, AddressOf Hyperlink_Click
            spn.Inlines.Add(hl)
        Next
        spn.Inlines.Add(New Run(value.ToString.Substring(lastindex)))

        Return spn.Inlines
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterToInlinesWithHyperlink is a OneWay converter.")
    End Function

    Public Sub Hyperlink_Click(sender As Object, e As RoutedEventArgs)
        Dim w As Window = Window.GetWindow(sender)
        If w Is Nothing Then Exit Sub
        If TypeOf w Is wndMain Then
            Dim wm = CType(w, wndMain)
            wm.StartSearch(Nothing, New clsFilter("""" & CType(sender, Hyperlink).Tag.ToString & """", attributesForSearchDefault, wm.searchobjectclasses))
        End If
    End Sub

End Class
