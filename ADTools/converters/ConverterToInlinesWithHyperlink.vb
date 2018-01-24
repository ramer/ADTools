'Imports System.Text.RegularExpressions
'Imports HandlebarsDotNet

'Public Class ConverterToInlinesWithHyperlink
'    Implements IValueConverter

'    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
'        If value Is Nothing Then Return Nothing
'        Dim spn = New Span()
'        Dim inlines As New List(Of Inline)
'        'inlines.Add(New Run(value.ToString))
'        'spn.Inlines.AddRange(inlines)
'        'Return spn.Inlines

'        Dim regexpatterns As New List(Of String)
'        For Each domain In domains
'            For Each template In domain.UsernamePatternTemplates
'                Dim data = New With {.n = "\d+"}
'                regexpatterns.Add(template(data))
'            Next
'            For Each template In domain.ComputerPatternTemplates
'                Dim data = New With {.n = "\d+"}
'                regexpatterns.Add(template(data))
'            Next
'        Next

'        Dim regexMatchSearcher As New Regex("(" & Join(regexpatterns.ToArray, ")|(") & ")", RegexOptions.IgnoreCase)
'        Dim matches As MatchCollection = regexMatchSearcher.Matches(value.ToString)
'        Dim lastindex As Integer = 0
'        For Each match As Match In matches
'            inlines.Add(New Run(value.ToString.Substring(lastindex, match.Index - lastindex)))
'            lastindex = match.Index + match.Length
'            Dim hl = New Hyperlink(New Run(match.Value)) With {.Tag = match.Value}
'            AddHandler hl.Click,
'                Sub()
'                    Dim w As Window = Window.GetWindow(hl)
'                    If w IsNot Nothing AndAlso
'                    TypeOf w Is NavigationWindow AndAlso
'                    CType(w, NavigationWindow).Content IsNot Nothing AndAlso
'                    TypeOf CType(w, NavigationWindow).Content Is pgMain AndAlso
'                    TypeOf CType(CType(w, NavigationWindow).Content, pgMain).frmObjects.Content Is pgObjects Then
'                        CType(CType(CType(w, NavigationWindow).Content, pgMain).frmObjects.Content, pgObjects).StartSearch(Nothing, New clsFilter("""" & match.Value & """", attributesForSearchDefault, preferences.SearchObjectClasses))
'                    End If
'                End Sub
'            inlines.Add(hl)
'        Next
'        inlines.Add(New Run(value.ToString.Substring(lastindex)))

'        spn.Inlines.AddRange(inlines)
'        Return spn.Inlines
'    End Function

'    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
'        Throw New NotSupportedException("ConverterToInlinesWithHyperlink is a OneWay converter.")
'    End Function

'End Class
