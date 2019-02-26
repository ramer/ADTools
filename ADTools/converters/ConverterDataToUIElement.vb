Imports System.Collections.ObjectModel
Imports System.Text.RegularExpressions

Public Class ConverterDataToUIElement
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        If value Is Nothing Then Return Nothing
        Dim content As Object

        Select Case value.GetType()
            Case GetType(Grid)

                Dim grid = CType(value, Grid)
                content = grid

            Case GetType(BitmapImage)

                Dim img As New Image
                img.Source = value
                content = img

            Case GetType(clsDirectoryObject)

                Dim obj As clsDirectoryObject = CType(value, clsDirectoryObject)
                Dim tblck As New TextBlock()
                Dim rn = New Run(obj.name)
                rn.Background = Brushes.Transparent
                rn.TextDecorations = TextDecorations.Underline
                AddHandler rn.MouseLeftButtonDown,
                    Sub()
                        ShowDirectoryObjectProperties(obj, Window.GetWindow(rn))
                    End Sub
                tblck.Inlines.Add(rn)
                content = tblck

            Case GetType(String)

                Dim tblck As New TextBlock()

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

                If regexpatterns.Count > 0 Then

                    Dim regexMatchSearcher As New Regex("(" & Join(regexpatterns.ToArray, ")|(") & ")", RegexOptions.IgnoreCase)
                    Dim matches As MatchCollection = regexMatchSearcher.Matches(value.ToString)
                    Dim lastindex As Integer = 0
                    For Each match As Match In matches
                        tblck.Inlines.Add(New Run(value.ToString.Substring(lastindex, match.Index - lastindex)))
                        lastindex = match.Index + match.Length
                        Dim rn = New Run(match.Value)
                        rn.Background = Brushes.Transparent
                        rn.TextDecorations = TextDecorations.Underline
                        Dim hl As New Hyperlink(rn)
                        AddHandler hl.Click,
                    Sub()
                        Dim w As Window = Window.GetWindow(hl)
                        If w IsNot Nothing AndAlso
                        TypeOf w Is NavigationWindow AndAlso
                        CType(w, NavigationWindow).Content IsNot Nothing AndAlso
                        TypeOf CType(w, NavigationWindow).Content Is pgMain AndAlso
                        TypeOf CType(CType(w, NavigationWindow).Content, pgMain).frmObjects.Content Is pgObjects Then
                            CType(CType(CType(w, NavigationWindow).Content, pgMain).frmObjects.Content, pgObjects).StartSearch(Nothing, New clsFilter("""" & match.Value & """", attributesForSearchDefault, preferences.SearchObjectClasses))
                        End If
                    End Sub
                        tblck.Inlines.Add(hl)
                    Next
                    tblck.Inlines.Add(New Run(value.ToString.Substring(lastindex)))
                    tblck.TextWrapping = TextWrapping.Wrap
                    content = tblck

                Else

                    tblck.Text = value
                    tblck.TextWrapping = TextWrapping.Wrap
                    content = tblck

                End If

            Case Else

                Dim tblck As New TextBlock()
                tblck.Text = value.ToString
                tblck.TextWrapping = TextWrapping.Wrap
                content = tblck

        End Select

        Return content
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterDataToUIElement is a OneWay converter.")
    End Function

End Class
