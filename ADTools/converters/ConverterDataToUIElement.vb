Imports System.Collections.ObjectModel
Imports System.Text.RegularExpressions

Public Class ConverterDataToUIElement
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        If value Is Nothing Then Return Nothing
        Dim items As New ObservableCollection(Of Object)

        Select Case value.GetType()
            Case GetType(Grid)
                items.Add(value)
            Case GetType(BitmapImage)
                Dim img As New Image
                img.Source = value
                items.Add(img)
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
                items.Add(tblck)

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
                        If w Is Nothing Then Exit Sub
                        If TypeOf w Is wndMain Then
                            Dim wm = CType(w, wndMain)
                            wm.StartSearch(Nothing, New clsFilter("""" & match.Value & """", attributesForSearchDefault, wm.searchobjectclasses))
                        End If
                    End Sub
                    tblck.Inlines.Add(hl)
                Next
                tblck.Inlines.Add(New Run(value.ToString.Substring(lastindex)))
                tblck.TextWrapping = TextWrapping.Wrap
                items.Add(tblck)


                '    valuetyped = value
                'Case GetType(Byte())
                '    valuetyped = value
                'Case GetType(Object())
                '    valuetyped = value
                'Case GetType(DateTime)
                '    valuetyped = value
                'Case GetType(Boolean)
                '    valuetyped = value
        End Select

        Return items
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException("ConverterToInlinesWithHyperlink is a OneWay converter.")
    End Function

End Class



'"

'            If attr.Name <> "Image" Then

'                Dim text As New FrameworkElementFactory(GetType(CustomTextBlock))
'                If first Then
'                    text.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold)
'                    first = False
'                    column.SetValue(DataGridColumn.SortMemberPathProperty, attr.Name)
'                End If
'                bind.Converter = New ConverterToInlinesWithHyperlink
'                text.SetBinding(CustomTextBlock.InlineCollectionProperty, bind)
'                text.SetValue(TextBlock.ToolTipProperty, attr.Label)
'                panel.AppendChild(text)

'            Else

'                Dim ttbind As New System.Windows.Data.Binding("Status")
'                ttbind.Mode = BindingMode.OneWay
'                Dim img As New FrameworkElementFactory(GetType(Image))
'                column.SetValue(clsSorter.PropertyNameProperty, "Image")
'                img.SetBinding(Image.SourceProperty, bind)
'                img.SetValue(Image.WidthProperty, 32.0)
'                img.SetValue(Image.HeightProperty, 32.0)
'                img.SetBinding(Image.ToolTipProperty, ttbind)
'                panel.AppendChild(img)

'            End If
'"
