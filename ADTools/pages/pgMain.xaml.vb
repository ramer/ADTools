Imports System.Reflection
Imports IPrint

Class pgMain

    Private Sub pgMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dpToolbar.DataContext = preferences
        mnuSearchDomains.ItemsSource = domains
    End Sub

#Region "Main Menu"

    Private Sub mnuFilePrint_Click(sender As Object, e As RoutedEventArgs) Handles mnuFilePrint.Click
        Dim currentobjects As clsThreadSafeObservableCollection(Of clsDirectoryObject) = Nothing
        If frmObjects IsNot Nothing AndAlso frmObjects.Content IsNot Nothing AndAlso TypeOf frmObjects.Content Is pgObjects Then currentobjects = CType(frmObjects.Content, pgObjects).currentobjects
        If currentobjects Is Nothing Then Exit Sub

        Dim fd As New FlowDocument
        fd.IsColumnWidthFlexible = True

        Dim table = New Table()
        table.CellSpacing = 0
        table.BorderBrush = Brushes.Black
        table.BorderThickness = New Thickness(1)
        Dim rowGroup = New TableRowGroup()
        table.RowGroups.Add(rowGroup)
        Dim header = New TableRow()
        header.Background = Brushes.AliceBlue
        rowGroup.Rows.Add(header)

        For Each column As clsViewColumnInfo In preferences.Columns
            Dim tableColumn = New TableColumn()
            'configure width and such
            tableColumn.Width = New GridLength(column.Width / 96, GridUnitType.Star)
            table.Columns.Add(tableColumn)
            Dim hc As New Paragraph(New Run(column.Header))
            hc.FontSize = 10.0
            hc.FontFamily = New FontFamily("Segoe UI")
            hc.FontWeight = FontWeights.Bold
            Dim cell = New TableCell(hc)
            cell.BorderBrush = Brushes.Gray
            cell.BorderThickness = New Thickness(0.1)
            cell.Padding = New Thickness(5, 5, 5, 5)
            header.Cells.Add(cell)
        Next

        For Each obj In currentobjects
            Dim tableRow = New TableRow()
            rowGroup.Rows.Add(tableRow)

            For Each column As clsViewColumnInfo In preferences.Columns
                Dim cell As New TableCell
                cell.BorderBrush = Brushes.Gray
                cell.BorderThickness = New Thickness(0.1)
                cell.Padding = New Thickness(5, 5, 5, 5)
                Dim first As Boolean = True
                For Each attr In column.Attributes
                    Dim t As Type = obj.GetType()
                    Dim pic() As PropertyInfo = t.GetProperties()

                    For Each pi In pic
                        If pi.Name = attr.Name Then
                            Dim value = pi.GetValue(obj)

                            If TypeOf value Is String Then
                                Dim p As New Paragraph(New Run(value))
                                p.FontSize = 8.0
                                p.FontFamily = New FontFamily("Segoe UI")
                                If first Then p.FontWeight = FontWeights.Bold : first = False
                                cell.Blocks.Add(p)
                            ElseIf TypeOf value Is BitmapImage Then
                                Dim img As New Image
                                img.Source = value
                                img.Width = 16
                                img.Height = 16
                                Dim p As New BlockUIContainer(img)
                                cell.Blocks.Add(p)
                            End If
                        End If
                    Next

                Next

                tableRow.Cells.Add(cell)
            Next
        Next

        fd.Blocks.Add(table)
        Try
            IPrintDialog.PreviewDocument(fd)
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub

    Private Sub mnuEditCreateObject_Click(sender As Object, e As RoutedEventArgs) Handles mnuEditCreateObject.Click
        ShowPage(New pgCreateObject(Nothing, Nothing), False, Window.GetWindow(Me), False)
    End Sub

    Private Sub mnuServiceDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceDomainOptions.Click
        ShowPage(New pgDomains, True, Window.GetWindow(Me), True)
        If frmObjects IsNot Nothing AndAlso frmObjects.Content IsNot Nothing AndAlso TypeOf frmObjects.Content Is pgObjects Then CType(frmObjects.Content, pgObjects).RefreshDomainTree()
    End Sub

    Private Sub mnuServicePreferences_Click(sender As Object, e As RoutedEventArgs) Handles mnuServicePreferences.Click
        ShowPage(New pgPreferences, True, Window.GetWindow(Me), True)
        If frmObjects IsNot Nothing AndAlso frmObjects.Content IsNot Nothing AndAlso TypeOf frmObjects.Content Is pgObjects Then CType(frmObjects.Content, pgObjects).RefreshDomainTree()
    End Sub

    Private Sub mnuServiceLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceLog.Click
        Dim w As New wndLog
        w.Owner = Window.GetWindow(Me)
        w.Show()
    End Sub

    Private Sub mnuServiceErrorLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceErrorLog.Click
        Dim w As New wndErrorLog
        w.Owner = Window.GetWindow(Me)
        w.Show()
    End Sub

    Private Sub mnuSearchSaveCurrentFilter_Click(sender As Object, e As RoutedEventArgs) Handles mnuSearchSaveCurrentFilter.Click
        If frmObjects IsNot Nothing AndAlso frmObjects.Content IsNot Nothing AndAlso TypeOf frmObjects.Content Is pgObjects Then CType(frmObjects.Content, pgObjects).SaveCurrentFilter()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As RoutedEventArgs) Handles mnuHelpAbout.Click
        Dim p As New pgAbout
        ShowPage(p, True, Window.GetWindow(Me), False)
    End Sub

#End Region

#Region "Events"

    Private Sub btnWindowClone_Click(sender As Object, e As RoutedEventArgs) Handles btnWindowClone.Click
        ADToolsApplication.ShowMainWindow()
    End Sub

    Private Sub btnDummy_Click(sender As Object, e As RoutedEventArgs) Handles btnDummy.Click

    End Sub

#End Region


End Class
