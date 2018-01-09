Imports System.Collections.ObjectModel
Imports System.Reflection
Imports System.Windows.Controls.Primitives
Imports IPrompt.VisualBasic
Imports IPrint
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.ComponentModel

Class pgMain

    Public Shared hkF5 As New RoutedCommand
    Public Shared hkF1 As New RoutedCommand
    Public Shared hkEsc As New RoutedCommand

    'Public Property currentcontainer As clsDirectoryObject
    'Public Property currentfilter As clsFilter
    'Public Property currentobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Private searchhistoryindex As Integer
    Private searchhistory As New List(Of clsSearchHistory)


    Public WithEvents clipboardTimer As New Threading.DispatcherTimer()

    Private clipboardlastdata As String

    Private Sub wndMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        'hkF5.InputGestures.Add(New KeyGesture(Key.F5))
        'Me.CommandBindings.Add(New CommandBinding(hkF5, AddressOf RefreshSearchResults))
        'hkEsc.InputGestures.Add(New KeyGesture(Key.Escape))
        'Me.CommandBindings.Add(New CommandBinding(hkEsc, AddressOf StopSearch))

        'dpToolbar.DataContext = preferences
        'mnuSearchDomains.ItemsSource = domains
        'dockpSearchObjectClasses.DataContext = searchobjectclasses

        'DataObject.AddPastingHandler(tbSearchPattern, AddressOf tbSearchPattern_OnPaste)

        'clipboardTimer.Interval = New TimeSpan(0, 0, 1)
        'clipboardTimer.Start()
    End Sub


    '#Region "Popups"

    '    Private Sub ShowPopups()
    '        poptvObjects.IsOpen = True
    '        poptbSearchPattern.IsOpen = True
    '        popF1Hint.IsOpen = True
    '    End Sub

    '    Private Sub MovePopups() Handles Me.LocationChanged, Me.SizeChanged, tvObjects.SizeChanged, tbSearchPattern.SizeChanged
    '        poptvObjects.HorizontalOffset += 1 : poptvObjects.HorizontalOffset -= 1
    '        poptbSearchPattern.HorizontalOffset += 1 : poptbSearchPattern.HorizontalOffset -= 1
    '        popF1Hint.HorizontalOffset += 1 : popF1Hint.HorizontalOffset -= 1
    '    End Sub

    '    Private Sub ClosePopups()
    '        poptvObjects.IsOpen = False
    '        poptbSearchPattern.IsOpen = False
    '        popF1Hint.IsOpen = False
    '    End Sub

    '    Private Sub ClosePopup(sender As Object, e As MouseButtonEventArgs) Handles poptvObjects.MouseLeftButtonDown, poptbSearchPattern.MouseLeftButtonDown, popF1Hint.MouseLeftButtonDown
    '        CType(sender, Popup).IsOpen = False
    '    End Sub

    '#End Region

#Region "Main Menu"

    Private Sub mnuFilePrint_Click(sender As Object, e As RoutedEventArgs) Handles mnuFilePrint.Click
        'Dim fd As New FlowDocument
        'fd.IsColumnWidthFlexible = True

        'Dim table = New Table()
        'table.CellSpacing = 0
        'table.BorderBrush = Brushes.Black
        'table.BorderThickness = New Thickness(1)
        'Dim rowGroup = New TableRowGroup()
        'table.RowGroups.Add(rowGroup)
        'Dim header = New TableRow()
        'header.Background = Brushes.AliceBlue
        'rowGroup.Rows.Add(header)

        'For Each column As clsViewColumnInfo In preferences.Columns
        '    Dim tableColumn = New TableColumn()
        '    'configure width and such
        '    tableColumn.Width = New GridLength(column.Width / 96, GridUnitType.Star)
        '    table.Columns.Add(tableColumn)
        '    Dim hc As New Paragraph(New Run(column.Header))
        '    hc.FontSize = 10.0
        '    hc.FontFamily = New FontFamily("Segoe UI")
        '    hc.FontWeight = FontWeights.Bold
        '    Dim cell = New TableCell(hc)
        '    cell.BorderBrush = Brushes.Gray
        '    cell.BorderThickness = New Thickness(0.1)
        '    cell.Padding = New Thickness(5, 5, 5, 5)
        '    header.Cells.Add(cell)
        'Next

        'For Each obj In currentobjects
        '    Dim tableRow = New TableRow()
        '    rowGroup.Rows.Add(tableRow)

        '    For Each column As clsViewColumnInfo In preferences.Columns
        '        Dim cell As New TableCell
        '        cell.BorderBrush = Brushes.Gray
        '        cell.BorderThickness = New Thickness(0.1)
        '        cell.Padding = New Thickness(5, 5, 5, 5)
        '        Dim first As Boolean = True
        '        For Each attr In column.Attributes
        '            Dim t As Type = obj.GetType()
        '            Dim pic() As PropertyInfo = t.GetProperties()

        '            For Each pi In pic
        '                If pi.Name = attr.Name Then
        '                    Dim value = pi.GetValue(obj)

        '                    If TypeOf value Is String Then
        '                        Dim p As New Paragraph(New Run(value))
        '                        p.FontSize = 8.0
        '                        p.FontFamily = New FontFamily("Segoe UI")
        '                        If first Then p.FontWeight = FontWeights.Bold : first = False
        '                        cell.Blocks.Add(p)
        '                    ElseIf TypeOf value Is BitmapImage Then
        '                        Dim img As New Image
        '                        img.Source = value
        '                        img.Width = 16
        '                        img.Height = 16
        '                        Dim p As New BlockUIContainer(img)
        '                        cell.Blocks.Add(p)
        '                    End If
        '                End If
        '            Next

        '        Next

        '        tableRow.Cells.Add(cell)
        '    Next
        'Next

        'fd.Blocks.Add(table)
        'Try
        '    IPrintDialog.PreviewDocument(fd)
        'Catch ex As Exception
        '    Debug.Print(ex.Message)
        'End Try
    End Sub

    Private Sub mnuEditCreateObject_Click(sender As Object, e As RoutedEventArgs) Handles mnuEditCreateObject.Click
        Dim w As New pgCreateObject

        'If currentcontainer IsNot Nothing Then
        '    w.destinationcontainer = currentcontainer
        '    w.destinationdomain = currentcontainer.Domain
        'Else
        '    w.destinationcontainer = Nothing
        '    w.destinationdomain = Nothing
        'End If

        ShowPage(w, False, Window.GetWindow(Me), False)
    End Sub

    Private Sub mnuServiceDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceDomainOptions.Click 'TODO, ctxmnutviDomainsDomainOptions.Click
        ShowPage(New pgDomains, True, Window.GetWindow(Me), True)
        'TODO RefreshDomainTree()
    End Sub

    Private Sub mnuServicePreferences_Click(sender As Object, e As RoutedEventArgs) Handles mnuServicePreferences.Click
        ShowPage(New pgPreferences, True, Window.GetWindow(Me), True)
        'TODO RefreshDomainTree()
    End Sub

    Private Sub mnuServiceLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceLog.Click
        Dim w As New wndLog
        'TODO   ShowWindow(w, True, Nothing, False)
    End Sub

    Private Sub mnuServiceErrorLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceErrorLog.Click
        Dim w As New wndErrorLog
        'TODO   ShowWindow(w, True, Nothing, False)
    End Sub

    Private Sub mnuSearchSaveCurrentFilter_Click(sender As Object, e As RoutedEventArgs) Handles mnuSearchSaveCurrentFilter.Click
        'If currentfilter Is Nothing OrElse String.IsNullOrEmpty(currentfilter.Filter) Then IMsgBox(My.Resources.wndMain_msg_CannotSaveCurrentFilter, vbOKOnly + vbExclamation,, Window.GetWindow(Me)) : Exit Sub

        'Dim name As String = IInputBox(My.Resources.wndMain_msg_EnterFilterName,,, vbQuestion, Window.GetWindow(Me))

        'If String.IsNullOrEmpty(name) Then Exit Sub

        'currentfilter.Name = name
        'preferences.Filters.Add(currentfilter)
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As RoutedEventArgs) Handles mnuHelpAbout.Click
        Dim p As New pgAbout
        ShowPage(p, True, Window.GetWindow(Me), False)
    End Sub

#End Region

#Region "Context Menu"

#End Region

#Region "Events"

    'Private Sub clipboardTimer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles clipboardTimer.Tick
    '    If preferences Is Nothing OrElse preferences.ClipboardSource = False Then Exit Sub
    '    Dim newclipboarddata As String = Clipboard.GetText
    '    If String.IsNullOrEmpty(newclipboarddata) Then Exit Sub
    '    If preferences.ClipboardSourceLimit AndAlso CountWords(newclipboarddata) > 3 Then Exit Sub ' only three words

    '    If clipboardlastdata <> newclipboarddata Then
    '        clipboardlastdata = newclipboarddata
    '        StartSearch(Nothing, New clsFilter(clipboardlastdata, preferences.AttributesForSearch, searchobjectclasses))
    '    End If
    'End Sub


    Private Sub btnWindowClone_Click(sender As Object, e As RoutedEventArgs) Handles btnWindowClone.Click
        Dim w As New wndMain
        w.Show()
    End Sub

    Private Sub btnDummy_Click(sender As Object, e As RoutedEventArgs) Handles btnDummy.Click

    End Sub

#End Region

#Region "Subs"

#End Region

End Class
