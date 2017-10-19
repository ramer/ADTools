Imports System.Collections.ObjectModel
Imports System.Reflection
Imports System.Windows.Controls.Primitives
Imports IPrompt.VisualBasic
Imports IPrint
Imports System.Text.RegularExpressions
Imports System.IO

Class wndMain
    Public WithEvents searcher As New clsSearcher

    Public Shared hkF5 As New RoutedCommand

    Public Property currentcontainer As clsDirectoryObject
    Public Property currentobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)
    Public Property currentfilter As clsFilter

    Private searchhistoryindex As Integer
    Private searchhistory As New List(Of clsSearchHistory)

    Public Property searchobjectclasses As New clsSearchObjectClasses(True, True, True, True, False)

    Private Sub wndMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        hkF5.InputGestures.Add(New KeyGesture(Key.F5))
        Me.CommandBindings.Add(New CommandBinding(hkF5, AddressOf HotKey_F5))

        DomainTreeUpdate()
        RebuildColumns()

        cmboSearchPattern.ItemsSource = ADToolsApplication.ocGlobalSearchHistory
        cmboSearchPattern.Focus()

        mnuSearchDomains.ItemsSource = domains
        tviFavorites.ItemsSource = preferences.Favorites
        tviFilters.ItemsSource = preferences.Filters
    End Sub

    Private Sub wndMain_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles MyBase.Closing, MyBase.Closing
        Dim count As Integer = 0

        For Each wnd As Window In ADToolsApplication.Current.Windows
            If GetType(wndMain) Is wnd.GetType Then count += 1
        Next

        If preferences.CloseOnXButton AndAlso count <= 1 Then Application.Current.Shutdown()
    End Sub

#Region "Main Menu"


    Private Sub mnuFilePrint_Click(sender As Object, e As RoutedEventArgs) Handles mnuFilePrint.Click
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

        For Each column As clsDataGridColumnInfo In preferences.Columns
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

            For Each column As clsDataGridColumnInfo In preferences.Columns
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

                            If attr.Name <> "Image" Then
                                Dim p As New Paragraph(New Run(value))
                                p.FontSize = 8.0
                                p.FontFamily = New FontFamily("Segoe UI")
                                If first Then p.FontWeight = FontWeights.Bold : first = False
                                cell.Blocks.Add(p)
                            Else
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
        IPrintDialog.PreviewDocument(fd)
    End Sub

    Private Sub mnuServiceDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceDomainOptions.Click
        ShowWindow(New wndDomains, True, Me, True)
        DomainTreeUpdate()
    End Sub

    Private Sub mnuServicePreferences_Click(sender As Object, e As RoutedEventArgs) Handles mnuServicePreferences.Click
        ShowWindow(New wndPreferences, True, Me, True)
    End Sub

    Private Sub mnuSearchSaveCurrentFilter_Click(sender As Object, e As RoutedEventArgs) Handles mnuSearchSaveCurrentFilter.Click
        If currentfilter Is Nothing OrElse String.IsNullOrEmpty(currentfilter.Filter) Then IMsgBox(My.Resources.wndMain_msg_CannotSaveCurrentFilter, vbOKOnly + vbExclamation,, Me) : Exit Sub

        Dim name As String = IInputBox(My.Resources.wndMain_msg_EnterFilterName,,, vbQuestion, Me)

        If String.IsNullOrEmpty(name) Then Exit Sub

        currentfilter.Name = name
        preferences.Filters.Add(currentfilter)
    End Sub

#End Region

#Region "Context Menu"

    Private Sub tviFilters_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)
        If TypeOf sp.Tag Is clsFilter Then
            StartSearch(Nothing, CType(sp.Tag, clsFilter))
        End If
    End Sub

    Private Sub tviFilters_TreeViewItem_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
        Dim flt As clsFilter = Nothing
        If TypeOf CType(sender, StackPanel).DataContext Is clsFilter Then flt = CType(CType(sender, StackPanel).DataContext, clsFilter)
        If flt Is Nothing Then Exit Sub
        CType(sender, StackPanel).ContextMenu.Tag = flt
    End Sub

    Private Sub tviDomains_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)

        If TypeOf sp.Tag Is clsDirectoryObject Then
            OpenObject(CType(sp.Tag, clsDirectoryObject))
        End If
    End Sub

    Private Sub tviDomains_TreeViewItem_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
        Dim obj As clsDirectoryObject = Nothing
        If TypeOf CType(sender, StackPanel).DataContext Is clsDirectoryObject Then obj = CType(CType(sender, StackPanel).DataContext, clsDirectoryObject)
        If obj Is Nothing And currentcontainer Is Nothing Then Exit Sub
        If obj Is Nothing Then obj = currentcontainer
        CType(sender, StackPanel).ContextMenu.Tag = {obj}
    End Sub

    Private Sub dgObjects_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs) Handles dgObjects.ContextMenuOpening
        Dim objects() As clsDirectoryObject = Nothing

        If dgObjects.SelectedItems.Count = 0 Then
            If currentcontainer IsNot Nothing Then objects = {currentcontainer}
        ElseIf dgObjects.SelectedItems.Count = 1 Then
            If TypeOf dgObjects.SelectedItem Is clsDirectoryObject Then objects = {dgObjects.SelectedItem}
        ElseIf dgObjects.SelectedItems.Count > 1 Then
            Dim isobjectsflag As Boolean = True ' all selected objects is clsDirectoryObject
            Dim tmpobjs As New List(Of clsDirectoryObject)
            For Each obj In dgObjects.SelectedItems
                If TypeOf obj IsNot clsDirectoryObject Then isobjectsflag = False : Exit For
                tmpobjs.Add(obj)
            Next
            If isobjectsflag Then objects = tmpobjs.ToArray
        End If
        If objects Is Nothing Then e.Handled = True : Exit Sub

        ctxmnuObjectsExternalSoftware.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer))
        If objects.Count = 1 Then
            ctxmnuObjectsExternalSoftware.Items.Clear()
            For Each es As clsExternalSoftware In preferences.ExternalSoftware
                Dim esmnu As New MenuItem
                esmnu.Header = es.Label
                esmnu.Icon = New Image With {.Source = es.Image}
                esmnu.Tag = es
                AddHandler esmnu.Click, AddressOf ctxmnuObjectsExternalSoftwareItem_Click
                ctxmnuObjectsExternalSoftware.Items.Add(esmnu)
            Next
        End If

        ctxmnuObjectsSelectAll.Visibility = BooleanToVisibility(dgObjects.Items.Count > 1)

        ctxmnuObjectsCreateObject.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.DomainDNS))
        ctxmnuObjectsCopy.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsCut.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsMove.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsRename.Visibility = BooleanToVisibility(objects.Count = 1)
        ctxmnuObjectsRemove.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsAddToFavorites.Visibility = BooleanToVisibility(objects.Count = 1)

        ctxmnuObjectsResetPassword.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)
        ctxmnuObjectsDisableEnable.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer))
        ctxmnuObjectsExpirationDate.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)
        ctxmnuObjectsCopyAttributes.Visibility = BooleanToVisibility(objects.Count > 0)

        ctxmnuObjectsProperties.Visibility = BooleanToVisibility(objects.Count = 1)

        dgObjects.ContextMenu.Tag = objects
    End Sub

    Private Sub ctxmnuObjectsExternalSoftwareItem_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(CType(sender, MenuItem).Parent, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(CType(sender, MenuItem).Parent, MenuItem).Parent, ContextMenu).Tag
        If objects.Count <> 1 Then Exit Sub

        Dim esmnu As MenuItem = CType(sender, MenuItem)
        If esmnu Is Nothing Then Exit Sub
        Dim es As clsExternalSoftware = CType(esmnu.Tag, clsExternalSoftware)
        If es Is Nothing Then Exit Sub

        Dim args As String = es.Arguments
        If args Is Nothing Then args = ""

        Dim patterns As MatchCollection = Regex.Matches(args, "{{(.*?)}}")

        For Each pattern As Match In patterns
            Dim val As String = If(objects(0).LdapProperty(pattern.Value.Replace("{{", "").Replace("}}", "")), pattern.Value).ToString
            args = Replace(args, pattern.Value, val)
        Next

        args = Replace(args, "{{myusername}}", objects(0).Domain.Username)
        args = Replace(args, "{{mypassword}}", objects(0).Domain.Password)
        args = Replace(args, "{{mydomain}}", objects(0).Domain.Name)

        Dim psi As New ProcessStartInfo(es.Path, args)

        If es.CurrentCredentials = True Then
            psi.WorkingDirectory = (New FileInfo(es.Path)).DirectoryName
            psi.UseShellExecute = False
            Process.Start(psi)
        Else
            psi.Domain = objects(0).Domain.Name
            psi.UserName = objects(0).Domain.Username
            psi.Password = StringToSecureString(objects(0).Domain.Password)
            psi.WorkingDirectory = (New FileInfo(es.Path)).DirectoryName
            psi.UseShellExecute = False
            Process.Start(psi)
        End If
    End Sub

    Private Sub ctxmnuObjectsSelectAll_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuObjectsSelectAll.Click
        dgObjects.SelectAll()
    End Sub

    Private Sub ctxmnuSharedProperties_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 Then ShowDirectoryObjectProperties(objects(0), Window.GetWindow(Me))
    End Sub

    Private Sub ctxmnuSharedAddToFavorites_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 Then preferences.Favorites.Add(objects(0))
    End Sub

    Private Sub ctxmnutviFavoritesRemoveFromFavorites_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 AndAlso preferences.Favorites.Contains(objects(0)) Then preferences.Favorites.Remove(objects(0))
    End Sub

    Private Sub ctxmnutviFiltersRemoveFromFilters_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsFilter Then Exit Sub
        Dim current As clsFilter = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If preferences.Filters.Contains(current) Then preferences.Filters.Remove(current)
    End Sub


#End Region

#Region "Events"




    Private Sub dgObjects_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles dgObjects.PreviewKeyDown
        Select Case e.Key
            Case Key.Enter
                If dgObjects.SelectedItem Is Nothing Then Exit Sub
                Dim current As clsDirectoryObject = CType(dgObjects.SelectedItem, clsDirectoryObject)
                OpenObject(current)
                e.Handled = True
            Case Key.Back
                OpenObjectParent()
                e.Handled = True
            Case Key.Home
                If dgObjects.Items.Count > 0 Then dgObjects.SelectedIndex = 0 : dgObjects.CurrentItem = dgObjects.SelectedItem
                e.Handled = True
            Case Key.End
                If dgObjects.Items.Count > 0 Then dgObjects.SelectedIndex = dgObjects.Items.Count - 1 : dgObjects.CurrentItem = dgObjects.SelectedItem
                e.Handled = True
        End Select
    End Sub

    Private Sub dgObjects_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles dgObjects.MouseDoubleClick
        If dgObjects.SelectedItem Is Nothing Or Not (e.ChangedButton = MouseButton.Left) Then Exit Sub
        Dim current As clsDirectoryObject = CType(dgObjects.SelectedItem, clsDirectoryObject)
        OpenObject(current)
    End Sub

    Private Sub dgObjects_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles dgObjects.MouseDown
        If e.ChangedButton = MouseButton.Left Or e.ChangedButton = MouseButton.Right Then
            Dim r As HitTestResult = VisualTreeHelper.HitTest(sender, e.GetPosition(sender))
            If TypeOf r.VisualHit Is ScrollViewer Then
                dgObjects.UnselectAll()
            End If
        End If
        If e.ChangedButton = MouseButton.XButton1 Then SearchPrevious()
        If e.ChangedButton = MouseButton.XButton2 Then SearchNext()
    End Sub

    Private Sub btnBack_Click(sender As Object, e As RoutedEventArgs) Handles btnBack.Click
        SearchPrevious()
    End Sub

    Private Sub btnForward_Click(sender As Object, e As RoutedEventArgs) Handles btnForward.Click
        SearchNext()
    End Sub

    Private Sub btnUp_Click(sender As Object, e As RoutedEventArgs) Handles btnUp.Click
        OpenObjectParent()
    End Sub

    Private Sub btnPath_Click(sender As Object, e As RoutedEventArgs)
        If sender.Tag Is Nothing Then Exit Sub
        Dim current As clsDirectoryObject = CType(sender.Tag, clsDirectoryObject)
        OpenObject(current)
    End Sub

    Private Sub cmboSearchPattern_KeyDown(sender As Object, e As KeyEventArgs) Handles cmboSearchPattern.KeyDown
        If e.Key = Key.Enter Then
            If mnuSearchModeDefault.IsChecked = True Then
                StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text, Nothing, searchobjectclasses, False))
            ElseIf mnuSearchModeAdvanced.IsChecked = True Then
                StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text))
            End If
        Else
            If cmboSearchPattern.IsDropDownOpen Then cmboSearchPattern.IsDropDownOpen = False
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If mnuSearchModeDefault.IsChecked = True Then
            StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text, Nothing, searchobjectclasses, False))
        ElseIf mnuSearchModeAdvanced.IsChecked = True Then
            StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text))
        End If
        cmboSearchPattern.Focus()
    End Sub

    Private Async Sub pbSearch_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles pbSearch.MouseDoubleClick
        Await searcher.BasicSearchStopAsync()
    End Sub

    Private Sub btnWindowClone_Click(sender As Object, e As RoutedEventArgs) Handles btnWindowClone.Click
        Dim w As New wndMain
        w.Show()
    End Sub

#End Region

#Region "Subs"

    Public Sub DomainTreeUpdate()
        tviDomains.ItemsSource = domains.Where(Function(d As clsDomain) d.Validated).Select(Function(d) If(d IsNot Nothing, New clsDirectoryObject(d.DefaultNamingContext, d), Nothing))
    End Sub

    Public Sub OpenObject(current As clsDirectoryObject)
        If current.SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or
           current.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or
           current.SchemaClass = clsDirectoryObject.enmSchemaClass.UnknownContainer Or
           current.SchemaClass = clsDirectoryObject.enmSchemaClass.DomainDNS Then
            StartSearch(current, Nothing)

        Else
            ShowDirectoryObjectProperties(current, Window.GetWindow(Me))
        End If
    End Sub

    Public Sub OpenObjectParent()
        If currentcontainer Is Nothing OrElse (currentcontainer.Entry.Parent Is Nothing OrElse currentcontainer.Entry.Path = currentcontainer.Domain.DefaultNamingContext.Path) Then Exit Sub
        Dim parent As New clsDirectoryObject(currentcontainer.Entry.Parent, currentcontainer.Domain)
        StartSearch(parent, Nothing)
    End Sub

    Private Sub ShowPath(Optional container As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        For Each child In tlbrPath.Items
            If TypeOf child Is Button Then RemoveHandler CType(child, Button).Click, AddressOf btnPath_Click
        Next

        tlbrPath.Items.Clear()

        If container IsNot Nothing Then

            Dim buttons As New List(Of Button)

            Do
                Dim btn As New Button
                Dim st As Style = Application.Current.TryFindResource("ToolbarButton")
                btn.Style = st
                btn.Content = If(container.objectClass.Contains("domaindns"), container.Domain.Name, container.name)
                btn.Margin = New Thickness(2, 0, 2, 0)
                btn.Padding = New Thickness(5, 0, 5, 0)
                btn.Tag = container
                buttons.Add(btn)

                If container.Entry.Parent Is Nothing OrElse container.Entry.Path = container.Domain.DefaultNamingContext.Path Then
                    Exit Do
                Else
                    container = New clsDirectoryObject(container.Entry.Parent, container.Domain)
                End If
            Loop

            buttons.Reverse()

            For Each child As Button In buttons
                AddHandler child.Click, AddressOf btnPath_Click
                tlbrPath.Items.Add(child)
            Next

        ElseIf filter IsNot Nothing Then

            Dim tblck As New TextBlock
            tblck.Background = Brushes.Transparent
            tblck.Text = My.Resources.wndMain_lbl_SearchResults & " " & If(Not String.IsNullOrEmpty(filter.Name), filter.Name, If(Not String.IsNullOrEmpty(filter.Pattern), filter.Pattern, " Advanced filter"))
            tblck.Margin = New Thickness(2, 0, 2, 0)
            tblck.Padding = New Thickness(5, 0, 5, 0)
            tlbrPath.Items.Add(tblck)

        End If
    End Sub

    Public Sub RebuildColumns()
        If currentobjects.Count > 10 Then currentobjects.Clear()

        dgObjects.Columns.Clear()
        For Each columninfo As clsDataGridColumnInfo In preferences.Columns
            dgObjects.Columns.Add(CreateColumn(columninfo))
        Next
    End Sub

    Public Sub UpdateColumns()
        For Each dgcolumn As DataGridColumn In dgObjects.Columns
            For Each pcolumn As clsDataGridColumnInfo In preferences.Columns
                If dgcolumn.Header.ToString = pcolumn.Header Then
                    pcolumn.DisplayIndex = dgcolumn.DisplayIndex
                    pcolumn.Width = dgcolumn.ActualWidth
                End If
            Next
        Next
    End Sub

    Private Sub HotKey_F5()
        If searchhistoryindex < 0 OrElse searchhistoryindex + 1 > searchhistory.Count Then Exit Sub
        Search(searchhistory(searchhistoryindex).Root, searchhistory(searchhistoryindex).Filter)
    End Sub

    Private Sub SearchPrevious()
        If searchhistoryindex <= 0 OrElse searchhistoryindex + 1 > searchhistory.Count Then Exit Sub
        Search(searchhistory(searchhistoryindex - 1).Root, searchhistory(searchhistoryindex - 1).Filter)
        searchhistoryindex -= 1
    End Sub

    Private Sub SearchNext()
        If searchhistoryindex < 0 OrElse searchhistoryindex + 2 > searchhistory.Count Then Exit Sub
        Search(searchhistory(searchhistoryindex + 1).Root, searchhistory(searchhistoryindex + 1).Filter)
        searchhistoryindex += 1
    End Sub

    Public Sub StartSearch(Optional root As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        While searchhistory.Count > searchhistoryindex + 1
            searchhistory.RemoveAt(searchhistory.Count - 1)
        End While

        searchhistory.Add(New clsSearchHistory(root, filter))
        If Not ADToolsApplication.ocGlobalSearchHistory.Contains(cmboSearchPattern.Text) Then ADToolsApplication.ocGlobalSearchHistory.Insert(0, cmboSearchPattern.Text)
        searchhistoryindex = searchhistory.Count - 1

        Search(root, filter)
    End Sub

    Public Async Sub Search(Optional root As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        Try
            CollectionViewSource.GetDefaultView(dgObjects.ItemsSource).GroupDescriptions.Clear()
        Catch
        End Try

        Dim cmboTextBoxChild As TextBox = cmboSearchPattern.Template.FindName("PART_EditableTextBox", cmboSearchPattern)
        If cmboTextBoxChild IsNot Nothing Then cmboTextBoxChild.SelectAll()

        cap.Visibility = Visibility.Visible
        pbSearch.Visibility = Visibility.Visible

        Dim domainlist As New ObservableCollection(Of clsDomain)(domains.Where(Function(x As clsDomain) x.IsSearchable = True).ToList)

        currentcontainer = root
        currentfilter = filter
        ShowPath(currentcontainer, filter)

        Await searcher.BasicSearchAsync(currentobjects, root, filter, domainlist)

        If preferences.SearchResultGrouping Then
            Try
                CollectionViewSource.GetDefaultView(dgObjects.ItemsSource).GroupDescriptions.Add(New PropertyGroupDescription("Domain.Name"))
            Catch
            End Try
        End If

        cap.Visibility = Visibility.Hidden
        pbSearch.Visibility = Visibility.Hidden
    End Sub

    Private Sub Searcher_BasicSearchAsyncDataRecieved() Handles searcher.BasicSearchAsyncDataRecieved
        cap.Visibility = Visibility.Hidden
    End Sub

    Public Function CreateColumn(columninfo As clsDataGridColumnInfo) As DataGridTemplateColumn
        Dim column As New DataGridTemplateColumn()
        column.Header = columninfo.Header
        column.SetValue(DataGridColumn.CanUserSortProperty, True)
        If columninfo.DisplayIndex > 0 Then column.DisplayIndex = columninfo.DisplayIndex
        If columninfo.Width > 0 Then column.Width = columninfo.Width
        Dim panel As New FrameworkElementFactory(GetType(VirtualizingStackPanel))
        panel.SetValue(VirtualizingStackPanel.VerticalAlignmentProperty, VerticalAlignment.Center)
        panel.SetValue(VirtualizingStackPanel.MarginProperty, New Thickness(5, 0, 5, 0))

        Dim first As Boolean = True
        For Each attr As clsAttribute In columninfo.Attributes
            Dim bind As System.Windows.Data.Binding

            bind = New System.Windows.Data.Binding(attr.Name)
            bind.Mode = BindingMode.OneWay

            If attr.Name <> "Image" Then

                Dim text As New FrameworkElementFactory(GetType(CustomTextBlock))
                If first Then
                    text.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold)
                    first = False
                    column.SetValue(DataGridColumn.SortMemberPathProperty, attr.Name)
                End If
                bind.Converter = New ConverterToInlinesWithHyperlink
                text.SetBinding(CustomTextBlock.InlineCollectionProperty, bind)
                text.SetValue(TextBlock.ToolTipProperty, attr.Label)
                'text.SetValue(TextBlock.TextWrappingProperty, TextWrapping.WrapWithOverflow)
                panel.AppendChild(text)

            Else

                Dim ttbind As New System.Windows.Data.Binding("Status")
                ttbind.Mode = BindingMode.OneWay
                Dim img As New FrameworkElementFactory(GetType(Image))
                column.SetValue(clsSorter.PropertyNameProperty, "Image")
                img.SetBinding(Image.SourceProperty, bind)
                img.SetValue(Image.WidthProperty, 32.0)
                img.SetValue(Image.HeightProperty, 32.0)
                img.SetBinding(Image.ToolTipProperty, ttbind)
                panel.AppendChild(img)

            End If
            'Status
        Next

        Dim template As New DataTemplate()
        template.VisualTree = panel

        column.CellTemplate = template

        Return column
    End Function



#End Region

End Class
