Imports System.Collections.ObjectModel
Imports System.Windows.Controls.Primitives
Imports IPrompt.VisualBasic

Class wndMain
    Public WithEvents searcher As New clsSearcher

    Public Shared hkF5 As New RoutedCommand

    Public Property currentcontainer As clsDirectoryObject
    Public Property currentobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)
    Public Property currentfilter As clsFilter

    Private searchhistoryindex As Integer
    Private searchhistory As New List(Of clsSearchHistory)

    Public Property searchobjectclasses As New clsSearchObjectClasses(True, True, True, True)


    Private Sub wndMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        hkF5.InputGestures.Add(New KeyGesture(Key.F5))
        Me.CommandBindings.Add(New CommandBinding(hkF5, AddressOf HotKey_F5))

        DomainTreeUpdate()
        RebuildColumns()

        cmboSearchPattern.ItemsSource = ADToolsApplication.ocGlobalSearchHistory
        cmboSearchPattern.Focus()

        mnuSearchDomains.ItemsSource = domains
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

    Private Sub mnuServiceDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceDomainOptions.Click
        ShowWindow(New wndDomains, True, Me, True)
        DomainTreeUpdate()
    End Sub

    Private Sub mnuSearchSaveCurrentFilter_Click(sender As Object, e As RoutedEventArgs) Handles mnuSearchSaveCurrentFilter.Click
        If currentfilter Is Nothing OrElse String.IsNullOrEmpty(currentfilter.Filter) Then IMsgBox(My.Resources.wndMain_msg_CannotSaveCurrentFilter, vbOKOnly + vbExclamation,, Me) : Exit Sub

        Dim name As String = IInputBox(My.Resources.wndMain_msg_EnterFilterName,,, vbQuestion, Me)

        If String.IsNullOrEmpty(name) Then Exit Sub

        currentfilter.Name = name
        preferences.Filters.Add(currentfilter)
    End Sub

#End Region


#Region "Events"

    Private Sub tviDomains_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)
        'System.Media.SystemSounds.Beep.Play()

        If TypeOf sp.Tag Is clsDirectoryObject Then
            OpenObject(CType(sp.Tag, clsDirectoryObject))
        End If
    End Sub

    Private Sub tviFilters_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)
        If TypeOf sp.Tag Is clsFilter Then
            StartSearch(Nothing, CType(sp.Tag, clsFilter))
        End If
    End Sub

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
        If dgObjects.SelectedItem Is Nothing Then Exit Sub
        Dim current As clsDirectoryObject = CType(dgObjects.SelectedItem, clsDirectoryObject)
        OpenObject(current)
    End Sub

    Private Sub dgObjects_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles dgObjects.MouseDown
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
                StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text, Nothing, Nothing, False))
            ElseIf mnuSearchModeAdvanced.IsChecked = True Then
                StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text))
            End If
        Else
            If cmboSearchPattern.IsDropDownOpen Then cmboSearchPattern.IsDropDownOpen = False
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If mnuSearchModeDefault.IsChecked = True Then
            StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text, Nothing, Nothing, False))
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
        If current.objectClass.Contains("computer") Or
           current.objectClass.Contains("group") Or
           current.objectClass.Contains("contact") Or
           current.objectClass.Contains("user") Then

            ShowDirectoryObjectProperties(current, Window.GetWindow(Me))

        ElseIf current.objectClass.Contains("organizationalunit") Or
               current.objectClass.Contains("container") Or
               current.objectClass.Contains("builtindomain") Or
               current.objectClass.Contains("domaindns") Then

            StartSearch(current, Nothing)
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






#End Region

End Class
