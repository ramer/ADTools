Imports System.Collections.ObjectModel

Class wndMain
    Public WithEvents searcher As New clsSearcher

    Public Shared hkF5 As New RoutedCommand

    Public Property objects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Private searchhistoryindex As Integer
    Private searchhistory As New List(Of String)

    Public Property searchobjectclasses As New clsSearchObjectClasses(True, True, True, True)

    Private Sub wndMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        hkF5.InputGestures.Add(New KeyGesture(Key.F5))
        Me.CommandBindings.Add(New CommandBinding(hkF5, AddressOf HotKey_F5))

        DomainTreeUpdate()
        RebuildColumns()

        cmboSearchPattern.ItemsSource = ADToolsApplication.ocGlobalSearchHistory
        cmboSearchPattern.Focus()

        cmboSearchObjectClasses.DataContext = searchobjectclasses
        cmboSearchDomains.ItemsSource = domains

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


#End Region


#Region "Events"

    Private Sub tvDomains_TreeViewItem_Selected(sender As Object, e As RoutedEventArgs)
        Dim trvitem As TreeViewItem = TryCast(sender, TreeViewItem)
        If trvitem Is e.OriginalSource AndAlso TypeOf CType(trvitem, TreeViewItem).DataContext Is clsDirectoryObject Then
            Dim parent As clsDirectoryObject = CType(CType(trvitem, TreeViewItem).DataContext, clsDirectoryObject)
            OpenObject(parent)
            trvitem.IsSelected = False
            'dgObjects.UpdateLayout()
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
            Search(cmboSearchPattern.Text)

            If Not String.IsNullOrEmpty(cmboSearchPattern.Text) Then
                While searchhistory.Count > searchhistoryindex + 1
                    searchhistory.RemoveAt(searchhistory.Count - 1)
                End While

                searchhistory.Add(cmboSearchPattern.Text)
                If Not ADToolsApplication.ocGlobalSearchHistory.Contains(cmboSearchPattern.Text) Then ADToolsApplication.ocGlobalSearchHistory.Insert(0, cmboSearchPattern.Text)
                searchhistoryindex = searchhistory.Count - 1
            End If
        Else
            If cmboSearchPattern.IsDropDownOpen Then cmboSearchPattern.IsDropDownOpen = False
        End If
    End Sub

#End Region

#Region "Subs"

    Public Sub DomainTreeUpdate()
        tvDomains.ItemsSource = domains.Where(Function(d As clsDomain) d.Validated).Select(Function(d) If(d IsNot Nothing, New clsDirectoryObject(d.DefaultNamingContext, d), Nothing))
    End Sub

    Public Sub OpenObject(current As clsDirectoryObject)
        If current.objectClass.Contains("computer") Or
           current.objectClass.Contains("group") Or
           current.objectClass.Contains("contact") Or
           current.objectClass.Contains("user") Then

            ShowDirectoryObjectProperties(current, Window.GetWindow(Me))

        ElseIf current.name = "Deleted Objects" Then
            'Search("""*""", True)
        ElseIf current.objectClass.Contains("organizationalunit") Or
               current.objectClass.Contains("container") Or
               current.objectClass.Contains("builtindomain") Or
               current.objectClass.Contains("domaindns") Then

            ShowChildren(current)
            dgObjects.Tag = current
        End If

        ShowPath()
    End Sub

    Public Sub OpenObjectParent()
        Dim current As clsDirectoryObject = CType(dgObjects.Tag, clsDirectoryObject)
        If current Is Nothing OrElse (current.Entry.Parent Is Nothing OrElse current.Entry.Path = current.Domain.DefaultNamingContext.Path) Then Exit Sub
        Dim parent As New clsDirectoryObject(current.Entry.Parent, current.Domain)
        ShowChildren(parent)
        dgObjects.Tag = parent

        ShowPath()
    End Sub

    Private Sub ShowChildren(parent As clsDirectoryObject)
        objects.Clear()
        For Each child As clsDirectoryObject In parent.Children
            objects.Add(child)
        Next
    End Sub

    Private Sub ShowPath()
        Dim current As clsDirectoryObject = CType(dgObjects.Tag, clsDirectoryObject)
        If current Is Nothing Then Exit Sub

        Dim children As New List(Of Button)

        Do
            Dim btn As New Button
            btn.Background = Brushes.Transparent
            btn.Content = If(current.objectClass.Contains("domaindns"), current.Domain.Name, current.name)
            btn.Margin = New Thickness(2, 0, 2, 0)
            btn.Padding = New Thickness(5, 0, 5, 0)
            btn.Tag = current
            children.Add(btn)

            If current.Entry.Parent Is Nothing OrElse current.Entry.Path = current.Domain.DefaultNamingContext.Path Then
                Exit Do
            Else
                current = New clsDirectoryObject(current.Entry.Parent, current.Domain)
            End If
        Loop

        children.Reverse()

        For Each child As Button In tlbrPath.Items
            RemoveHandler child.Click, AddressOf btnPath_Click
        Next
        tlbrPath.Items.Clear()
        For Each child As Button In children
            AddHandler child.Click, AddressOf btnPath_Click
            tlbrPath.Items.Add(child)
        Next
    End Sub

    Public Sub RebuildColumns()
        If objects.Count > 10 Then objects.Clear()

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
        SearchRepeat()
    End Sub

    Private Sub SearchRepeat()
        Search(cmboSearchPattern.Text)
    End Sub

    Private Sub SearchPrevious()
        If searchhistoryindex <= 0 OrElse searchhistoryindex + 1 > searchhistory.Count Then Exit Sub
        Search(searchhistory(searchhistoryindex - 1))
        searchhistoryindex -= 1
    End Sub

    Private Sub SearchNext()
        If searchhistoryindex < 0 OrElse searchhistoryindex + 2 > searchhistory.Count Then Exit Sub
        Search(searchhistory(searchhistoryindex + 1))
        searchhistoryindex += 1
    End Sub

    Public Async Sub Search(pattern As String)
        If String.IsNullOrEmpty(pattern) Then Exit Sub

        Try
            CollectionViewSource.GetDefaultView(dgObjects.ItemsSource).GroupDescriptions.Clear()
        Catch
        End Try

        cmboSearchPattern.Text = pattern
        Dim cmboTextBoxChild As TextBox = cmboSearchPattern.Template.FindName("PART_EditableTextBox", cmboSearchPattern)
        If cmboTextBoxChild IsNot Nothing Then cmboTextBoxChild.SelectAll()

        cap.Visibility = Visibility.Visible
        pbSearch.Visibility = Visibility.Visible

        Dim domainlist As New ObservableCollection(Of clsDomain)(domains.Where(Function(x As clsDomain) x.IsSearchable = True).ToList)

        Await searcher.BasicSearchAsync(objects, pattern, domainlist, preferences.AttributesForSearch, searchobjectclasses, False)

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
