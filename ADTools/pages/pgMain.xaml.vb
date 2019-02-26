Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Reflection
Imports System.Windows.Controls.Primitives
Imports IPrint
Imports IPrompt.VisualBasic

Class pgMain

    Public WithEvents clipboardTimer As Threading.DispatcherTimer
    Public Property clipboardlastdata As String

    Private Sub pgMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dpToolbar.DataContext = preferences
        mnuSearchDomains.ItemsSource = domains

        tviFavorites.ItemsSource = preferences.Favorites
        tviFilters.ItemsSource = preferences.Filters
        gcdNavigation.DataContext = preferences

        clipboardTimer = New Threading.DispatcherTimer
        clipboardTimer.Interval = New TimeSpan(0, 0, 1)
        clipboardTimer.Start()

        AddHandler domains.CollectionChanged, AddressOf DomainsChangedHandler
        For Each domain As clsDomain In domains
            AddHandler domain.PropertyChanged, AddressOf DomainValidatedHandler
        Next
        DomainRefreshHandler()
    End Sub



    Public Function CurrentObjectsPage() As pgObjects
        If frmObjects IsNot Nothing AndAlso frmObjects.Content IsNot Nothing AndAlso TypeOf frmObjects.Content Is pgObjects Then
            Return CType(frmObjects.Content, pgObjects)
        Else
            Return Nothing
        End If
    End Function

#Region "Main Toolbar"

    Private Sub mnuFilePrint_Click(sender As Object, e As RoutedEventArgs) Handles mnuFilePrint.Click
        Dim currentobjects As clsThreadSafeObservableCollection(Of clsDirectoryObject) = Nothing
        Dim pg = CurrentObjectsPage()
        If pg Is Nothing Then Exit Sub
        currentobjects = pg.currentobjects
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
                        If pi.Name = attr Then
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
        ShowPage(New pgDomains, True, Window.GetWindow(Me), False)
    End Sub

    Private Sub mnuServicePreferences_Click(sender As Object, e As RoutedEventArgs) Handles mnuServicePreferences.Click
        ShowPage(New pgPreferences, True, Window.GetWindow(Me), False)
    End Sub

    'Private Sub mnuServiceLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceLog.Click
    '    Dim w As New wndLog
    '    w.Owner = Window.GetWindow(Me)
    '    w.Show()
    'End Sub

    Private Sub mnuServiceErrorLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceErrorLog.Click
        Dim w As New wndErrorLog
        w.Owner = Window.GetWindow(Me)
        w.Show()
    End Sub

    Private Sub mnuSearchSaveCurrentFilter_Click(sender As Object, e As RoutedEventArgs) Handles mnuSearchSaveCurrentFilter.Click
        Dim pg = CurrentObjectsPage()
        If pg Is Nothing Then Exit Sub
        pg.SaveCurrentFilter()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As RoutedEventArgs) Handles mnuHelpAbout.Click
        Dim p As New pgAbout
        ShowPage(p, True, Window.GetWindow(Me), False)
    End Sub

#End Region

#Region "Context Menu"


    Private Sub tviFilters_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)
        Dim pg = CurrentObjectsPage()
        If pg Is Nothing Then Exit Sub

        If TypeOf sp.Tag Is clsFilter Then
            pg.StartSearch(Nothing, CType(sp.Tag, clsFilter))
        End If
    End Sub

    Private Sub tviDomainstviFavorites_TreeViewItem_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
        Dim obj As clsDirectoryObject = Nothing
        If TypeOf CType(sender, StackPanel).DataContext Is clsDirectoryObject Then obj = CType(CType(sender, StackPanel).DataContext, clsDirectoryObject)
        If obj Is Nothing Then Exit Sub

        If obj.IsDeleted Then e.Handled = True : Exit Sub

        CType(sender, StackPanel).ContextMenu.Tag = {obj}
    End Sub

    Private Sub tviDomainstviFavorites_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)

        If TypeOf sp.Tag Is clsDirectoryObject Then
            Dim pg = CurrentObjectsPage()
            If pg Is Nothing Then Exit Sub

            pg.OpenObject(CType(sp.Tag, clsDirectoryObject))
        End If
    End Sub

    Private Sub tviDomains_TreeViewItem_DragEnterDragOver(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            If e.KeyStates.HasFlag(DragDropKeyStates.ControlKey) Then
                e.Effects = DragDropEffects.Copy
            Else
                e.Effects = DragDropEffects.Move
            End If

            For Each obj As clsDirectoryObject In e.Data.GetData(GetType(clsDirectoryObject()))
                If Not (obj.SchemaClass = enmDirectoryObjectSchemaClass.User Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Contact Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Computer Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Group Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit) Then e.Effects = DragDropEffects.None : Exit For
            Next
        Else
            e.Effects = DragDropEffects.None
        End If

        e.Handled = True
    End Sub

    Private Sub tviDomains_TreeViewItem_Drop(sender As Object, e As DragEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)
        If TypeOf sp.Tag IsNot clsDirectoryObject Then Exit Sub
        Dim destination As clsDirectoryObject = sp.Tag

        If Not (destination.SchemaClass = enmDirectoryObjectSchemaClass.Container Or
            destination.SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit Or
            destination.SchemaClass = enmDirectoryObjectSchemaClass.UnknownContainer) Then Exit Sub

        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            Dim dropped = e.Data.GetData(GetType(clsDirectoryObject()))
            For Each obj As clsDirectoryObject In dropped
                If Not (obj.SchemaClass = enmDirectoryObjectSchemaClass.User Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Contact Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Computer Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Group Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit) Then Exit Sub
            Next

            If e.KeyStates.HasFlag(DragDropKeyStates.ControlKey) Then

                Dim pg = CurrentObjectsPage()
                If pg Is Nothing Then Exit Sub
                pg.ObjectsCopy(destination, dropped)

            Else

                Dim pg = CurrentObjectsPage()
                If pg Is Nothing Then Exit Sub
                pg.ObjectsMove(destination, dropped)

            End If

        End If
    End Sub

    Private Sub ctxmnutviDomainsDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnutviDomainsDomainOptions.Click
        ShowPage(New pgDomains, True, Window.GetWindow(Me), True)
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

#Region "Context Menu Shared"

    Private Sub tviFilters_TreeViewItem_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
        Dim flt As clsFilter = Nothing
        If TypeOf CType(sender, StackPanel).DataContext Is clsFilter Then flt = CType(CType(sender, StackPanel).DataContext, clsFilter)
        If flt Is Nothing Then Exit Sub
        CType(sender, StackPanel).ContextMenu.Tag = flt
    End Sub

    Private Sub ctxmnuSharedCreateObject_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag

        If objects.Count = 1 AndAlso (
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.Container Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.DomainDNS Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.UnknownContainer) Then

            ShowPage(New pgCreateObject(objects(0), objects(0).Domain), False, Window.GetWindow(Me), False)
        Else
            ShowPage(New pgCreateObject(Nothing, Nothing), False, Window.GetWindow(Me), False)
        End If

    End Sub

    Private Sub ctxmnuSharedCopy_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag

        Try
            Clipboard.SetDataObject(Join(objects.Select(Function(o) o.name).ToArray, vbCrLf), True)
        Catch ex As Exception
            IMsgBox(ex.Message, vbExclamation)
        End Try

        ClipboardBuffer = objects.Where(
            Function(obj)
                Return obj.SchemaClass = enmDirectoryObjectSchemaClass.User Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Contact Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Computer Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Group Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit
            End Function).ToArray

        ClipboardAction = enmClipboardAction.Copy
    End Sub

    Private Sub ctxmnuSharedCut_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag

        Try
            Clipboard.SetDataObject(Join(objects.Select(Function(o) o.name).ToArray, vbCrLf), True)
        Catch ex As Exception
            IMsgBox(ex.Message, vbExclamation)
        End Try

        ClipboardBuffer = objects.Where(
            Function(obj)
                Return obj.SchemaClass = enmDirectoryObjectSchemaClass.User Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Contact Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Computer Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.Group Or
                       obj.SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit
            End Function).ToArray

        ClipboardAction = enmClipboardAction.Cut
    End Sub

    Private Sub ctxmnuSharedPaste_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim destinations() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (destinations.Count = 1 AndAlso (
            destinations(0).SchemaClass = enmDirectoryObjectSchemaClass.Container Or
            destinations(0).SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit Or
            destinations(0).SchemaClass = enmDirectoryObjectSchemaClass.UnknownContainer)) Then Exit Sub

        If Not (ClipboardBuffer IsNot Nothing AndAlso ClipboardBuffer.Count > 0) Then Exit Sub

        Dim destination As clsDirectoryObject = destinations(0)
        Dim sourceobjects() As clsDirectoryObject = ClipboardBuffer

        If ClipboardAction = enmClipboardAction.Copy Then ' copy

            Dim pg = CurrentObjectsPage()
            If pg Is Nothing Then Exit Sub
            pg.ObjectsCopy(destination, sourceobjects)

        ElseIf ClipboardAction = enmClipboardAction.Cut Then ' cut

            Dim pg = CurrentObjectsPage()
            If pg Is Nothing Then Exit Sub
            pg.ObjectsMove(destination, sourceobjects)

        End If

    End Sub

    Private Sub ctxmnuSharedRename_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.User Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.Contact Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.Computer Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.Group Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit)) Then Exit Sub

        Try
            Dim organizationalunitaffected As Boolean = False

            Dim obj As clsDirectoryObject = objects(0)
            Dim name As String = IInputBox(My.Resources.str_EnterObjectName, My.Resources.str_RenameObject, objects(0).name, vbQuestion, Window.GetWindow(Me))
            If Len(name) > 0 Then
                If obj.SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit Then
                    obj.Rename("OU=" & name)
                    organizationalunitaffected = True
                Else
                    obj.Rename("CN=" & name)
                End If

                'RefreshSearchResults()
                'TODO
                'If organizationalunitaffected Then RefreshDomainTree()
            End If
        Catch ex As Exception
            ThrowException(ex, "ctxmnuSharedRename_Click")
        End Try
    End Sub

    Private Sub ctxmnuSharedRemove_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.User Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.Contact Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.Computer Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.Group Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit)) Then Exit Sub

        Try
            Dim currentcontaineraffected As Boolean = False

            If IMsgBox(My.Resources.str_AreYouSure, vbYesNo + vbQuestion, My.Resources.str_RemoveObject, Window.GetWindow(Me)) <> MsgBoxResult.Yes Then Exit Sub
            Dim pg = CurrentObjectsPage()
            If pg Is Nothing Then Exit Sub

            If objects(0).SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit Then
                If IMsgBox(My.Resources.str_ThisIsOrganizationaUnit & vbCrLf & vbCrLf & My.Resources.str_AreYouSure, vbYesNo + vbExclamation, My.Resources.str_RemoveObject, Window.GetWindow(Me)) <> MsgBoxResult.Yes Then Exit Sub
                If pg.currentcontainer IsNot Nothing AndAlso objects(0).distinguishedName = pg.currentcontainer.distinguishedName Then currentcontaineraffected = True
            End If

            objects(0).DeleteTree()

            If currentcontaineraffected Then
                pg.OpenObjectParent()
            End If

        Catch ex As Exception
            ThrowException(ex, "ctxmnuSharedRemove_Click")
        End Try
    End Sub

    Private Sub ctxmnuSharedAddToFavorites_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 Then preferences.Favorites.Add(objects(0))
    End Sub

    Private Sub ctxmnuSharedOpenObjectLocation_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        Dim pg = CurrentObjectsPage()
        If pg Is Nothing Then Exit Sub

        If objects.Count = 1 Then pg.OpenObject(objects(0).Parent)
    End Sub

    Private Sub ctxmnuSharedOpenObjectTree_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 Then
            Dim w As NavigationWindow = CType(Window.GetWindow(Me), NavigationWindow)
            If TypeOf w.Content Is pgMain Then CType(w.Content, pgMain).ShowInTree(objects(0).Parent)
        End If
    End Sub

    Private Sub ctxmnuSharedResetPassword_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.User)) Then Exit Sub

        Try
            If IMsgBox(My.Resources.str_AreYouSure, vbYesNo + vbQuestion, My.Resources.str_PasswordReset, Window.GetWindow(Me)) = MsgBoxResult.Yes Then
                objects(0).ResetPassword()
                objects(0).passwordNeverExpires = False
                IMsgBox(My.Resources.str_PasswordChanged, vbOKOnly + vbInformation, My.Resources.str_PasswordReset)
            End If

        Catch ex As Exception
            ThrowException(ex, "ctxmnuSharedResetPassword_Click")
        End Try
    End Sub

    Private Sub ctxmnuSharedDisableEnable_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.User Or
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.Computer)) Then Exit Sub

        Try
            If IMsgBox(My.Resources.str_AreYouSure, vbYesNo + vbQuestion, If(objects(0).disabled, My.Resources.str_Enable, My.Resources.str_Disable), Window.GetWindow(Me)) = MsgBoxResult.Yes Then
                objects(0).disabled = Not objects(0).disabled
            End If

        Catch ex As Exception
            ThrowException(ex, "ctxmnuSharedDisableEnable_Click")
        End Try
    End Sub

    Private Sub ctxmnuSharedExpirationDate_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = enmDirectoryObjectSchemaClass.User)) Then Exit Sub

        Try
            objects(0).accountExpiresDate = Today.AddDays(1)
            Dim w = ShowDirectoryObjectProperties(objects(0), Window.GetWindow(Me))
            If w IsNot Nothing AndAlso w.Content IsNot Nothing AndAlso TypeOf w.Content Is pgObject Then CType(w.Content, pgObject).FirstPage = New pgUserObject(objects(0))
        Catch ex As Exception
            ThrowException(ex, "ctxmnuSharedExpirationDate_Click")
        End Try
    End Sub

    Private Sub ctxmnuSharedProperties_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 Then ShowDirectoryObjectProperties(objects(0), Window.GetWindow(Me))
    End Sub

#End Region

#Region "Events"

    <DebuggerStepThrough>
    Private Sub clipboardTimer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles clipboardTimer.Tick
        If preferences Is Nothing OrElse preferences.ClipboardSource = False Then Exit Sub
        If Not Clipboard.ContainsText(TextDataFormat.Text) Then Exit Sub
        Dim newclipboarddata As String

        Try
            newclipboarddata = Clipboard.GetText(TextDataFormat.Text)
            If String.IsNullOrEmpty(newclipboarddata) Then Exit Sub
            If preferences.ClipboardSourceLimit AndAlso CountWords(newclipboarddata) > 3 Then Exit Sub ' only three words

            If clipboardlastdata <> newclipboarddata Then
                clipboardlastdata = newclipboarddata

                Dim pg = CurrentObjectsPage()
                If pg Is Nothing Then Exit Sub

                pg.StartSearch(Nothing, New clsFilter(newclipboarddata.Replace(vbNewLine, " / "), preferences.AttributesForSearch, preferences.SearchObjectClasses))
            End If
        Catch
        End Try
    End Sub

    Private Sub btnWindowClone_Click(sender As Object, e As RoutedEventArgs) Handles btnWindowClone.Click
        ADToolsApplication.CreateMainWindow()
    End Sub

    Private Sub btnDummy_Click(sender As Object, e As RoutedEventArgs) Handles btnDummy.Click

    End Sub

#End Region

#Region "Subs"

    Public Sub DomainsChangedHandler(sender As Object, e As NotifyCollectionChangedEventArgs)
        If e.OldItems IsNot Nothing Then
            For Each domain As clsDomain In e.OldItems
                RemoveHandler domain.PropertyChanged, AddressOf DomainValidatedHandler
            Next
        End If
        If e.NewItems IsNot Nothing Then
            For Each domain As clsDomain In e.NewItems
                AddHandler domain.PropertyChanged, AddressOf DomainValidatedHandler
            Next
        End If
        DomainRefreshHandler()
    End Sub

    Public Sub DomainValidatedHandler(sender As Object, e As PropertyChangedEventArgs)
        If e.PropertyName = "Validated" Then DomainRefreshHandler()
    End Sub

    Public Sub DomainRefreshHandler()
        tviDomains.ItemsSource = domains.Where(Function(d As clsDomain) d.Validated).Select(Function(d As clsDomain) New clsDirectoryObject(d.DefaultNamingContext, d))
    End Sub

    Public Async Sub ShowInTree(container As clsDirectoryObject)
        If container Is Nothing Then Exit Sub

        Dim containerDN As String = container.distinguishedName

        Dim treepath As New List(Of String)

        Do ' find all parents to the root
            treepath.Add(containerDN)

            If containerDN.StartsWith("DC=", StringComparison.OrdinalIgnoreCase) Then
                Exit Do
            Else
                containerDN = GetParentDNFromDN(containerDN)
            End If
        Loop

        treepath.Reverse()

        Dim currentparentnode = tviDomains
        For I = 0 To treepath.Count - 1
            For j = 0 To currentparentnode.Items.Count - 1
                If CType(currentparentnode.Items(j), clsDirectoryObject).distinguishedName <> treepath(I) Then Continue For

                If currentparentnode.ItemContainerGenerator.Status <> GeneratorStatus.ContainersGenerated Then
                    Await Task.Run(
                        Sub()
                            Dim timeout As Integer
                            Do While currentparentnode.ItemContainerGenerator.Status <> GeneratorStatus.ContainersGenerated And timeout < 100 ' 10 sec
                                Threading.Thread.Sleep(100)
                                timeout += 1
                            Loop
                        End Sub)
                End If

                Dim st = currentparentnode.ItemContainerGenerator.Status
                Dim childnode As TreeViewItem = currentparentnode.ItemContainerGenerator.ContainerFromIndex(j)
                If childnode Is Nothing Then Exit Sub

                childnode.SetValue(TreeViewItem.IsExpandedProperty, True)
                childnode.ApplyTemplate()
                childnode.UpdateLayout()

                currentparentnode = childnode
                Exit For
            Next
        Next
    End Sub

#End Region

End Class
