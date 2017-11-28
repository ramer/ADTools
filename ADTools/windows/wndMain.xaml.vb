Imports System.Collections.ObjectModel
Imports System.Reflection
Imports System.Windows.Controls.Primitives
Imports IPrompt.VisualBasic
Imports IPrint
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.ComponentModel

Class wndMain
    Public WithEvents searcher As New clsSearcher

    Public Shared hkF5 As New RoutedCommand
    Public Shared hkF1 As New RoutedCommand
    Public Shared hkEsc As New RoutedCommand

    Public Property currentcontainer As clsDirectoryObject
    Public Property currentfilter As clsFilter
    Public Property currentobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Public Property cvscurrentobjects As New CollectionViewSource
    Public Property cvcurrentobjects As ICollectionView

    Private searchhistoryindex As Integer
    Private searchhistory As New List(Of clsSearchHistory)

    Public Property searchobjectclasses As New clsSearchObjectClasses(True, True, True, True, False)

    Public WithEvents clipboardTimer As New Threading.DispatcherTimer()
    Private clipboardlastdata As String

    Private Sub wndMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        hkF1.InputGestures.Add(New KeyGesture(Key.F1))
        Me.CommandBindings.Add(New CommandBinding(hkF1, AddressOf ShowPopups))
        hkF5.InputGestures.Add(New KeyGesture(Key.F5))
        Me.CommandBindings.Add(New CommandBinding(hkF5, AddressOf RefreshDataGrid))
        hkEsc.InputGestures.Add(New KeyGesture(Key.Escape))
        Me.CommandBindings.Add(New CommandBinding(hkEsc, AddressOf StopSearch))

        RefreshDomainTree()
        UpdateDetailsViewGroupStyle(True)

        DataObject.AddPastingHandler(tbSearchPattern, AddressOf tbSearchPattern_OnPaste)
        tbSearchPattern.Focus()

        mnuSearchDomains.ItemsSource = domains
        tviFavorites.ItemsSource = preferences.Favorites
        tviFilters.ItemsSource = preferences.Filters

        cvscurrentobjects = New CollectionViewSource() With {.Source = currentobjects}
        cvcurrentobjects = cvscurrentobjects.View
        cvcurrentobjects.GroupDescriptions.Add(New PropertyGroupDescription("Domain.Name"))

        lvObjects.SetBinding(ItemsControl.ItemsSourceProperty, New Binding("cvcurrentobjects") With {.IsAsync = True})

        If preferences.FirstRun Then ShowPopups()

        clipboardTimer.Interval = New TimeSpan(0, 0, 1)
        clipboardTimer.Start()
    End Sub

    Private Sub wndMain_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles MyBase.Closing, MyBase.Closing
        Dim count As Integer = 0

        For Each wnd As Window In ADToolsApplication.Current.Windows
            If GetType(wndMain) Is wnd.GetType Then count += 1
        Next

        If preferences.CloseOnXButton AndAlso count <= 1 Then ApplicationDeactivate()
    End Sub

#Region "Popups"

    Private Sub ShowPopups()
        poptvObjects.IsOpen = True
        poptbSearchPattern.IsOpen = True
        popF1Hint.IsOpen = True
    End Sub

    Private Sub MovePopups() Handles Me.LocationChanged, Me.SizeChanged, tvObjects.SizeChanged, tbSearchPattern.SizeChanged
        poptvObjects.HorizontalOffset += 1 : poptvObjects.HorizontalOffset -= 1
        poptbSearchPattern.HorizontalOffset += 1 : poptbSearchPattern.HorizontalOffset -= 1
        popF1Hint.HorizontalOffset += 1 : popF1Hint.HorizontalOffset -= 1
    End Sub

    Private Sub ClosePopups()
        poptvObjects.IsOpen = False
        poptbSearchPattern.IsOpen = False
        popF1Hint.IsOpen = False
    End Sub

    Private Sub ClosePopup(sender As Object, e As MouseButtonEventArgs) Handles poptvObjects.MouseLeftButtonDown, poptbSearchPattern.MouseLeftButtonDown, popF1Hint.MouseLeftButtonDown
        CType(sender, Popup).IsOpen = False
    End Sub

#End Region

#Region "Main Menu"

    Private Sub mnuFileExit_Click(sender As Object, e As RoutedEventArgs) Handles mnuFileExit.Click
        Me.Close()
    End Sub

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

    Private Sub mnuEditCreateObject_Click(sender As Object, e As RoutedEventArgs) Handles mnuEditCreateObject.Click
        Dim w As New wndCreateObject
        w.destinationcontainer = Nothing
        w.destinationdomain = Nothing
        ShowWindow(w, False, Me, False)
    End Sub

    Private Sub mnuServiceDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceDomainOptions.Click, ctxmnutviDomainsDomainOptions.Click, hlpoptvObjects.Click
        ShowWindow(New wndDomains, True, Me, True)
        RefreshDomainTree()
    End Sub

    Private Sub mnuServicePreferences_Click(sender As Object, e As RoutedEventArgs) Handles mnuServicePreferences.Click
        ShowWindow(New wndPreferences, True, Me, True)
        RefreshDomainTree()
    End Sub

    Private Sub mnuServiceLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceLog.Click
        Dim w As New wndLog
        ShowWindow(w, True, Nothing, False)
    End Sub

    Private Sub mnuServiceErrorLog_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceErrorLog.Click
        Dim w As New wndErrorLog
        ShowWindow(w, True, Nothing, False)
    End Sub

    Private Sub mnuSearchSaveCurrentFilter_Click(sender As Object, e As RoutedEventArgs) Handles mnuSearchSaveCurrentFilter.Click
        If currentfilter Is Nothing OrElse String.IsNullOrEmpty(currentfilter.Filter) Then IMsgBox(My.Resources.wndMain_msg_CannotSaveCurrentFilter, vbOKOnly + vbExclamation,, Me) : Exit Sub

        Dim name As String = IInputBox(My.Resources.wndMain_msg_EnterFilterName,,, vbQuestion, Me)

        If String.IsNullOrEmpty(name) Then Exit Sub

        currentfilter.Name = name
        preferences.Filters.Add(currentfilter)
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As RoutedEventArgs) Handles mnuHelpAbout.Click
        Dim w As New wndAbout
        ShowWindow(w, True, Me, False)
    End Sub

#End Region

#Region "Context Menu"

    Private Sub tviDomainstviFavorites_TreeViewItem_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
        Dim obj As clsDirectoryObject = Nothing
        If TypeOf CType(sender, StackPanel).DataContext Is clsDirectoryObject Then obj = CType(CType(sender, StackPanel).DataContext, clsDirectoryObject)
        If obj Is Nothing Then Exit Sub

        If obj.IsDeleted Then e.Handled = True : Exit Sub

        CType(sender, StackPanel).ContextMenu.Tag = {obj}
    End Sub

    Private Sub tviFilters_TreeViewItem_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
        Dim flt As clsFilter = Nothing
        If TypeOf CType(sender, StackPanel).DataContext Is clsFilter Then flt = CType(CType(sender, StackPanel).DataContext, clsFilter)
        If flt Is Nothing Then Exit Sub
        CType(sender, StackPanel).ContextMenu.Tag = flt
    End Sub

    Private Sub lvObjects_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs) Handles lvObjects.ContextMenuOpening
        Dim objects As New List(Of clsDirectoryObject)

        If lvObjects.SelectedItems.Count = 0 Then
            If currentcontainer IsNot Nothing Then objects = New List(Of clsDirectoryObject) From {currentcontainer}
        ElseIf lvObjects.SelectedItems.Count = 1 Then
            If TypeOf lvObjects.SelectedItem Is clsDirectoryObject Then objects = New List(Of clsDirectoryObject) From {lvObjects.SelectedItem}
        ElseIf lvObjects.SelectedItems.Count > 1 Then
            Dim isobjectsflag As Boolean = True ' all selected objects is clsDirectoryObject
            Dim tmpobjs As New List(Of clsDirectoryObject)
            For Each obj In lvObjects.SelectedItems
                If TypeOf obj IsNot clsDirectoryObject Then isobjectsflag = False : Exit For
                tmpobjs.Add(obj)
            Next
            If isobjectsflag Then objects = tmpobjs
        End If

        Dim alldeleted As Boolean = objects.Count > 0 AndAlso objects.Where(Function(o) If(o.IsDeleted IsNot Nothing, o.IsDeleted = True, False)).Count = objects.Count
        ctxmnuObjectsRestore.Visibility = BooleanToVisibility(alldeleted)
        ctxmnuObjectsRestoreToContainer.Visibility = BooleanToVisibility(alldeleted)
        ctxmnuObjectsRestoreSeparator.Visibility = ctxmnuObjectsRestore.Visibility

        ctxmnuObjectsExternalSoftware.Visibility = BooleanToVisibility(preferences.ExternalSoftware.Count > 0 AndAlso objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer))

        ctxmnuObjectsSelectAll.Visibility = BooleanToVisibility(lvObjects.Items.Count > 1)
        ctxmnuObjectsSelectAllSeparator.Visibility = ctxmnuObjectsSelectAll.Visibility

        ctxmnuObjectsCreateObject.Visibility = BooleanToVisibility(True)

        ctxmnuObjectsCopy.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsCopySeparator.Visibility = ctxmnuObjectsCopy.Visibility
        ctxmnuObjectsCut.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsPaste.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.DomainDNS))
        ctxmnuObjectsPaste.IsEnabled = ClipboardBuffer IsNot Nothing AndAlso ClipboardBuffer.Count > 0
        ctxmnuObjectsRename.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User))
        ctxmnuObjectsRemove.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User))
        ctxmnuObjectsCopyData.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsCopyDataSeparator.Visibility = ctxmnuObjectsCopyData.Visibility
        ctxmnuObjectsFilterData.Visibility = BooleanToVisibility(objects.Count = 1)
        ctxmnuObjectsAddToFavorites.Visibility = BooleanToVisibility(objects.Count = 1)
        ctxmnuObjectsOpenObjectLocation.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso currentcontainer Is Nothing)
        ctxmnuObjectsOpenObjectTree.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso currentcontainer Is Nothing)

        ctxmnuObjectsResetPassword.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)
        ctxmnuObjectsDisableEnable.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer))
        ctxmnuObjectsExpirationDate.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)

        ctxmnuObjectsProperties.Visibility = BooleanToVisibility(objects.Count = 1)
        ctxmnuObjectsPropertiesSeparator.Visibility = ctxmnuObjectsProperties.Visibility

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

            ctxmnuObjectsFilterData.Items.Clear()
            If cvcurrentobjects.Filter IsNot Nothing Then
                Dim fdmnuclear As New MenuItem
                fdmnuclear.Header = My.Resources.wndMain_ctxmnuObjectsFilterDataClear
                fdmnuclear.Tag = New clsAttribute("", "", "")
                AddHandler fdmnuclear.Click, AddressOf ctxmnuObjectsFilterDataItem_Click
                ctxmnuObjectsFilterData.Items.Add(fdmnuclear)
                ctxmnuObjectsFilterData.Items.Add(New Separator)
            End If

            Dim statusvalue As clsDirectoryObject.enmStatus = GetType(clsDirectoryObject).GetProperty("Status").GetValue(objects(0))
            Dim fdmnustatus As New MenuItem
            fdmnustatus.Header = statusvalue.ToString
            fdmnustatus.Tag = New clsAttribute("Status", "Status", statusvalue)
            AddHandler fdmnustatus.Click, AddressOf ctxmnuObjectsFilterDataItem_Click
            ctxmnuObjectsFilterData.Items.Add(fdmnustatus)

            For Each columninfo In preferences.Columns
                For Each attr In columninfo.Attributes
                    If GetType(clsDirectoryObject).GetProperty(attr.Name).PropertyType Is GetType(String) Then
                        Dim value As String = GetType(clsDirectoryObject).GetProperty(attr.Name).GetValue(objects(0))
                        If Not String.IsNullOrEmpty(value) Then
                            Dim fdmnu As New MenuItem
                            fdmnu.Header = value
                            fdmnu.Tag = New clsAttribute(attr.Name, attr.Label, value)
                            AddHandler fdmnu.Click, AddressOf ctxmnuObjectsFilterDataItem_Click
                            ctxmnuObjectsFilterData.Items.Add(fdmnu)
                        End If
                    End If
                Next
            Next
        End If

        If objects.Count > 0 Then
            ctxmnuObjectsCopyData.Items.Clear()

            Dim cdmnudn As New MenuItem
            cdmnudn.Header = My.Resources.wndMain_ctxmnuObjectsCopyDataDisplayName
            cdmnudn.Tag = "displayName"
            AddHandler cdmnudn.Click, AddressOf ctxmnuObjectsCopyDataItem_Click
            ctxmnuObjectsCopyData.Items.Add(cdmnudn)

            Dim cdmnuba As New MenuItem
            cdmnuba.Header = My.Resources.wndMain_ctxmnuObjectsCopyDataBasicAttributes
            cdmnuba.Tag = "basicAttributes"
            AddHandler cdmnuba.Click, AddressOf ctxmnuObjectsCopyDataItem_Click
            ctxmnuObjectsCopyData.Items.Add(cdmnuba)

            ctxmnuObjectsCopyData.Items.Add(New Separator)

            For Each col As clsViewColumnInfo In preferences.Columns
                For Each attr In col.Attributes
                    Dim cdmnu As New MenuItem
                    cdmnu.Header = attr.Label
                    cdmnu.Tag = attr.Name
                    AddHandler cdmnu.Click, AddressOf ctxmnuObjectsCopyDataItem_Click
                    ctxmnuObjectsCopyData.Items.Add(cdmnu)
                Next
            Next
        End If

        lvObjects.ContextMenu.Tag = objects.ToArray
    End Sub

    Private Sub ctxmnuObjectsRestore_Click(sender As Object, e As RoutedEventArgs)
        'TODO
    End Sub

    Private Sub ctxmnuObjectsRestoreToContainer_Click(sender As Object, e As RoutedEventArgs)
        'TODO
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
            Dim val As String = If(objects(0).GetAttribute(pattern.Value.Replace("{{", "").Replace("}}", "")), pattern.Value).ToString
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

    Private Sub ctxmnuObjectsCopyDataItem_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(CType(sender, MenuItem).Parent, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(CType(sender, MenuItem).Parent, MenuItem).Parent, ContextMenu).Tag

        Dim cdmnu As MenuItem = CType(sender, MenuItem)
        If cdmnu Is Nothing Then Exit Sub
        Try
            If Not String.IsNullOrEmpty(cdmnu.Tag) AndAlso cdmnu.Tag = "basicAttributes" Then
                Clipboard.SetDataObject(Join(objects.Select(Function(o) o.displayName & vbTab & o.userPrincipalName & vbTab & o.mail & vbTab & o.telephoneNumber).ToArray, vbCrLf), True)
            Else
                Clipboard.SetDataObject(Join(objects.Select(Function(o) o.GetType().GetProperty(cdmnu.Tag).GetValue(o).ToString).ToArray, vbCrLf), True)
            End If
        Catch ex As Exception
            IMsgBox(ex.Message, vbExclamation)
        End Try
    End Sub

    Private Sub ctxmnuObjectsFilterDataItem_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(CType(sender, MenuItem).Parent, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(CType(sender, MenuItem).Parent, MenuItem).Parent, ContextMenu).Tag
        Dim cdmnu As MenuItem = CType(sender, MenuItem)
        If cdmnu Is Nothing Then Exit Sub
        If cdmnu.Tag Is Nothing Then Exit Sub

        If TypeOf cdmnu.Tag IsNot clsAttribute Then Exit Sub

        Dim attr As clsAttribute = cdmnu.Tag
        Try
            ApplyPostfilter(attr.Name, attr.Value)
        Catch ex As Exception
            IMsgBox(ex.Message, vbExclamation)
        End Try
    End Sub

    Private Sub ctxmnuObjectsSelectAll_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuObjectsSelectAll.Click
        lvObjects.SelectAll()
    End Sub

    Private Sub ctxmnuSharedCreateObject_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag

        Dim w As New wndCreateObject

        If objects.Count = 1 AndAlso (
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.DomainDNS Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.UnknownContainer) Then

            w.destinationcontainer = objects(0)
            w.destinationdomain = objects(0).Domain
        Else
            w.destinationcontainer = Nothing
            w.destinationdomain = Nothing
        End If

        ShowWindow(w, False, Me, False)
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
                Return obj.SchemaClass = clsDirectoryObject.enmSchemaClass.User Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit
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
                Return obj.SchemaClass = clsDirectoryObject.enmSchemaClass.User Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit
            End Function).ToArray

        ClipboardAction = enmClipboardAction.Cut
    End Sub

    Private Sub ctxmnuSharedPaste_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim destinations() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (destinations.Count = 1 AndAlso (
            destinations(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or
            destinations(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or
            destinations(0).SchemaClass = clsDirectoryObject.enmSchemaClass.UnknownContainer)) Then Exit Sub

        If Not (ClipboardBuffer IsNot Nothing AndAlso ClipboardBuffer.Count > 0) Then Exit Sub

        Dim destination As clsDirectoryObject = destinations(0)
        Dim sourceobjects() As clsDirectoryObject = ClipboardBuffer

        If ClipboardAction = enmClipboardAction.Copy Then ' copy

            ObjectsCopy(destination, sourceobjects)

        ElseIf ClipboardAction = enmClipboardAction.Cut Then ' cut

            ObjectsMove(destination, sourceobjects)

        End If

    End Sub

    Private Sub ctxmnuSharedRename_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit)) Then Exit Sub

        Try
            Dim organizationalunitaffected As Boolean = False

            Dim obj As clsDirectoryObject = objects(0)
            Dim name As String = IInputBox(My.Resources.wndMain_msg_EnterObjectName, My.Resources.wndMain_msg_RenameObject, objects(0).name, vbQuestion, Me)
            If Len(name) > 0 Then
                If obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Then
                    obj.Rename("OU=" & name)
                    organizationalunitaffected = True
                Else
                    obj.Rename("CN=" & name)
                End If

                RefreshDataGrid()
                If organizationalunitaffected Then RefreshDomainTree()
            End If
        Catch ex As Exception
            ThrowException(ex, "ctxmnuSharedRename_Click")
        End Try
    End Sub

    Private Sub ctxmnuSharedRemove_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit)) Then Exit Sub

        Try
            Dim currentcontaineraffected As Boolean = False

            If IMsgBox(My.Resources.wndMain_msg_AreYouSure, vbYesNo + vbQuestion, My.Resources.wndMain_msg_RemoveObject, Me) <> MsgBoxResult.Yes Then Exit Sub
            If objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Then
                If IMsgBox(My.Resources.wndMain_msg_ThisIsOrganizationaUnit & vbCrLf & vbCrLf & My.Resources.wndMain_msg_AreYouSure, vbYesNo + vbExclamation, My.Resources.wndMain_msg_RemoveObject, Me) <> MsgBoxResult.Yes Then Exit Sub
                If currentcontainer IsNot Nothing AndAlso objects(0).distinguishedName = currentcontainer.distinguishedName Then currentcontaineraffected = True
            End If

            objects(0).DeleteTree()

            If currentcontaineraffected Then
                OpenObjectParent()
                RefreshDomainTree()
            Else
                RefreshDataGrid()
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
        If objects.Count = 1 Then OpenObject(objects(0).Parent)
    End Sub

    Private Sub ctxmnuSharedOpenObjectTree_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 Then ShowInTree(objects(0).Parent)
    End Sub

    Private Sub ctxmnuSharedResetPassword_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)) Then Exit Sub

        Try
            If IMsgBox(My.Resources.wndObject_msg_AreYouSure, vbYesNo + vbQuestion, My.Resources.wndObject_msg_ResetPassword, Me) = MsgBoxResult.Yes Then
                objects(0).ResetPassword()
                objects(0).passwordNeverExpires = False
                IMsgBox(My.Resources.wndObject_msg_PasswordChanged, vbOKOnly + vbInformation, My.Resources.wndObject_msg_ResetPassword)
            End If

        Catch ex As Exception
            ThrowException(ex, "ctxmnuSharedResetPassword_Click")
        End Try
    End Sub

    Private Sub ctxmnuSharedDisableEnable_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If Not (objects.Count = 1 AndAlso (
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer)) Then Exit Sub

        Try
            If IMsgBox(My.Resources.wndObject_msg_AreYouSure, vbYesNo + vbQuestion, If(objects(0).disabled, My.Resources.wndObject_msg_Enable, My.Resources.wndObject_msg_Disable), Me) = MsgBoxResult.Yes Then
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
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)) Then Exit Sub

        Try
            objects(0).accountExpiresDate = Today.AddDays(1)
            Dim w As Window = ShowDirectoryObjectProperties(objects(0), Me)
            If GetType(wndUser) Is w.GetType Then CType(w, wndUser).tabctlUser.SelectedIndex = 1

        Catch ex As Exception
            ThrowException(ex, "ctxmnuSharedExpirationDate_Click")
        End Try
    End Sub

    Private Sub ctxmnuSharedProperties_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 Then ShowDirectoryObjectProperties(objects(0), Window.GetWindow(Me))
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

    Private Sub clipboardTimer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles clipboardTimer.Tick
        If preferences Is Nothing OrElse preferences.ClipboardSource = False Then Exit Sub
        Dim newclipboarddata As String = Clipboard.GetText
        If String.IsNullOrEmpty(newclipboarddata) Then Exit Sub
        If preferences.ClipboardSourceLimit AndAlso CountWords(newclipboarddata) > 3 Then Exit Sub ' only three words

        If clipboardlastdata <> newclipboarddata Then
            clipboardlastdata = newclipboarddata
            StartSearch(Nothing, New clsFilter(clipboardlastdata, preferences.AttributesForSearch, searchobjectclasses))
        End If
    End Sub

    Private Sub tviDomainstviFavorites_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)

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

    Private Sub tviDomains_TreeViewItem_DragEnterDragOver(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            If e.KeyStates.HasFlag(DragDropKeyStates.ControlKey) Then
                e.Effects = DragDropEffects.Copy
            Else
                e.Effects = DragDropEffects.Move
            End If

            For Each obj As clsDirectoryObject In e.Data.GetData(GetType(clsDirectoryObject()))
                If Not (obj.SchemaClass = clsDirectoryObject.enmSchemaClass.User Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit) Then e.Effects = DragDropEffects.None : Exit For
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

        If Not (destination.SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or
            destination.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or
            destination.SchemaClass = clsDirectoryObject.enmSchemaClass.UnknownContainer) Then Exit Sub

        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            Dim dropped = e.Data.GetData(GetType(clsDirectoryObject()))
            For Each obj As clsDirectoryObject In dropped
                If Not (obj.SchemaClass = clsDirectoryObject.enmSchemaClass.User Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or
                       obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit) Then Exit Sub
            Next

            If e.KeyStates.HasFlag(DragDropKeyStates.ControlKey) Then

                ObjectsCopy(destination, dropped)

            Else

                ObjectsMove(destination, dropped)

            End If

        End If
    End Sub

    Private Sub lvObjects_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles lvObjects.PreviewKeyDown
        Select Case e.Key
            Case Key.Enter
                If lvObjects.SelectedItem Is Nothing Then Exit Sub
                Dim current As clsDirectoryObject = CType(lvObjects.SelectedItem, clsDirectoryObject)
                OpenObject(current)
                e.Handled = True
            Case Key.Back
                OpenObjectParent()
                e.Handled = True
            Case Key.Home
                If lvObjects.Items.Count > 0 Then lvObjects.SelectedIndex = 0
                e.Handled = True
            Case Key.End
                If lvObjects.Items.Count > 0 Then lvObjects.SelectedIndex = lvObjects.Items.Count - 1
                e.Handled = True
        End Select
    End Sub

    Private Sub lvObjects_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvObjects.MouseDoubleClick
        If lvObjects.SelectedItem Is Nothing Or Not (e.ChangedButton = MouseButton.Left) Then Exit Sub
        Dim current As clsDirectoryObject = CType(lvObjects.SelectedItem, clsDirectoryObject)
        OpenObject(current)
    End Sub

    Private Sub lvObjects_ColumnReordered_LayoutUpdated() Handles lvObjects.LayoutUpdated
        Dim view = lvObjects.View
        If view Is Nothing OrElse Not TypeOf view Is GridView Then Exit Sub
        Dim gridview As GridView = CType(view, GridView)
        For Each dgcolumn As GridViewColumn In gridview.Columns
            For Each pcolumn As clsViewColumnInfo In preferences.Columns
                If dgcolumn.Header.ToString = pcolumn.Header Then
                    'pcolumn.DisplayIndex = dgcolumn.DisplayIndex
                    pcolumn.Width = dgcolumn.ActualWidth
                End If
            Next
        Next
    End Sub

    Private Sub lvObjects_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles lvObjects.MouseDown
        'If (e.ChangedButton = MouseButton.Left Or e.ChangedButton = MouseButton.Right) AndAlso Keyboard.Modifiers = 0 Then
        '    Dim r As HitTestResult = VisualTreeHelper.HitTest(sender, e.GetPosition(sender))
        '    If r IsNot Nothing Then
        '        If TypeOf r.VisualHit Is ScrollViewer Then
        '            lvObjects.UnselectAll()
        '        Else
        '            Dim dp As DependencyObject = r.VisualHit
        '            While (dp IsNot Nothing) AndAlso Not (TypeOf dp Is DataGridRow)
        '                dp = VisualTreeHelper.GetParent(dp)
        '            End While
        '            If dp Is Nothing Then Return

        '            If TypeOf dp Is DataGridRow Then
        '                lvObjects.SelectedItem = CType(dp, DataGridRow).DataContext
        '            End If
        '        End If
        '    End If
        'End If
        If e.ChangedButton = MouseButton.XButton1 Then SearchPrevious()
        If e.ChangedButton = MouseButton.XButton2 Then SearchNext()
    End Sub

    'Private Sub lvObjects_MouseMove(sender As Object, e As MouseEventArgs) Handles lvObjects.MouseMove
    '    Dim datagrid As DataGrid = TryCast(sender, DataGrid)

    '    If e.LeftButton = MouseButtonState.Pressed And
    '        e.GetPosition(sender).X < datagrid.ActualWidth - SystemParameters.VerticalScrollBarWidth And
    '        e.GetPosition(sender).Y < datagrid.ActualHeight - SystemParameters.HorizontalScrollBarHeight And
    '        datagrid.SelectedItems.Count > 0 Then

    '        Dim dragData As New DataObject(datagrid.SelectedItems.Cast(Of clsDirectoryObject).ToArray)

    '        DragDrop.DoDragDrop(datagrid, dragData, DragDropEffects.All)
    '    End If
    'End Sub

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
        If currentcontainer Is Nothing Then Exit Sub
        Dim current As clsDirectoryObject = New clsDirectoryObject(CType(sender.Tag, String), currentcontainer.Domain)
        OpenObject(current)
    End Sub

    Private Sub tbSearchPattern_OnPaste(sender As Object, e As DataObjectPastingEventArgs)
        Dim istext = e.SourceDataObject.GetDataPresent(DataFormats.UnicodeText, True)
        If Not istext Then Exit Sub

        Dim texttopaste As String = e.SourceDataObject.GetData(DataFormats.UnicodeText).ToString.Replace(vbNewLine, " / ")

        Dim start As Integer = tbSearchPattern.SelectionStart
        Dim length As Integer = tbSearchPattern.SelectionLength
        Dim caret As Integer = tbSearchPattern.CaretIndex

        Dim text As String = tbSearchPattern.Text.Substring(0, start)
        text += tbSearchPattern.Text.Substring(start + length)

        Dim newText As String = text.Substring(0, tbSearchPattern.CaretIndex) + texttopaste
        newText += text.Substring(caret)
        tbSearchPattern.Text = newText
        tbSearchPattern.CaretIndex = caret + texttopaste.Length

        e.CancelCommand()
    End Sub

    Private Sub tbSearchPattern_KeyDown(sender As Object, e As KeyEventArgs) Handles tbSearchPattern.KeyDown
        If e.Key = Key.Enter Then
            If mnuSearchModeDefault.IsChecked = True Then
                StartSearch(Nothing, New clsFilter(tbSearchPattern.Text, preferences.AttributesForSearch, searchobjectclasses))
            ElseIf mnuSearchModeAdvanced.IsChecked = True Then
                StartSearch(Nothing, New clsFilter(tbSearchPattern.Text))
            End If
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If mnuSearchModeDefault.IsChecked = True Then
            StartSearch(Nothing, New clsFilter(tbSearchPattern.Text, preferences.AttributesForSearch, searchobjectclasses))
        ElseIf mnuSearchModeAdvanced.IsChecked = True Then
            StartSearch(Nothing, New clsFilter(tbSearchPattern.Text))
        End If
        tbSearchPattern.Focus()
    End Sub

    Private Sub pbSearch_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles pbSearch.MouseDoubleClick
        StopSearch()
    End Sub

    Private Sub imgFilterStatus_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles imgFilterStatus.MouseLeftButtonDown
        ApplyPostfilter()
    End Sub

    Private Sub btnWindowClone_Click(sender As Object, e As RoutedEventArgs) Handles btnWindowClone.Click
        Dim w As New wndMain
        w.Show()
    End Sub

    Private Sub btnDummy_Click(sender As Object, e As RoutedEventArgs) Handles btnDummy.Click
        UpdateDetailsViewGroupStyle(True)
    End Sub

    Private Sub btnViewPreferences_Click(sender As Object, e As RoutedEventArgs) Handles btnViewPreferences.Click
        Dim w As New wndPreferences
        w.InitializeComponent()
        w.tcPreferences.SelectedIndex = 2
        ShowWindow(New wndPreferences, True, Me, True)
    End Sub

#End Region

#Region "Subs"

    Public Sub RefreshDomainTree()
        tviDomains.ItemsSource = domains.Where(Function(d As clsDomain) d.Validated).Select(Function(d As clsDomain) New clsDirectoryObject(d.DefaultNamingContext, d))
    End Sub

    Public Sub UpdateDetailsViewGroupStyle(show As Boolean)
        If show Then
            Dim groupstyle = New GroupStyle()
            groupstyle.ContainerStyle = TryFindResource("ListView_ViewDetails_GroupItem")
            If groupstyle.ContainerStyle Is Nothing Then Exit Sub
            Debug.WriteLine(lvObjects.Style.ToString)
            lvObjects.GroupStyle.Clear()
            lvObjects.GroupStyle.Add(groupstyle)
        Else
            lvObjects.GroupStyle.Clear()
        End If
    End Sub

    Public Sub OpenObject(current As clsDirectoryObject)
        If (current.SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or
           current.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or
           current.SchemaClass = clsDirectoryObject.enmSchemaClass.UnknownContainer Or
           current.SchemaClass = clsDirectoryObject.enmSchemaClass.DomainDNS) AndAlso
           (current.IsRecycled Is Nothing OrElse current.IsRecycled = False) Then
            StartSearch(current, Nothing)
        Else
            ShowDirectoryObjectProperties(current, Window.GetWindow(Me))
        End If
    End Sub

    Public Sub OpenObjectParent()
        If currentcontainer Is Nothing OrElse (currentcontainer.Parent Is Nothing OrElse currentcontainer.distinguishedName = currentcontainer.Domain.DefaultNamingContext) Then Exit Sub
        StartSearch(currentcontainer.Parent, Nothing)
    End Sub

    Private Sub ShowPath(Optional container As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)

        For Each child In spPath.Children
            If TypeOf child Is Button Then RemoveHandler CType(child, Button).Click, AddressOf btnPath_Click
        Next

        spPath.Children.Clear()

        If container IsNot Nothing Then

            Dim buttons As New List(Of Button)

            Dim containerDN As String = container.distinguishedName

            Do
                Dim btn As New Button
                Dim st As Style = Application.Current.TryFindResource("ToolbarButton")
                btn.Style = st
                btn.Content = GetNameFromDN(containerDN)
                btn.Margin = New Thickness(2, 0, 2, 0)
                btn.Padding = New Thickness(5, 0, 5, 0)
                btn.VerticalAlignment = VerticalAlignment.Stretch
                btn.Tag = containerDN
                buttons.Add(btn)

                If containerDN.StartsWith("DC=", StringComparison.OrdinalIgnoreCase) Then
                    Exit Do
                Else
                    containerDN = GetParentDNFromDN(containerDN)
                End If
            Loop

            buttons.Reverse()
            spPath.Children.Add(New Shapes.Path() With {.Style = FindResource("PathSeparator")})
            For Each child As Button In buttons
                AddHandler child.Click, AddressOf btnPath_Click
                spPath.Children.Add(child)
                spPath.Children.Add(New Shapes.Path() With {.Style = FindResource("PathSeparator")})
            Next

        ElseIf filter IsNot Nothing Then

            Dim tblck As New TextBlock
            tblck.VerticalAlignment = VerticalAlignment.Center
            tblck.Background = Brushes.Transparent
            tblck.Text = My.Resources.wndMain_lbl_SearchResults & " " & If(Not String.IsNullOrEmpty(filter.Name), filter.Name, If(Not String.IsNullOrEmpty(filter.Pattern), filter.Pattern.Replace(vbNewLine, " "), " Advanced filter"))
            tblck.Margin = New Thickness(2, 0, 2, 0)
            tblck.Padding = New Thickness(5, 0, 5, 0)
            spPath.Children.Add(tblck)

        End If
    End Sub

    Private Async Sub ShowInTree(container As clsDirectoryObject)
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

    Private Sub ApplyPostfilter(Optional prop As String = Nothing, Optional value As String = Nothing)
        If String.IsNullOrEmpty(prop) Then
            imgFilterStatus.Visibility = Visibility.Collapsed
            cvcurrentobjects.Filter = Nothing
        Else
            imgFilterStatus.Visibility = Visibility.Visible
            cvcurrentobjects.Filter = New Predicate(Of Object)(
                Function(obj As clsDirectoryObject)
                    Return GetType(clsDirectoryObject).GetProperty(prop).GetValue(obj) = value
                End Function)
        End If
    End Sub

    Public Sub RefreshDataGrid()
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

    Public Sub StopSearch()
        searcher.StopAllSearchAsync()
    End Sub

    Public Sub StartSearch(Optional root As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        If root Is Nothing And (filter Is Nothing OrElse String.IsNullOrEmpty(filter.Filter)) Then Exit Sub

        While searchhistory.Count > searchhistoryindex + 1
            searchhistory.RemoveAt(searchhistory.Count - 1)
        End While

        searchhistory.Add(New clsSearchHistory(root, filter))
        searchhistoryindex = searchhistory.Count - 1

        Search(root, filter)
    End Sub

    Public Sub Search(Optional root As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        tbSearchPattern.SelectAll()
        'UpdateDetailsViewGroupStyle(False)

        ApplyPostfilter(Nothing, Nothing)

        cap.Visibility = Visibility.Visible
        pbSearch.Visibility = Visibility.Visible

        Dim domainlist As New ObservableCollection(Of clsDomain)(domains.Where(Function(x As clsDomain) x.IsSearchable = True).ToList)

        currentcontainer = root
        currentfilter = filter
        ShowPath(currentcontainer, filter)

        searcher.SearchAsync(currentobjects, root, filter, domainlist)
    End Sub

    Private Sub Searcher_SearchAsyncDataRecieved() Handles searcher.SearchAsyncDataRecieved
        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub Searcher_SearchAsyncCompleted() Handles searcher.SearchAsyncCompleted
        cap.Visibility = Visibility.Hidden
        pbSearch.Visibility = Visibility.Hidden

        'If preferences.SearchResultGrouping AndAlso currentcontainer Is Nothing Then
        '    UpdateDetailsViewGroupStyle(True)
        'End If
    End Sub

    Public Function CreateColumn(columninfo As clsViewColumnInfo) As DataGridTemplateColumn
        Dim column As New DataGridTemplateColumn()
        column.Header = columninfo.Header
        column.SetValue(DataGridColumn.CanUserSortProperty, True)
        If columninfo.DisplayIndex > 0 Then column.DisplayIndex = columninfo.DisplayIndex
        column.Width = If(columninfo.Width > 0, columninfo.Width, 0)
        column.MinWidth = 58
        Dim panel As New FrameworkElementFactory(GetType(VirtualizingStackPanel))
        panel.SetValue(VerticalAlignmentProperty, VerticalAlignment.Center)
        panel.SetValue(MarginProperty, New Thickness(5, 0, 5, 0))

        Dim first As Boolean = True
        For Each attr As clsAttribute In columninfo.Attributes
            Dim bind As New Binding(attr.Name) With {.Mode = BindingMode.OneWay, .Converter = New ConverterDataToUIElement, .ConverterParameter = attr.Name}


            Dim container As New FrameworkElementFactory(GetType(ItemsControl))
            If first Then
                first = False
                container.SetValue(FontWeightProperty, FontWeights.Bold)
                column.SetValue(DataGridColumn.SortMemberPathProperty, attr.Name)
            End If

            container.SetBinding(ItemsControl.ItemsSourceProperty, bind)
            container.SetValue(ToolTipProperty, attr.Label)
            panel.AppendChild(container)
        Next

        Dim template As New DataTemplate()
        template.VisualTree = panel
        column.CellTemplate = template

        Return column
    End Function

    Public Sub ObjectsCopy(destination As clsDirectoryObject, sourceobjects() As clsDirectoryObject)
        If sourceobjects(0).Domain Is destination.Domain Then ' same domain
            If sourceobjects.Count > 1 Then IMsgBox(My.Resources.wndMain_msg_CopyOneObjectWithinDomain, vbOKOnly + vbExclamation, My.Resources.wndMain_msg_CopyObject, Me) : Exit Sub

            ' single object

            Dim w As New wndCreateObject
            w.destinationcontainer = destination
            w.destinationdomain = destination.Domain

            Select Case sourceobjects(0).SchemaClass
                Case clsDirectoryObject.enmSchemaClass.User
                    w.copyingobject = sourceobjects(0)
                    w.tbUserDisplayname.Text = sourceobjects(0).displayName
                    w.tbUserObjectName.Text = sourceobjects(0).displayName & "_copy"
                    w.cmboUserUserPrincipalName.Text = sourceobjects(0).userPrincipalNameName & "_copy"
                    w.cmboUserUserPrincipalNameDomain.SelectedItem = sourceobjects(0).userPrincipalNameDomain
                    w.tabctlObject.SelectedIndex = 0

                Case clsDirectoryObject.enmSchemaClass.Computer
                    w.copyingobject = sourceobjects(0)
                    w.cmboComputerObjectName.Text = sourceobjects(0).name & "_copy"
                    w.tabctlObject.SelectedIndex = 1

                Case clsDirectoryObject.enmSchemaClass.Group
                    w.copyingobject = sourceobjects(0)
                    w.rbGroupScopeDomainLocal.IsChecked = sourceobjects(0).groupTypeScopeDomainLocal
                    w.rbGroupScopeGlobal.IsChecked = sourceobjects(0).groupTypeScopeGlobal
                    w.rbGroupScopeUniversal.IsChecked = sourceobjects(0).groupTypeScopeUniversal
                    w.rbGroupTypeSecurity.IsChecked = sourceobjects(0).groupTypeSecurity
                    w.rbGroupTypeDistribution.IsChecked = sourceobjects(0).groupTypeDistribution
                    w.tbGroupObjectName.Text = sourceobjects(0).name & "_copy"
                    w.tabctlObject.SelectedIndex = 2

                Case clsDirectoryObject.enmSchemaClass.Contact
                    w.copyingobject = sourceobjects(0)
                    w.tbContactDisplayname.Text = sourceobjects(0).displayName
                    w.tbContactObjectName.Text = sourceobjects(0).displayName & "_copy"
                    w.tabctlObject.SelectedIndex = 3

                Case clsDirectoryObject.enmSchemaClass.OrganizationalUnit
                    w.copyingobject = sourceobjects(0)
                    w.tbOrganizationalUnitObjectName.Text = sourceobjects(0).name & "_copy"
                    w.tabctlObject.SelectedIndex = 4

                Case Else
                    IMsgBox(My.Resources.wndMain_msg_ObjectUnknownClass, vbOKOnly + vbExclamation, My.Resources.wndMain_msg_CopyObject, Me)
            End Select
            ShowWindow(w, False, Me, False)

        Else ' another domain
            ' TODO
            If IMsgBox("А это не допилено еще", vbYesNo + vbQuestion, "Вставка", Me) <> MessageBoxResult.Yes Then Exit Sub

        End If
    End Sub

    Public Sub ObjectsMove(destination As clsDirectoryObject, sourceobjects() As clsDirectoryObject)
        Dim organizationalunitaffected As Boolean = False

        ' another domain
        If sourceobjects(0).Domain IsNot destination.Domain Then IMsgBox(My.Resources.wndMain_msg_MovingIsProhibited, vbOKOnly + vbExclamation, My.Resources.wndMain_msg_MoveObject, Me) : Exit Sub

        ' same domain
        If IMsgBox(My.Resources.wndMain_msg_AreYouSure & vbCrLf & vbCrLf &
                   String.Format(My.Resources.wndMain_msg_MoveObjectToContainer,
                                 If(sourceobjects.Count > 1, String.Format(My.Resources.wndMain_msg_NObjects, sourceobjects.Count), sourceobjects(0).name),
                                 destination.distinguishedNameFormated), vbYesNo + vbQuestion, My.Resources.wndMain_msg_MoveObject, Me) <> MessageBoxResult.Yes Then Exit Sub

        Try
            For Each obj In sourceobjects
                If obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Then organizationalunitaffected = True

                obj.MoveTo(destination.distinguishedName)
            Next
        Catch ex As Exception
            ThrowException(ex, "ObjectsMove")
        End Try

        RefreshDataGrid()
        If organizationalunitaffected Then RefreshDomainTree()
    End Sub

    Private Sub rbViewMediumIcons_Checked(sender As Object, e As RoutedEventArgs) Handles rbViewMediumIcons.Checked, rbViewList.Checked, rbViewTiles.Checked, rbViewDetails.Checked
        If sender Is rbViewMediumIcons Then
            UpdateDetailsViewGroupStyle(False)
            lvObjects.SetResourceReference(StyleProperty, "ListView_ViewMediumIcons")
        ElseIf sender Is rbViewList Then
            UpdateDetailsViewGroupStyle(False)
            lvObjects.SetResourceReference(StyleProperty, "ListView_ViewList")
        ElseIf sender Is rbViewTiles Then
            UpdateDetailsViewGroupStyle(False)
            lvObjects.SetResourceReference(StyleProperty, "ListView_ViewTiles")
        ElseIf sender Is rbViewDetails Then
            UpdateDetailsViewGroupStyle(True)
            lvObjects.SetResourceReference(StyleProperty, "ListView_ViewDetails")
        End If
    End Sub


#End Region

End Class
