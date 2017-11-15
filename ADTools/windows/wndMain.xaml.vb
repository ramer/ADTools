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
    Public Shared hkF1 As New RoutedCommand
    Public Shared hkEsc As New RoutedCommand

    Public Property currentcontainer As clsDirectoryObject
    Public Property currentobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)
    Public Property currentfilter As clsFilter

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
        RebuildColumns()

        DataObject.AddPastingHandler(cmboSearchPattern, AddressOf cmboSearchPattern_OnPaste)
        cmboSearchPattern.ItemsSource = ADToolsApplication.ocGlobalSearchHistory
        cmboSearchPattern.Focus()

        mnuSearchDomains.ItemsSource = domains
        tviFavorites.ItemsSource = preferences.Favorites
        tviFilters.ItemsSource = preferences.Filters

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
        popcmboSearchPattern.IsOpen = True
        popF1Hint.IsOpen = True
    End Sub

    Private Sub MovePopups() Handles Me.LocationChanged, Me.SizeChanged, tvObjects.SizeChanged, cmboSearchPattern.SizeChanged
        poptvObjects.HorizontalOffset += 1 : poptvObjects.HorizontalOffset -= 1
        popcmboSearchPattern.HorizontalOffset += 1 : popcmboSearchPattern.HorizontalOffset -= 1
        popF1Hint.HorizontalOffset += 1 : popF1Hint.HorizontalOffset -= 1
    End Sub

    Private Sub ClosePopups()
        poptvObjects.IsOpen = False
        popcmboSearchPattern.IsOpen = False
        popF1Hint.IsOpen = False
    End Sub

    Private Sub ClosePopup(sender As Object, e As MouseButtonEventArgs) Handles poptvObjects.MouseLeftButtonDown, popcmboSearchPattern.MouseLeftButtonDown, popF1Hint.MouseLeftButtonDown
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

    Private Sub tviDomains_TreeViewItem_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
        Dim obj As clsDirectoryObject = Nothing
        If TypeOf CType(sender, StackPanel).DataContext Is clsDirectoryObject Then obj = CType(CType(sender, StackPanel).DataContext, clsDirectoryObject)
        If obj Is Nothing Then Exit Sub

        CType(sender, StackPanel).ContextMenu.Tag = {obj}
    End Sub

    Private Sub dgObjects_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs) Handles dgObjects.ContextMenuOpening
        Dim objects As New List(Of clsDirectoryObject)

        If dgObjects.SelectedItems.Count = 0 Then
            If currentcontainer IsNot Nothing Then objects = New List(Of clsDirectoryObject) From {currentcontainer}
        ElseIf dgObjects.SelectedItems.Count = 1 Then
            If TypeOf dgObjects.SelectedItem Is clsDirectoryObject Then objects = New List(Of clsDirectoryObject) From {dgObjects.SelectedItem}
        ElseIf dgObjects.SelectedItems.Count > 1 Then
            Dim isobjectsflag As Boolean = True ' all selected objects is clsDirectoryObject
            Dim tmpobjs As New List(Of clsDirectoryObject)
            For Each obj In dgObjects.SelectedItems
                If TypeOf obj IsNot clsDirectoryObject Then isobjectsflag = False : Exit For
                tmpobjs.Add(obj)
            Next
            If isobjectsflag Then objects = tmpobjs
        End If

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
        ctxmnuObjectsSelectAllSeparator.Visibility = ctxmnuObjectsSelectAll.Visibility

        ctxmnuObjectsCreateObject.Visibility = BooleanToVisibility(True)

        ctxmnuObjectsCopy.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsCopySeparator.Visibility = ctxmnuObjectsCopy.Visibility
        ctxmnuObjectsCut.Visibility = BooleanToVisibility(objects.Count > 0)
        ctxmnuObjectsPaste.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.DomainDNS))
        ctxmnuObjectsPaste.IsEnabled = ClipboardBuffer IsNot Nothing AndAlso ClipboardBuffer.Count > 0
        ctxmnuObjectsRename.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User))
        ctxmnuObjectsRemove.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Group Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User))
        ctxmnuObjectsAddToFavorites.Visibility = BooleanToVisibility(objects.Count = 1)
        ctxmnuObjectsOpenObjectLocation.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso currentcontainer Is Nothing)
        ctxmnuObjectsOpenObjectTree.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso currentcontainer Is Nothing)
        ctxmnuObjectsAddToFavoritesSeparator.Visibility = ctxmnuObjectsAddToFavorites.Visibility

        ctxmnuObjectsResetPassword.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)
        ctxmnuObjectsDisableEnable.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso (objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer))
        ctxmnuObjectsExpirationDate.Visibility = BooleanToVisibility(objects.Count = 1 AndAlso objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)

        ctxmnuObjectsProperties.Visibility = BooleanToVisibility(objects.Count = 1)
        ctxmnuObjectsPropertiesSeparator.Visibility = ctxmnuObjectsProperties.Visibility

        dgObjects.ContextMenu.Tag = objects.ToArray
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
        Clipboard.SetDataObject(Join(objects.Select(Function(o) o.name & vbTab & o.userPrincipalName & vbTab & o.telephoneNumber).ToArray, vbCrLf), True)

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
        Clipboard.SetDataObject(Join(objects.Select(Function(o) o.name & vbTab & o.userPrincipalName & vbTab & o.telephoneNumber).ToArray, vbCrLf), True)

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
                    obj.Entry.Rename("OU=" & name)
                    organizationalunitaffected = True
                Else
                    obj.Entry.Rename("CN=" & name)
                End If
                obj.Entry.CommitChanges()
                obj.NotifyRenamed()
                obj.NotifyMoved()

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
                If currentcontainer IsNot Nothing AndAlso objects(0).Entry.Path = currentcontainer.Entry.Path Then currentcontaineraffected = True
            End If

            objects(0).Entry.DeleteTree()

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
        If objects.Count = 1 Then OpenObject(New clsDirectoryObject(objects(0).Entry.Parent, objects(0).Domain))
    End Sub

    Private Sub ctxmnuSharedOpenObjectTree_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag
        If objects.Count = 1 Then ShowTree(New clsDirectoryObject(objects(0).Entry.Parent, objects(0).Domain))
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

    Private Sub tviDomains_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)

        If TypeOf sp.Tag Is clsDirectoryObject Then
            OpenObject(CType(sp.Tag, clsDirectoryObject))
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
            If r IsNot Nothing AndAlso TypeOf r.VisualHit Is ScrollViewer Then
                dgObjects.UnselectAll()
            End If
        End If
        If e.ChangedButton = MouseButton.XButton1 Then SearchPrevious()
        If e.ChangedButton = MouseButton.XButton2 Then SearchNext()
    End Sub

    Private Sub dgObjects_MouseMove(sender As Object, e As MouseEventArgs) Handles dgObjects.MouseMove
        Dim datagrid As DataGrid = TryCast(sender, DataGrid)

        If e.LeftButton = MouseButtonState.Pressed And
            e.GetPosition(sender).X < datagrid.ActualWidth - SystemParameters.VerticalScrollBarWidth And
            e.GetPosition(sender).Y < datagrid.ActualHeight - SystemParameters.HorizontalScrollBarHeight And
            datagrid.SelectedItems.Count > 0 Then

            Dim dragData As New DataObject(datagrid.SelectedItems.Cast(Of clsDirectoryObject).ToArray)

            DragDrop.DoDragDrop(datagrid, dragData, DragDropEffects.All)
        End If
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

    Private Sub cmboSearchPattern_OnPaste(sender As Object, e As DataObjectPastingEventArgs)
        Dim istext = e.SourceDataObject.GetDataPresent(DataFormats.UnicodeText, True)
        If Not istext Then Exit Sub

        Dim texttopase As String = e.SourceDataObject.GetData(DataFormats.UnicodeText).ToString.Replace(vbNewLine, " / ")
        Dim cmboTextBoxChild As TextBox = cmboSearchPattern.Template.FindName("PART_EditableTextBox", cmboSearchPattern)

        Dim start As Integer = cmboTextBoxChild.SelectionStart
        Dim length As Integer = cmboTextBoxChild.SelectionLength
        Dim caret As Integer = cmboTextBoxChild.CaretIndex

        Dim text As String = cmboTextBoxChild.Text.Substring(0, start)
        text += cmboTextBoxChild.Text.Substring(start + length)

        Dim newText As String = text.Substring(0, cmboTextBoxChild.CaretIndex) + texttopase
        newText += text.Substring(caret)
        cmboTextBoxChild.Text = newText
        cmboTextBoxChild.CaretIndex = caret + texttopase.Length

        e.CancelCommand()
    End Sub

    Private Sub cmboSearchPattern_KeyDown(sender As Object, e As KeyEventArgs) Handles cmboSearchPattern.KeyDown
        If e.Key = Key.Enter Then
            If mnuSearchModeDefault.IsChecked = True Then
                StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text, preferences.AttributesForSearch, searchobjectclasses))
            ElseIf mnuSearchModeAdvanced.IsChecked = True Then
                StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text))
            End If
        Else
            If cmboSearchPattern.IsDropDownOpen Then cmboSearchPattern.IsDropDownOpen = False
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If mnuSearchModeDefault.IsChecked = True Then
            StartSearch(Nothing, New clsFilter(cmboSearchPattern.Text, preferences.AttributesForSearch, searchobjectclasses))
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

    Private Sub btnDummy_Click(sender As Object, e As RoutedEventArgs) Handles btnDummy.Click
        Throw New StackOverflowException("some bad data")
    End Sub

#End Region

#Region "Subs"

    Public Sub RefreshDomainTree()
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
        For Each child In spPath.Children
            If TypeOf child Is Button Then RemoveHandler CType(child, Button).Click, AddressOf btnPath_Click
        Next

        spPath.Children.Clear()

        If container IsNot Nothing Then

            Dim buttons As New List(Of Button)

            Do
                Dim btn As New Button
                Dim st As Style = Application.Current.TryFindResource("ToolbarButton")
                btn.Style = st
                btn.Content = If(container.objectClass.Contains("domaindns"), container.Domain.Name, container.name)
                btn.Margin = New Thickness(2, 0, 2, 0)
                btn.Padding = New Thickness(5, 0, 5, 0)
                btn.VerticalAlignment = VerticalAlignment.Stretch
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
                spPath.Children.Add(child)
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

    Private Async Sub ShowTree(container As clsDirectoryObject)
        If container Is Nothing Then Exit Sub

        Dim treepath As New List(Of clsDirectoryObject)

        Do ' find all parents to the root
            treepath.Add(container)

            If container.Entry.Parent Is Nothing OrElse container.Entry.Path = container.Domain.DefaultNamingContext.Path Then
                Exit Do
            Else
                container = New clsDirectoryObject(container.Entry.Parent, container.Domain)
            End If
        Loop

        treepath.Reverse()

        Dim currentparentnode = tviDomains
        For I = 0 To treepath.Count - 1
            For j = 0 To currentparentnode.Items.Count - 1
                If CType(currentparentnode.Items(j), clsDirectoryObject).Entry.Path <> treepath(I).Entry.Path Then Continue For

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

    Public Async Sub StopSearch()
        Await searcher.BasicSearchStopAsync()
    End Sub

    Public Sub StartSearch(Optional root As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        If root Is Nothing And (filter Is Nothing OrElse String.IsNullOrEmpty(filter.Filter)) Then Exit Sub

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

        If preferences.SearchResultGrouping AndAlso root Is Nothing Then
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

                obj.Entry.MoveTo(destination.Entry)
                obj.NotifyMoved()
            Next
        Catch ex As Exception
            ThrowException(ex, "cmdMove")
        End Try

        RefreshDataGrid()
        If organizationalunitaffected Then RefreshDomainTree()
    End Sub


#End Region

End Class
