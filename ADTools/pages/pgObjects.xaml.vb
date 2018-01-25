Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Controls.Primitives
Imports IPrompt.VisualBasic

Class pgObjects
    Public Shared hkF5 As New RoutedCommand
    Public Shared hkEsc As New RoutedCommand

    Public WithEvents clipboardTimer As New Threading.DispatcherTimer()
    Private clipboardlastdata As String

    Public WithEvents searcher As New clsSearcher

    Public Property currentcontainer As clsDirectoryObject
    Public Property currentfilter As clsFilter
    Public Property currentobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Public Property cvscurrentobjects As New CollectionViewSource
    Public Property cvcurrentobjects As ICollectionView

    Sub New()
        InitializeComponent()
        InitializeObject()
    End Sub

    Sub New(Optional root As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        InitializeComponent()
        InitializeObject()
        Search(root, filter)
    End Sub

    Private Sub InitializeObject()
        RefreshDomainTree()

        DataObject.AddPastingHandler(tbSearchPattern, AddressOf tbSearchPattern_OnPaste)
        tbSearchPattern.Focus()

        tviFavorites.ItemsSource = preferences.Favorites
        tviFilters.ItemsSource = preferences.Filters

        gcdNavigation.DataContext = preferences
        gcdPreview.DataContext = preferences

        hkEsc.InputGestures.Add(New KeyGesture(Key.Escape))
        Me.CommandBindings.Add(New CommandBinding(hkEsc, Sub() StopSearch()))
        hkF5.InputGestures.Add(New KeyGesture(Key.F5))
        Me.CommandBindings.Add(New CommandBinding(hkF5, Sub() Search(currentcontainer, currentfilter)))

        clipboardTimer.Interval = New TimeSpan(0, 0, 1)
        clipboardTimer.Start()

        cvscurrentobjects = New CollectionViewSource() With {.Source = currentobjects}
        cvcurrentobjects = cvscurrentobjects.View
        cvcurrentobjects.GroupDescriptions.Add(New PropertyGroupDescription("Domain.Name"))

        lvObjects.SetBinding(ItemsControl.ItemsSourceProperty, New Binding("cvcurrentobjects") With {.IsAsync = True})
        lvObjects.ViewStyleDetails = GetViewDetailsStyle()
        ' mega-crutch - lvObjects binded to local cvcurrentobjects property and remote preferences properties
        AddHandler preferences.PropertyChanged, Sub(s, a) If a.PropertyName = "ViewResultGrouping" Then lvObjects.EnableGrouping = preferences.ViewResultGrouping
        AddHandler preferences.PropertyChanged, Sub(s, a) If a.PropertyName = "CurrentView" Then lvObjects.CurrentView = preferences.CurrentView
        lvObjects.CurrentView = preferences.CurrentView
        lvObjects.EnableGrouping = preferences.ViewResultGrouping
    End Sub

#Region "ContextMenu"

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

    Private Sub ctxmnuSharedCreateObject_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(CType(sender, MenuItem).Parent, ContextMenu).Tag IsNot clsDirectoryObject() Then Exit Sub
        Dim objects() As clsDirectoryObject = CType(CType(sender, MenuItem).Parent, ContextMenu).Tag

        If objects.Count = 1 AndAlso (
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Container Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.DomainDNS Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.UnknownContainer) Then

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
            Dim name As String = IInputBox(My.Resources.str_EnterObjectName, My.Resources.str_RenameObject, objects(0).name, vbQuestion, Window.GetWindow(Me))
            If Len(name) > 0 Then
                If obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Then
                    obj.Rename("OU=" & name)
                    organizationalunitaffected = True
                Else
                    obj.Rename("CN=" & name)
                End If

                'RefreshSearchResults()
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

            If IMsgBox(My.Resources.str_AreYouSure, vbYesNo + vbQuestion, My.Resources.str_RemoveObject, Window.GetWindow(Me)) <> MsgBoxResult.Yes Then Exit Sub
            If objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Then
                If IMsgBox(My.Resources.str_ThisIsOrganizationaUnit & vbCrLf & vbCrLf & My.Resources.str_AreYouSure, vbYesNo + vbExclamation, My.Resources.str_RemoveObject, Window.GetWindow(Me)) <> MsgBoxResult.Yes Then Exit Sub
                If currentcontainer IsNot Nothing AndAlso objects(0).distinguishedName = currentcontainer.distinguishedName Then currentcontaineraffected = True
            End If

            objects(0).DeleteTree()

            If currentcontaineraffected Then
                OpenObjectParent()
                RefreshDomainTree()
            Else
                'RefreshSearchResults()
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
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User Or
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.Computer)) Then Exit Sub

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
            objects(0).SchemaClass = clsDirectoryObject.enmSchemaClass.User)) Then Exit Sub

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
                fdmnuclear.Header = My.Resources.ctxmnu_PostfilterClear
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
            cdmnudn.Header = My.Resources.ctxmnu_CopyDataDisplayName
            cdmnudn.Tag = "displayName"
            AddHandler cdmnudn.Click, AddressOf ctxmnuObjectsCopyDataItem_Click
            ctxmnuObjectsCopyData.Items.Add(cdmnudn)

            Dim cdmnuba As New MenuItem
            cdmnuba.Header = My.Resources.ctxmnu_CopyDataBasicAttributes
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

    Private Sub ctxmnutviDomainsDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnutviDomainsDomainOptions.Click
        ShowPage(New pgDomains, True, Window.GetWindow(Me), True)
        RefreshDomainTree()
    End Sub

    Private Sub ctxmnuObjectsRestore_Click(sender As Object, e As RoutedEventArgs)
        'TODO
        IMsgBox("Не допилено :/", vbOKOnly, "ADTools", Window.GetWindow(Me))
    End Sub

    Private Sub ctxmnuObjectsRestoreToContainer_Click(sender As Object, e As RoutedEventArgs)
        'TODO
        IMsgBox("Не допилено :/", vbOKOnly, "ADTools", Window.GetWindow(Me))
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

#End Region

#Region "Subs"

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

    Public Sub SaveCurrentFilter()
        If currentfilter Is Nothing OrElse String.IsNullOrEmpty(currentfilter.Filter) Then IMsgBox(My.Resources.str_CannotSaveCurrentFilter, vbOKOnly + vbExclamation,, Window.GetWindow(Me)) : Exit Sub

        Dim name As String = IInputBox(My.Resources.str_EnterFilterName,,, vbQuestion, Window.GetWindow(Me))

        If String.IsNullOrEmpty(name) Then Exit Sub

        currentfilter.Name = name
        preferences.Filters.Add(currentfilter)
    End Sub

    Public Sub RefreshDomainTree()
        tviDomains.ItemsSource = domains.Where(Function(d As clsDomain) d.Validated).Select(Function(d As clsDomain) New clsDirectoryObject(d.DefaultNamingContext, d))
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

        btnUp.IsEnabled = container IsNot Nothing AndAlso (container.Parent IsNot Nothing And Not container.distinguishedName = container.Domain.DefaultNamingContext)

        If container IsNot Nothing Then

            Dim buttons As New List(Of Button)

            Dim containerDN As String = container.distinguishedName

            Title = GetNameFromDN(containerDN)

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
            tblck.Text = My.Resources.str_SearchResults & " " & If(Not String.IsNullOrEmpty(filter.Name), filter.Name, If(Not String.IsNullOrEmpty(filter.Pattern), filter.Pattern.Replace(vbNewLine, " "), " Advanced filter"))
            tblck.Margin = New Thickness(2, 0, 2, 0)
            tblck.Padding = New Thickness(5, 0, 5, 0)
            spPath.Children.Add(tblck)
            Title = tblck.Text

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
                container.SetValue(ItemsControl.FontWeightProperty, FontWeights.Bold)
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
            If sourceobjects.Count > 1 Then IMsgBox(My.Resources.str_CopyOneObjectWithinDomain, vbOKOnly + vbExclamation, My.Resources.str_CopyObject, Window.GetWindow(Me)) : Exit Sub

            ' single object

            Dim w As New pgCreateObject(destination, destination.Domain)

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
                    IMsgBox(My.Resources.str_UnknownObjectClass, vbOKOnly + vbExclamation, My.Resources.str_CopyObject, Window.GetWindow(Me))
            End Select
            ShowPage(w, False, Window.GetWindow(Me), False)

        Else ' another domain
            ' TODO
            IMsgBox("Не допилено :/", vbOKOnly, "ADTools", Window.GetWindow(Me))

        End If
    End Sub

    Public Sub ObjectsMove(destination As clsDirectoryObject, sourceobjects() As clsDirectoryObject)
        Dim organizationalunitaffected As Boolean = False

        ' another domain
        If sourceobjects(0).Domain IsNot destination.Domain Then IMsgBox(My.Resources.str_MovingIsProhibited, vbOKOnly + vbExclamation, My.Resources.str_MoveObject, Window.GetWindow(Me)) : Exit Sub

        ' same domain
        If IMsgBox(My.Resources.str_AreYouSure & vbCrLf & vbCrLf &
                   String.Format(My.Resources.str_MoveObjectToContainer,
                                 If(sourceobjects.Count > 1, String.Format(My.Resources.str_nObjects, sourceobjects.Count), sourceobjects(0).name),
                                 destination.distinguishedNameFormated), vbYesNo + vbQuestion, My.Resources.str_MoveObject, Window.GetWindow(Me)) <> MessageBoxResult.Yes Then Exit Sub

        Try
            For Each obj In sourceobjects
                If obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Then organizationalunitaffected = True

                obj.MoveTo(destination.distinguishedName)
            Next
        Catch ex As Exception
            ThrowException(ex, "ObjectsMove")
        End Try

        'RefreshSearchResults()
        If organizationalunitaffected Then RefreshDomainTree()
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
            StartSearch(Nothing, New clsFilter(clipboardlastdata, preferences.AttributesForSearch, preferences.SearchObjectClasses))
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
        Dim obj = FindVisualParent(Of ListViewItem)(e.OriginalSource)
        If obj Is Nothing OrElse TypeOf obj.DataContext IsNot clsDirectoryObject Then Exit Sub
        OpenObject(obj.DataContext)
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
            If preferences.SearchMode = enmSearchMode.Default Then
                StartSearch(Nothing, New clsFilter(tbSearchPattern.Text, preferences.AttributesForSearch, preferences.SearchObjectClasses))
            Else
                StartSearch(Nothing, New clsFilter(tbSearchPattern.Text))
            End If
        End If
    End Sub

    Private Sub tbSearchPattern_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles tbSearchPattern.PreviewMouseDown
        If e.ClickCount = 3 Then tbSearchPattern.SelectAll()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If preferences.SearchMode = enmSearchMode.Default Then
            StartSearch(Nothing, New clsFilter(tbSearchPattern.Text, preferences.AttributesForSearch, preferences.SearchObjectClasses))
        Else
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

#End Region

#Region "Search"

    Public Sub StopSearch()
        searcher.StopAllSearchAsync()
    End Sub

    Public Sub StartSearch(Optional root As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        If root Is Nothing And (filter Is Nothing OrElse String.IsNullOrEmpty(filter.Filter)) Then Exit Sub
        NavigationService.Navigate(New pgObjects(root, filter))
    End Sub

    Public Sub Search(Optional root As clsDirectoryObject = Nothing, Optional filter As clsFilter = Nothing)
        If root Is Nothing And filter Is Nothing Then Exit Sub

        If filter IsNot Nothing AndAlso Not String.IsNullOrEmpty(filter.Pattern) Then tbSearchPattern.Text = filter.Pattern
        If root IsNot Nothing Then ShowInTree(root)

        tbSearchPattern.SelectAll()
        ApplyPostfilter(Nothing, Nothing)

        cap.Visibility = Visibility.Visible
        pbSearch.Visibility = Visibility.Visible

        Dim domainlist As New ObservableCollection(Of clsDomain)(domains.Where(Function(x As clsDomain) x.IsSearchable = True).ToList)

        currentcontainer = root
        currentfilter = filter
        ShowPath(currentcontainer, filter)

        searcher.SearchAsync(currentobjects, root, filter, domainlist, attributesToLoadDefault, preferences.ViewShowDeletedObjects)
    End Sub

    Private Sub Searcher_SearchAsyncDataRecieved() Handles searcher.SearchAsyncDataRecieved
        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub Searcher_SearchAsyncCompleted() Handles searcher.SearchAsyncCompleted
        cap.Visibility = Visibility.Hidden
        pbSearch.Visibility = Visibility.Hidden
    End Sub

#End Region

End Class
