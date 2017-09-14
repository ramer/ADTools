Imports System.Collections.ObjectModel
Imports System.DirectoryServices

Class wndMain

    Public Property objects As New ObservableCollection(Of clsDirectoryObject)

    Private Sub wndMain_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DomainTreeUpdate()
    End Sub

    Private Sub wndMain_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles MyBase.Closing, MyBase.Closing
        Windows.Application.Current.Shutdown()
    End Sub

#Region "Main Menu"

    Private Sub mnuServiceDomainOptions_Click(sender As Object, e As RoutedEventArgs) Handles mnuServiceDomainOptions.Click
        ShowWindow(New wndDomains, True, Me, True)
    End Sub



#End Region


#Region "Events"

    Private Sub tvDomains_TreeViewItem_Selected(sender As Object, e As RoutedEventArgs)
        Dim trvitem As TreeViewItem = TryCast(sender, TreeViewItem)
        If trvitem Is e.OriginalSource AndAlso TypeOf CType(trvitem, TreeViewItem).DataContext Is clsDirectoryObject Then
            Dim parent As clsDirectoryObject = CType(CType(trvitem, TreeViewItem).DataContext, clsDirectoryObject)
            DomainTreeOpenNode(parent)
            trvitem.IsSelected = False
            'dgObjects.UpdateLayout()
        End If
    End Sub

    Private Sub dgObjects_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles dgObjects.PreviewKeyDown
        Select Case e.Key
            Case Key.Enter
                If dgObjects.SelectedItem Is Nothing Then Exit Sub
                Dim current As clsDirectoryObject = CType(dgObjects.SelectedItem, clsDirectoryObject)
                DomainTreeOpenNode(current)
                e.Handled = True
            Case Key.Back
                DomainTreeOpenParentNode()
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
        DomainTreeOpenNode(current)
    End Sub

    Private Sub btnUp_Click(sender As Object, e As RoutedEventArgs) Handles btnUp.Click
        DomainTreeOpenParentNode()
    End Sub

    Private Sub btnPath_Click(sender As Object, e As RoutedEventArgs)
        If sender.Tag Is Nothing Then Exit Sub
        Dim current As clsDirectoryObject = CType(sender.Tag, clsDirectoryObject)
        DomainTreeOpenNode(current)
    End Sub

#End Region

#Region "Subs"

    Public Sub DomainTreeUpdate()
        tvDomains.ItemsSource = domains.Where(Function(d As clsDomain) d.Validated).Select(Function(d) If(d IsNot Nothing, New clsDirectoryObject(d.DefaultNamingContext, d), Nothing))
    End Sub

    Public Sub DomainTreeOpenNode(current As clsDirectoryObject)
        If current.SchemaClassName = "computer" Or
           current.SchemaClassName = "group" Or
           current.SchemaClassName = "contact" Or
           current.SchemaClassName = "user" Then

            'ShowDirectoryObjectProperties(current, Window.GetWindow(Me))
        ElseIf current.SchemaClassName = "organizationalUnit" Or
               current.SchemaClassName = "container" Or
               current.SchemaClassName = "builtinDomain" Or
               current.SchemaClassName = "domainDNS" Then

            ShowChildren(current)
            dgObjects.Tag = current
        End If

        ShowPath()
    End Sub

    Public Sub DomainTreeOpenParentNode()
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
            btn.Content = If(current.SchemaClassName = "domainDNS", current.Domain.Name, current.name)
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

#End Region

End Class
