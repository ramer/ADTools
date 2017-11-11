
Imports System.Collections.ObjectModel
Imports IPrompt.VisualBasic

Public Class ctlGroupMemberOf


    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlGroupMemberOf),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))
    Public Shared ReadOnly CurrentDomainProperty As DependencyProperty = DependencyProperty.Register("CurrentDomain",
                                                    GetType(clsDomain),
                                                    GetType(ctlGroupMemberOf),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf CurrentDomainPropertyChanged))

    Enum enmMode
        DomainDefaultGroups
        ObjectMemberOf
    End Enum

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomain As clsDomain
    Private Property _currentdomaingroups As New ObservableCollection(Of clsDirectoryObject)

    WithEvents searcher As New clsSearcher

    Private Mode As enmMode

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Public Property CurrentDomain() As clsDirectoryObject
        Get
            Return GetValue(CurrentDomainProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentDomainProperty, value)
        End Set
    End Property

    Sub New()
        InitializeComponent()
    End Sub

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlGroupMemberOf = CType(d, ctlGroupMemberOf)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
            ._currentdomain = CType(e.NewValue, clsDirectoryObject).Domain

            .Mode = enmMode.ObjectMemberOf

            ._currentdomaingroups.Clear()
            .lvSelectedGroups.Items.Clear()
            ._currentobject.memberOf.ToList.ForEach(Sub(x As clsDirectoryObject) .lvSelectedGroups.Items.Add(x))
            .lvDomainGroups.ItemsSource = If(._currentobject IsNot Nothing, ._currentdomaingroups, Nothing)
        End With
    End Sub

    Private Shared Sub CurrentDomainPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlGroupMemberOf = CType(d, ctlGroupMemberOf)
        With instance
            ._currentobject = Nothing
            ._currentdomain = CType(e.NewValue, clsDomain)

            .Mode = enmMode.DomainDefaultGroups

            ._currentdomaingroups.Clear()
            .lvSelectedGroups.Items.Clear()
            ._currentdomain.DefaultGroups.ToList.ForEach(Sub(x As clsDirectoryObject) .lvSelectedGroups.Items.Add(x))
            .lvDomainGroups.ItemsSource = If(._currentdomain IsNot Nothing, ._currentdomaingroups, Nothing)
        End With
    End Sub

    Private Sub ctlGroupMemberOf_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDomainGroupsFilter.Focus()
        btnDefaultGroups.Visibility = If(Mode = enmMode.ObjectMemberOf, Visibility.Visible, Visibility.Hidden)
    End Sub

    Public Async Sub InitializeAsync()
        If lvSelectedGroups.Items.Count = 0 Then
            cap.Visibility = Visibility.Visible

            Dim groups As New ObservableCollection(Of clsDirectoryObject)

            If _currentobject IsNot Nothing Then
                groups = Await Task.Run(Function() _currentobject.memberOf)
            Else
                groups = Await Task.Run(Function() _currentdomain.DefaultGroups)
            End If

            lvSelectedGroups.Items.Clear()
            For Each g In groups
                lvSelectedGroups.Items.Add(g)
            Next

            cap.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Async Sub tbDomainGroupsFilter_KeyDown(sender As Object, e As KeyEventArgs) Handles tbDomainGroupsFilter.KeyDown
        If e.Key = Key.Enter Then
            tbDomainGroupsFilter.SelectAll()
            Await searcher.BasicSearchAsync(
                _currentdomaingroups,
                Nothing,
                New clsFilter("*" & tbDomainGroupsFilter.Text & "*", Nothing, New clsSearchObjectClasses(False, False, False, True, False)),
                New ObservableCollection(Of clsDomain)({_currentdomain}))
        End If
    End Sub

    Private Sub lvDomainGroups_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvDomainGroups.MouseDoubleClick
        If lvDomainGroups.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvDomainGroups.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lvSelectedGroups_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lvSelectedGroups.MouseDoubleClick
        If lvSelectedGroups.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvSelectedGroups.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvSelectedGroups.MouseMove,
                                                                            lvDomainGroups.MouseMove
        Dim listView As ListView = TryCast(sender, ListView)

        If e.LeftButton = MouseButtonState.Pressed And
            e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth And
            listView.SelectedItems.Count > 0 Then

            Dim dragData As New DataObject(listView.SelectedItems.Cast(Of clsDirectoryObject).ToArray)

            DragDrop.DoDragDrop(listView, dragData, DragDropEffects.All)
        End If
    End Sub

    Private Sub lvSelectedGroups_DragEnter(sender As Object, e As DragEventArgs) Handles lvSelectedGroups.DragEnter,
                                                                                         trashSelectedGroups.DragEnter,
                                                                                         lvSelectedGroups.DragOver,
                                                                                         trashSelectedGroups.DragOver
        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            e.Effects = DragDropEffects.Copy
            For Each obj As clsDirectoryObject In e.Data.GetData(GetType(clsDirectoryObject()))
                If obj.SchemaClass <> clsDirectoryObject.enmSchemaClass.Group Then e.Effects = DragDropEffects.None : Exit For
            Next
        Else
            e.Effects = DragDropEffects.None
        End If

        If e.Effects = DragDropEffects.Copy Then
            If sender Is lvSelectedGroups Then trashSelectedGroups.Visibility = Visibility.Visible
            If sender Is trashSelectedGroups Then trashSelectedGroups.Visibility = Visibility.Visible : trashSelectedGroups.Background = Application.Current.Resources("ColorButtonBackground")
        End If

        e.Handled = True
    End Sub

    Private Sub lvSelectedGroups_DragLeave(sender As Object, e As DragEventArgs) Handles lvSelectedGroups.DragLeave,
                                                                                            trashSelectedGroups.DragLeave
        If sender Is lvSelectedGroups Then trashSelectedGroups.Visibility = Visibility.Collapsed
        If sender Is trashSelectedGroups Then trashSelectedGroups.Visibility = Visibility.Collapsed : trashSelectedGroups.Background = Brushes.Transparent
    End Sub

    Private Sub lvSelectedGroups_Drop(sender As Object, e As DragEventArgs) Handles lvSelectedGroups.Drop,
                                                                                    trashSelectedGroups.Drop
        trashSelectedGroups.Visibility = Visibility.Collapsed : trashSelectedGroups.Background = Brushes.Transparent

        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            Dim dropped = e.Data.GetData(GetType(clsDirectoryObject()))
            For Each obj As clsDirectoryObject In dropped
                If obj.SchemaClass <> clsDirectoryObject.enmSchemaClass.Group Then Exit Sub
            Next

            For Each obj In dropped
                If sender Is lvSelectedGroups Then ' adding member
                    If Mode = enmMode.ObjectMemberOf AndAlso obj.Domain IsNot _currentobject.Domain Then IMsgBox(My.Resources.ctlGroupMember_msg_AnotherDomain, vbOKOnly + vbExclamation, My.Resources.ctlGroupMember_msg_AnotherDomainTitle) : Exit Sub
                    AddMember(obj)
                ElseIf sender Is trashSelectedGroups Then
                    RemoveMember(obj)
                End If
            Next
        End If
    End Sub

    Private Sub AddMember([object] As clsDirectoryObject)
        Try

            If Mode = enmMode.ObjectMemberOf Then

                For Each group As clsDirectoryObject In _currentobject.memberOf
                    If group.name = [object].name Then Exit Sub
                Next

                [object].Entry.Invoke("Add", _currentobject.Entry.Path)
                [object].Entry.CommitChanges()
                _currentobject.memberOf.Add([object])
                lvSelectedGroups.Items.Add([object])

            ElseIf Mode = enmMode.DomainDefaultGroups Then

                For Each group As clsDirectoryObject In _currentdomain.DefaultGroups
                    If group.name = [object].name Then Exit Sub
                Next

                _currentdomain.DefaultGroups.Add([object])
                _currentdomain.DefaultGroups = _currentdomain.DefaultGroups ' update setter
                lvSelectedGroups.Items.Add([object])

            End If

        Catch ex As Exception
            ThrowException(ex, "AddMember")
            If _currentobject IsNot Nothing And [object] IsNot Nothing AndAlso _currentobject.objectClass.Contains("group") And [object].objectClass.Contains("group") Then ShowWrongMemberMessage()
        End Try
    End Sub

    Private Sub RemoveMember([object] As clsDirectoryObject)
        Try
            If Mode = enmMode.ObjectMemberOf Then

                If Not _currentobject.memberOf.Contains([object]) Then Exit Sub
                [object].Entry.Invoke("Remove", _currentobject.Entry.Path)
                [object].Entry.CommitChanges()
                _currentobject.memberOf.Remove([object])
                lvSelectedGroups.Items.Remove([object])

            ElseIf Mode = enmMode.DomainDefaultGroups Then

                If Not _currentdomain.DefaultGroups.Contains([object]) Then Exit Sub
                _currentdomain.DefaultGroups.Remove([object])
                _currentdomain.DefaultGroups = _currentdomain.DefaultGroups ' update setter
                lvSelectedGroups.Items.Remove([object])

            End If

        Catch ex As Exception
            ThrowException(ex, "RemoveMember")
        End Try
    End Sub

    Private Sub btnDefaultGroups_Click(sender As Object, e As RoutedEventArgs) Handles btnDefaultGroups.Click
        If Mode = enmMode.DomainDefaultGroups Then Exit Sub

        For Each group In _currentobject.memberOf
            group.Entry.Invoke("Remove", _currentobject.Entry.Path)
            group.Entry.CommitChanges()
        Next
        _currentobject.memberOf.Clear()
        lvSelectedGroups.Items.Clear()
        For Each group In _currentobject.Domain.DefaultGroups
            group.Entry.Invoke("Add", _currentobject.Entry.Path)
            group.Entry.CommitChanges()
            _currentobject.memberOf.Add(group)
            lvSelectedGroups.Items.Add(group)
        Next
    End Sub

End Class
