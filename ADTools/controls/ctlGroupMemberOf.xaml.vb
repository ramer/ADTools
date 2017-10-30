
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

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomain As clsDomain
    Private Property _currentdomaingroups As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    WithEvents searcher As New clsSearcher

    Private sourceobject As Object
    Private dragallow As Boolean
    Private dragobjects As New List(Of clsDirectoryObject)

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
            ._currentdomaingroups.Clear()
            .lvSelectedGroups.Items.Clear()
            ._currentdomain.DefaultGroups.ToList.ForEach(Sub(x As clsDirectoryObject) .lvSelectedGroups.Items.Add(x))
            .lvDomainGroups.ItemsSource = If(._currentdomain IsNot Nothing, ._currentdomaingroups, Nothing)
        End With
    End Sub

    Private Sub ctlGroupMemberOf_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDomainGroupsFilter.Focus()
        btnDefaultGroups.Visibility = If(mode() = 0, Visibility.Visible, Visibility.Hidden)
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

    Private Function mode() As Integer
        If _currentobject IsNot Nothing Then
            Return 0
        ElseIf _currentdomain IsNot Nothing Then
            Return 1
        Else
            Return -1
        End If
    End Function

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

    Private Sub lv_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles lvSelectedGroups.PreviewMouseLeftButtonDown,
                                                                                                   lvDomainGroups.PreviewMouseLeftButtonDown
        Dim listView As ListView = TryCast(sender, ListView)
        dragallow = e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth 'And e.GetPosition(sender).Y < listView.ActualHeight - SystemParameters.HorizontalScrollBarHeight
        dragobjects = listView.SelectedItems.Cast(Of clsDirectoryObject).ToList
    End Sub

    Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvSelectedGroups.MouseMove,
                                                                            lvDomainGroups.MouseMove
        Dim listView As ListView = TryCast(sender, ListView)

        If dragobjects.Count = 0 And listView.SelectedItem IsNot Nothing Then dragobjects.Add(listView.SelectedItem)

        sourceobject = Nothing

        If e.LeftButton = MouseButtonState.Pressed And dragobjects.Count > 0 And dragallow Then
            sourceobject = listView

            Dim dragData As New DataObject("dolist", dragobjects)

            DragDrop.DoDragDrop(listView, dragData, DragDropEffects.Copy)
        End If
    End Sub

    Private Sub lv_DragEnter(sender As Object, e As DragEventArgs) Handles lvSelectedGroups.DragEnter,
                                                                            lvDomainGroups.DragEnter

        If Not e.Data.GetDataPresent("dolist") OrElse sender Is sourceobject Then
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub lv_Drop(sender As Object, e As DragEventArgs) Handles lvSelectedGroups.Drop,
                                                                        lvDomainGroups.Drop

        If e.Data.GetDataPresent("dolist") And sender IsNot sourceobject Then

            Dim dragged As List(Of clsDirectoryObject) = TryCast(e.Data.GetData("dolist"), List(Of clsDirectoryObject))

            For Each obj In dragged
                If sender Is lvSelectedGroups Then ' adding member
                    If mode() = 0 AndAlso obj.Domain IsNot _currentobject.Domain Then IMsgBox(My.Resources.ctlGroupMember_msg_AnotherDomain, vbOKOnly + vbExclamation, My.Resources.ctlGroupMember_msg_AnotherDomainTitle) : Exit Sub
                    AddMember(obj)
                Else
                    RemoveMember(obj)
                End If
            Next
        End If
    End Sub

    Private Sub AddMember([object] As clsDirectoryObject)
        Try

            If mode() = 0 Then

                For Each group As clsDirectoryObject In _currentobject.memberOf
                    If group.name = [object].name Then Exit Sub
                Next

                [object].Entry.Invoke("Add", _currentobject.Entry.Path)
                [object].Entry.CommitChanges()
                _currentobject.memberOf.Add([object])
                lvSelectedGroups.Items.Add([object])

            ElseIf mode() = 1 Then

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

            If mode() = 0 Then

                If Not _currentobject.memberOf.Contains([object]) Then Exit Sub
                [object].Entry.Invoke("Remove", _currentobject.Entry.Path)
                [object].Entry.CommitChanges()
                _currentobject.memberOf.Remove([object])
                lvSelectedGroups.Items.Remove([object])

            ElseIf mode() = 1 Then

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
        If mode() <> 0 Then Exit Sub

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
