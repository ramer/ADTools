
Imports System.Collections.ObjectModel
Imports IPrompt.VisualBasic

Public Class ctlGroupMember


    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlGroupMember),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomainobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

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

    Sub New()
        InitializeComponent()
    End Sub

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlGroupMember = CType(d, ctlGroupMember)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
            ._currentdomainobjects.Clear()
            '.lvSelectedObjects.Items.Clear()
            'CType(._currentobject.member, ObservableCollection(Of clsDirectoryObject)).ToList.ForEach(Sub(x As clsDirectoryObject) .lvSelectedObjects.Items.Add(x))
            .lvDomainObjects.ItemsSource = If(._currentobject IsNot Nothing, ._currentdomainobjects, Nothing)
        End With
    End Sub

    Private Sub ctlGroupMember_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDomainObjectsFilter.Focus()
    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing Then Exit Sub

        If lvSelectedObjects.Items.Count = 0 Then
            cap.Visibility = Visibility.Visible

            Dim groups As New ObservableCollection(Of clsDirectoryObject)

            groups = Await Task.Run(Function() _currentobject.member)

            lvSelectedObjects.Items.Clear()
            For Each g In groups
                lvSelectedObjects.Items.Add(g)
            Next

            cap.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Async Sub tbDomainObjectsFilter_KeyDown(sender As Object, e As KeyEventArgs) Handles tbDomainObjectsFilter.KeyDown
        If e.Key = Key.Enter Then
            tbDomainObjectsFilter.SelectAll()
            Await searcher.BasicSearchAsync(
                _currentdomainobjects,
                Nothing,
                New clsFilter(tbDomainObjectsFilter.Text, Nothing, New clsSearchObjectClasses(True, False, True, True, False), False),
                New ObservableCollection(Of clsDomain)({_currentobject.Domain}))
        End If
    End Sub

    Private Sub lvDomainObjects_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvDomainObjects.MouseDoubleClick
        If lvDomainObjects.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvDomainObjects.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lvSelectedObjects_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lvSelectedObjects.MouseDoubleClick
        If lvSelectedObjects.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvSelectedObjects.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lv_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles lvSelectedObjects.PreviewMouseLeftButtonDown,
                                                                                                   lvDomainObjects.PreviewMouseLeftButtonDown
        Dim listView As ListView = TryCast(sender, ListView)
        sourceobject = Nothing
        dragallow = e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth 'And e.GetPosition(sender).Y < listView.ActualHeight - SystemParameters.HorizontalScrollBarHeight
        dragobjects = listView.SelectedItems.Cast(Of clsDirectoryObject).ToList
    End Sub

    Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvSelectedObjects.MouseMove,
                                                                            lvDomainObjects.MouseMove
        Dim listView As ListView = TryCast(sender, ListView)

        If dragobjects.Count = 0 And listView.SelectedItem IsNot Nothing Then dragobjects.Add(listView.SelectedItem)

        sourceobject = Nothing

        If e.LeftButton = MouseButtonState.Pressed And dragobjects.Count > 0 And dragallow Then
            sourceobject = listView

            Dim dragData As New DataObject("dolist", dragobjects)

            DragDrop.DoDragDrop(listView, dragData, DragDropEffects.Copy)
        End If
    End Sub

    Private Sub lv_DragEnter(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.DragEnter,
                                                                            lvDomainObjects.DragEnter

        If Not e.Data.GetDataPresent("dolist") OrElse sender Is sourceobject Then
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub lv_Drop(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.Drop,
                                                                        lvDomainObjects.Drop

        If e.Data.GetDataPresent("dolist") And sender IsNot sourceobject Then

            Dim dragged As List(Of clsDirectoryObject) = TryCast(e.Data.GetData("dolist"), List(Of clsDirectoryObject))

            For Each obj In dragged
                If sender Is lvSelectedObjects Then ' adding member
                    If obj.Domain IsNot _currentobject.Domain Then IMsgBox("Из другого домена нельзя же!", vbOKOnly + vbExclamation, "Ну тычоваще :(") : Exit Sub
                    AddMember(obj)
                Else
                    RemoveMember(obj)
                End If
            Next
        End If
    End Sub

    Private Sub AddMember([object] As clsDirectoryObject)
        Try

            For Each obj As clsDirectoryObject In _currentobject.member
                If obj.name = [object].name Then Exit Sub
            Next

            _currentobject.Entry.Invoke("Add", [object].Entry.Path)
            _currentobject.Entry.CommitChanges()
            _currentobject.member.Add([object])
            lvSelectedObjects.Items.Add([object])

        Catch ex As Exception
            ThrowException(ex, "AddMember")
            If _currentobject IsNot Nothing And [object] IsNot Nothing AndAlso _currentobject.objectClass.Contains("group") And [object].objectClass.Contains("group") Then ShowWrongMemberMessage()
        End Try
    End Sub

    Private Sub RemoveMember([object] As clsDirectoryObject)
        Try

            If Not _currentobject.member.Contains([object]) Then Exit Sub
            _currentobject.Entry.Invoke("Remove", [object].Entry.Path)
            _currentobject.Entry.CommitChanges()
            _currentobject.member.Remove([object])
            lvSelectedObjects.Items.Remove([object])

        Catch ex As Exception
            ThrowException(ex, "RemoveMember")
        End Try
    End Sub

End Class
