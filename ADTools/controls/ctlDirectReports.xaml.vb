﻿Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports IPrompt.VisualBasic

Public Class ctlDirectReports


    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlDirectReports),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomainobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    WithEvents searcher As New clsSearcher

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
        Dim instance As ctlDirectReports = CType(d, ctlDirectReports)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
            ._currentdomainobjects.Clear()
            .lvDomainObjects.ItemsSource = If(._currentobject IsNot Nothing, ._currentdomainobjects, Nothing)

            If .lvDomainObjects.ItemsSource IsNot Nothing Then
                Dim view As CollectionView = CollectionViewSource.GetDefaultView(.lvDomainObjects.ItemsSource)
                view.SortDescriptions.Clear()
                view.SortDescriptions.Add(New SortDescription("name", ListSortDirection.Ascending))
            End If
        End With
    End Sub

    Private Sub ctlDirectReports_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDomainObjectsFilter.Focus()

        InitializeAsync()
    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing Then Exit Sub

        If lvSelectedObjects.Items.Count = 0 Then
            cap.Visibility = Visibility.Visible

            Dim objects As New ObservableCollection(Of clsDirectoryObject)

            objects = Await Task.Run(Function() _currentobject.directReports)

            lvSelectedObjects.Items.Clear()
            For Each o In objects
                lvSelectedObjects.Items.Add(o)
            Next
            lvSelectedObjects.Items.SortDescriptions.Clear()
            lvSelectedObjects.Items.SortDescriptions.Add(New SortDescription("name", ListSortDirection.Ascending))

            cap.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Sub tbDomainObjectsFilter_KeyDown(sender As Object, e As KeyEventArgs) Handles tbDomainObjectsFilter.KeyDown
        If e.Key = Key.Enter Then
            tbDomainObjectsFilter.SelectAll()
            searcher.SearchAsync(
                _currentdomainobjects,
                Nothing,
                New clsFilter(tbDomainObjectsFilter.Text, Nothing, New clsSearchObjectClasses(True, False, False, False, False)),
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

    'Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvSelectedObjects.MouseMove,
    '                                                                        lvDomainObjects.MouseMove
    '    Dim listView As ListView = TryCast(sender, ListView)

    '    If e.LeftButton = MouseButtonState.Pressed And
    '        e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth And
    '        listView.SelectedItems.Count > 0 Then

    '        Dim dragData As New DataObject(listView.SelectedItems.Cast(Of clsDirectoryObject).ToArray)

    '        DragDrop.DoDragDrop(listView, dragData, DragDropEffects.All)
    '    End If
    'End Sub

    Private Sub lv_TouchMove(sender As Object, e As TouchEventArgs) Handles lvSelectedObjects.TouchMove,
                                                                            lvDomainObjects.TouchMove
        Dim listView As ListView = TryCast(sender, ListView)

        If e.GetTouchPoint(sender).Position.X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth And
            listView.SelectedItems.Count > 0 Then

            Dim dragData As New DataObject(listView.SelectedItems.Cast(Of clsDirectoryObject).ToArray)

            DragDrop.DoDragDrop(listView, dragData, DragDropEffects.All)
        End If
    End Sub

    Private Sub lvSelectedObjects_DragEnter(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.DragEnter,
                                                                            trashSelectedObjects.DragEnter,
                                                                            lvSelectedObjects.DragOver,
                                                                            trashSelectedObjects.DragOver
        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            e.Effects = DragDropEffects.Copy
            For Each obj As clsDirectoryObject In e.Data.GetData(GetType(clsDirectoryObject()))
                If Not (obj.SchemaClass = enmDirectoryObjectSchemaClass.User) Then e.Effects = DragDropEffects.None : Exit For
            Next
        Else
            e.Effects = DragDropEffects.None
        End If

        If e.Effects = DragDropEffects.Copy Then
            If sender Is lvSelectedObjects Then trashSelectedObjects.Visibility = Visibility.Visible
            If sender Is trashSelectedObjects Then trashSelectedObjects.Visibility = Visibility.Visible : trashSelectedObjects.Background = Application.Current.Resources("ColorButtonBackground")
        End If

        e.Handled = True
    End Sub

    Private Sub lvSelectedObjects_DragLeave(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.DragLeave,
                                                                                            trashSelectedObjects.DragLeave
        If sender Is lvSelectedObjects Then trashSelectedObjects.Visibility = Visibility.Collapsed
        If sender Is trashSelectedObjects Then trashSelectedObjects.Visibility = Visibility.Collapsed : trashSelectedObjects.Background = Brushes.Transparent
    End Sub

    Private Sub lv_Drop(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.Drop,
                                                                      trashSelectedObjects.Drop
        trashSelectedObjects.Visibility = Visibility.Collapsed : trashSelectedObjects.Background = Brushes.Transparent

        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            Dim dropped = e.Data.GetData(GetType(clsDirectoryObject()))
            For Each obj As clsDirectoryObject In dropped
                If Not (obj.SchemaClass = enmDirectoryObjectSchemaClass.User) Then Exit Sub
            Next

            For Each obj In dropped
                If sender Is lvSelectedObjects Then ' adding member
                    If obj.Domain IsNot _currentobject.Domain Then IMsgBox(My.Resources.str_AnotherDomain, vbOKOnly + vbExclamation, My.Resources.str_AnotherDomainTitle) : Exit Sub
                    AddMember(obj)
                ElseIf sender Is trashSelectedObjects Then
                    RemoveMember(obj)
                End If
            Next
        End If
    End Sub

    Private Sub AddMember([object] As clsDirectoryObject)
        Try

            For Each obj As clsDirectoryObject In _currentobject.directReports
                If obj.name = [object].name Then Exit Sub
            Next

            [object].manager = _currentobject
            _currentobject.directReports.Add([object])
            lvSelectedObjects.Items.Add([object])

        Catch ex As Exception
            ThrowException(ex, "AddMember")
        End Try
    End Sub

    Private Sub RemoveMember([object] As clsDirectoryObject)
        Try

            If Not _currentobject.directReports.Contains([object]) Then Exit Sub
            [object].manager = Nothing
            _currentobject.directReports.Remove([object])
            lvSelectedObjects.Items.Remove([object])

        Catch ex As Exception
            ThrowException(ex, "RemoveMember")
        End Try
    End Sub

End Class



