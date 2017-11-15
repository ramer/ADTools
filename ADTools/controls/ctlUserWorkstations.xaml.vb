Imports System.Collections.ObjectModel
Imports IPrompt.VisualBasic

Public Class ctlUserWorkstations


    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlUserWorkstations),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentselectedobjects As New ObservableCollection(Of clsDirectoryObject)
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
        Dim instance As ctlUserWorkstations = CType(d, ctlUserWorkstations)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
            ._currentdomainobjects.Clear()

            .lvSelectedObjects.ItemsSource = If(._currentselectedobjects IsNot Nothing, ._currentselectedobjects, Nothing)
            .lvDomainObjects.ItemsSource = If(._currentobject IsNot Nothing, ._currentdomainobjects, Nothing)
        End With
    End Sub

    Private Sub ctlUserWorkstations_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDomainObjectsFilter.Focus()

        InitializeAsync()
    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing OrElse _currentobject.Entry Is Nothing OrElse _currentobject.userWorkstations.Count = 0 Then Exit Sub

        If _currentselectedobjects.Count = 0 Then
            cap.Visibility = Visibility.Visible

            Dim objects As New ObservableCollection(Of clsDirectoryObject)

            objects = Await Task.Run(
                Function()
                    Return searcher.BasicSearchSync(
                    New clsDirectoryObject(_currentobject.Domain.DefaultNamingContext, _currentobject.Domain),
                    New clsFilter("""" & Join(_currentobject.userWorkstations, """/""") & """", Nothing, New clsSearchObjectClasses(False, False, True, False, False)))
                End Function)

            For Each o In objects
                _currentselectedobjects.Add(o)
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
                New clsFilter(tbDomainObjectsFilter.Text, Nothing, New clsSearchObjectClasses(False, False, True, False, False)),
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

    Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvSelectedObjects.MouseMove,
                                                                            lvDomainObjects.MouseMove
        Dim listView As ListView = TryCast(sender, ListView)

        If e.LeftButton = MouseButtonState.Pressed And
            e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth And
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
                If Not (obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Computer) Then e.Effects = DragDropEffects.None : Exit For
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
                If Not (obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Computer) Then Exit Sub
            Next

            For Each obj In dropped
                If sender Is lvSelectedObjects Then ' adding member
                    If obj.Domain IsNot _currentobject.Domain Then IMsgBox(My.Resources.ctlGroupMember_msg_AnotherDomain, vbOKOnly + vbExclamation, My.Resources.ctlGroupMember_msg_AnotherDomainTitle) : Exit Sub
                    AddMember(obj)
                ElseIf sender Is trashSelectedObjects Then
                    RemoveMember(obj)
                End If
            Next
        End If
    End Sub

    Private Sub AddMember([object] As clsDirectoryObject)
        Try
            For Each obj As clsDirectoryObject In _currentselectedobjects
                If obj.name = [object].name Then Exit Sub
            Next
            _currentselectedobjects.Add([object])
            _currentobject.userWorkstations = _currentselectedobjects.Select(Function(x As clsDirectoryObject) x.name).ToArray
        Catch ex As Exception
            ThrowException(ex, "AddComputer")
        End Try
    End Sub

    Private Sub RemoveMember([object] As clsDirectoryObject)
        Try
            _currentselectedobjects.Remove([object])
            _currentobject.userWorkstations = _currentselectedobjects.Select(Function(x As clsDirectoryObject) x.name).ToArray
        Catch ex As Exception
            ThrowException(ex, "RemoveComputer")
        End Try
    End Sub

End Class
