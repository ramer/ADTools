
Imports System.Collections.ObjectModel

Public Class ctlUserWorkstations


    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlUserWorkstations),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentselectedobjects As New ObservableCollection(Of clsDirectoryObject)
    Private Property _currentdomainobjects As New ObservableCollection(Of clsDirectoryObject)

    WithEvents searcher As New clsSearcher

    Private sourceobject As Object
    Private allowdrag As Boolean

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
            If ._currentobject IsNot Nothing AndAlso ._currentobject.userWorkstations.Count > 0 Then
                ._currentselectedobjects = .searcher.BasicSearchSync(
                    New clsDirectoryObject(._currentobject.Domain.DefaultNamingContext, ._currentobject.Domain),
                    New clsFilter("""" & Join(._currentobject.userWorkstations, """/""") & """", Nothing, New clsSearchObjectClasses(False, False, True, False, False)))
            End If
            .lvSelectedObjects.ItemsSource = If(._currentselectedobjects IsNot Nothing, ._currentselectedobjects, Nothing)
            .lvDomainObjects.ItemsSource = If(._currentobject IsNot Nothing, ._currentdomainobjects, Nothing)
        End With
    End Sub

    Private Async Sub tbDomainObjectsFilter_KeyDown(sender As Object, e As KeyEventArgs) Handles tbDomainObjectsFilter.KeyDown
        If e.Key = Key.Enter Then
            _currentdomainobjects.Clear()
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

    Private Sub lv_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles lvSelectedObjects.PreviewMouseLeftButtonDown,
                                                                                                   lvDomainObjects.PreviewMouseLeftButtonDown
        Dim listView As ListView = TryCast(sender, ListView)
        allowdrag = e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth And e.GetPosition(sender).Y < listView.ActualHeight - SystemParameters.HorizontalScrollBarHeight
    End Sub

    Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvSelectedObjects.MouseMove,
                                                                            lvDomainObjects.MouseMove
        Dim listView As ListView = TryCast(sender, ListView)

        If e.LeftButton = MouseButtonState.Pressed And listView.SelectedItem IsNot Nothing And allowdrag Then
            sourceobject = listView

            Dim obj As clsDirectoryObject = CType(listView.SelectedItem, clsDirectoryObject)
            Dim dragData As New DataObject("clsDirectoryObject", obj)

            DragDrop.DoDragDrop(listView, dragData, DragDropEffects.Move)
        End If
    End Sub

    Private Sub lv_DragEnter(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.DragEnter,
                                                                            lvDomainObjects.DragEnter

        If Not e.Data.GetDataPresent("clsDirectoryObject") OrElse sender Is sourceobject Then
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub lv_Drop(sender As Object, e As DragEventArgs) Handles lvSelectedObjects.Drop,
                                                                        lvDomainObjects.Drop

        If e.Data.GetDataPresent("clsDirectoryObject") And sender IsNot sourceobject Then
            Dim draggedObject As clsDirectoryObject = TryCast(e.Data.GetData("clsDirectoryObject"), clsDirectoryObject)

            If sender Is lvSelectedObjects Then ' adding computer
                If draggedObject Is Nothing Then Exit Sub
                AddComputer(draggedObject)
            Else
                If draggedObject Is Nothing Then Exit Sub
                RemoveComputer(draggedObject)
            End If

        End If
    End Sub

    Private Sub AddComputer([object] As clsDirectoryObject)
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

    Private Sub RemoveComputer([object] As clsDirectoryObject)
        Try
            _currentselectedobjects.Remove([object])
            _currentobject.userWorkstations = _currentselectedobjects.Select(Function(x As clsDirectoryObject) x.name).ToArray
        Catch ex As Exception
            ThrowException(ex, "RemoveComputer")
        End Try
    End Sub

    Private Sub ctlUserWorkstations_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbDomainObjectsFilter.Focus()
    End Sub
End Class
