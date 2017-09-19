Imports System.ComponentModel


Public Class clsSorter

#Region "Attached properties"

    Public Shared Function GetCommand(obj As DependencyObject) As ICommand
        Return DirectCast(obj.GetValue(CommandProperty), ICommand)
    End Function

    Public Shared Sub SetCommand(obj As DependencyObject, value As ICommand)
        obj.SetValue(CommandProperty, value)
    End Sub

    ' Using a DependencyProperty as the backing store for Command.  This enables animation, styling, binding, etc...
    Public Shared ReadOnly CommandProperty As DependencyProperty = DependencyProperty.RegisterAttached(
        "Command",
        GetType(ICommand),
        GetType(clsSorter),
        New UIPropertyMetadata(Nothing,
                               Sub(o, e)
                                   Dim listView As ItemsControl = TryCast(o, ItemsControl)
                                   If listView IsNot Nothing Then
                                       If Not GetAutoSort(listView) Then
                                           ' Don't change click handler if AutoSort enabled
                                           If e.OldValue IsNot Nothing AndAlso e.NewValue Is Nothing Then
                                               listView.[RemoveHandler](GridViewColumnHeader.ClickEvent, New RoutedEventHandler(AddressOf ColumnHeader_Click))
                                           End If
                                           If e.OldValue Is Nothing AndAlso e.NewValue IsNot Nothing Then
                                               listView.[AddHandler](GridViewColumnHeader.ClickEvent, New RoutedEventHandler(AddressOf ColumnHeader_Click))
                                           End If
                                       End If
                                   End If
                               End Sub))

    Public Shared Function GetAutoSort(obj As DependencyObject) As Boolean
        Return CBool(obj.GetValue(AutoSortProperty))
    End Function

    Public Shared Sub SetAutoSort(obj As DependencyObject, value As Boolean)
        obj.SetValue(AutoSortProperty, value)
    End Sub

    ' Using a DependencyProperty as the backing store for AutoSort.  This enables animation, styling, binding, etc...
    Public Shared ReadOnly AutoSortProperty As DependencyProperty = DependencyProperty.RegisterAttached(
        "AutoSort",
        GetType(Boolean),
        GetType(clsSorter),
        New UIPropertyMetadata(False,
                               Sub(o, e)
                                   Dim listView As ListView = TryCast(o, ListView)
                                   If listView IsNot Nothing Then
                                       If GetCommand(listView) Is Nothing Then
                                           ' Don't change click handler if a command is set
                                           Dim oldValue As Boolean = CBool(e.OldValue)
                                           Dim newValue As Boolean = CBool(e.NewValue)
                                           If oldValue AndAlso Not newValue Then
                                               listView.[RemoveHandler](GridViewColumnHeader.ClickEvent, New RoutedEventHandler(AddressOf ColumnHeader_Click))
                                           End If
                                           If Not oldValue AndAlso newValue Then
                                               listView.[AddHandler](GridViewColumnHeader.ClickEvent, New RoutedEventHandler(AddressOf ColumnHeader_Click))
                                           End If
                                       End If
                                   End If
                               End Sub))

    Public Shared Function GetPropertyName(obj As DependencyObject) As String
        If obj Is Nothing Then Return Nothing
        Return DirectCast(obj.GetValue(PropertyNameProperty), String)
    End Function

    Public Shared Sub SetPropertyName(obj As DependencyObject, value As String)
        obj.SetValue(PropertyNameProperty, value)
    End Sub

    ' Using a DependencyProperty as the backing store for PropertyName.  This enables animation, styling, binding, etc...
    Public Shared ReadOnly PropertyNameProperty As DependencyProperty = DependencyProperty.RegisterAttached("PropertyName", GetType(String), GetType(clsSorter), New UIPropertyMetadata(Nothing))

#End Region

#Region "Column header click event handler"

    Private Shared Sub ColumnHeader_Click(sender As Object, e As RoutedEventArgs)
        Dim headerClicked As GridViewColumnHeader = TryCast(e.OriginalSource, GridViewColumnHeader)
        If headerClicked IsNot Nothing Then
            Dim propertyName As String = GetPropertyName(headerClicked.Column)
            If Not String.IsNullOrEmpty(propertyName) Then
                Dim listView As ListView = GetAncestor(Of ListView)(headerClicked)
                If listView IsNot Nothing Then
                    Dim command As ICommand = GetCommand(listView)
                    If command IsNot Nothing Then
                        If command.CanExecute(propertyName) Then
                            command.Execute(propertyName)
                        End If
                    ElseIf GetAutoSort(listView) Then
                        ApplySort(listView.Items, propertyName)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Helper methods"

    Public Shared Function GetAncestor(Of T As DependencyObject)(reference As DependencyObject) As T
        Dim parent As DependencyObject = VisualTreeHelper.GetParent(reference)
        While Not (TypeOf parent Is T)
            parent = VisualTreeHelper.GetParent(parent)
        End While
        If parent IsNot Nothing Then
            Return DirectCast(parent, T)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Sub ApplySort(view As ICollectionView, propertyName As String)
        Dim direction As ListSortDirection = ListSortDirection.Ascending
        If view.SortDescriptions.Count > 0 Then
            Dim currentSort As SortDescription = view.SortDescriptions(0)
            If currentSort.PropertyName = propertyName Then
                If currentSort.Direction = ListSortDirection.Ascending Then
                    direction = ListSortDirection.Descending
                Else
                    direction = ListSortDirection.Ascending
                End If
            End If
            view.SortDescriptions.Clear()
        End If
        If Not String.IsNullOrEmpty(propertyName) Then
            view.SortDescriptions.Add(New SortDescription(propertyName, direction))
        End If
    End Sub

#End Region

End Class
