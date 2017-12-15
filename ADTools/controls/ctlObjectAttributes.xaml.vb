Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class ctlObjectAttributes

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlObjectAttributes),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject

    Private attributes As New ObservableCollection(Of clsAttribute)
    Private cvsAttributes As New CollectionViewSource
    Private cvAttributes As ICollectionView

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Sub New()
        cvsAttributes = New CollectionViewSource() With {.Source = attributes}
        cvAttributes = cvsAttributes.View
        cvAttributes.SortDescriptions.Add(New SortDescription("Name", ListSortDirection.Ascending))

        InitializeComponent()

        dgAttributes.ItemsSource = cvAttributes
    End Sub

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlObjectAttributes = CType(d, ctlObjectAttributes)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Sub ctlAttributes_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbAttributesFilter.Focus()

        InitializeAsync()
    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing Then Exit Sub

        If attributes.Count = 0 Then
            cap.Visibility = Visibility.Visible

            Dim ldapattributes As New ObservableCollection(Of clsAttribute)
            ldapattributes = Await Task.Run(Function()
                                                _currentobject.RefreshAllAllowedAttributes()
                                                Return _currentobject.AllAttributes
                                            End Function)

            attributes.Clear()
            For Each a In ldapattributes
                attributes.Add(New clsAttribute(a.Name, "", a.Value))
            Next

            cap.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Sub btnAttributesRefresh_Click(sender As Object, e As RoutedEventArgs) Handles btnAttributesRefresh.Click
        RefreshAsync()
    End Sub

    Public Async Sub RefreshAsync()
        If _currentobject Is Nothing Then Exit Sub

        cap.Visibility = Visibility.Visible
        Dim ldapattributes As New ObservableCollection(Of clsAttribute)

        Await Task.Run(
            Sub()
                _currentobject.Refresh()
                ldapattributes = _currentobject.AllAttributes
            End Sub)

        For I = 0 To ldapattributes.Count - 1
            For J = 0 To attributes.Count - 1
                If ldapattributes(I).Name = attributes(J).Name Then
                    'magic
                    If ldapattributes(I).Value Is Nothing And attributes(J).Value Is Nothing Then Continue For

                    If (ldapattributes(I).Value IsNot Nothing And attributes(J).Value Is Nothing) OrElse
                       (ldapattributes(I).Value Is Nothing And attributes(J).Value IsNot Nothing) OrElse
                       (Not ldapattributes(I).Value.ToString = attributes(J).Value.ToString) Then
                        attributes(J).NewValue = ldapattributes(I).Value
                    End If

                End If
            Next
        Next

        ApplyFilter()

        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub rbAttributesWithValue_Checked(sender As Object, e As RoutedEventArgs) Handles rbAttributesWithValue.Checked, rbAttributesAll.Checked, rbAttributesChanged.Checked
        ApplyFilter()
    End Sub

    Private Sub tbAttributesFilter_TextChanged(sender As Object, e As TextChangedEventArgs) Handles tbAttributesFilter.TextChanged
        ApplyFilter()
    End Sub

    Private Sub dgAttributes_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgAttributes.SelectionChanged
        e.Handled = True
    End Sub

    Private Sub ctxmnuCopy_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuCopy.Click
        If dgAttributes.SelectedItems Is Nothing Then Exit Sub

        Dim str As String = ""

        For Each a As clsAttribute In dgAttributes.SelectedItems
            str &= a.Name & vbTab & a.Value & vbCrLf
        Next

        Clipboard.SetText(str)
    End Sub

    Private Sub ctxmnuCopyValues_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuCopyValues.Click
        If dgAttributes.SelectedItems Is Nothing Then Exit Sub

        Dim str As String = ""

        For Each a As clsAttribute In dgAttributes.SelectedItems
            str &= a.Value & vbCrLf
        Next

        Clipboard.SetText(str)
    End Sub

    Private Sub ApplyFilter()
        cvAttributes.Filter = New Predicate(Of Object)(
            Function(a As clsAttribute)
                Dim filter As String = tbAttributesFilter.Text

                If rbAttributesWithValue.IsChecked Then ' with values
                    If filter.Length = 0 Then
                        Return (a.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(a.Value.ToString))
                    Else
                        Return (a.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(a.Value.ToString)) _
                               AndAlso
                               ((a.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) Or
                               (a.Value IsNot Nothing AndAlso a.Value.ToString.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) Or
                               (a.NewValue IsNot Nothing AndAlso a.NewValue.ToString.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0))
                    End If
                ElseIf rbAttributesChanged.IsChecked Then 'changed
                    If filter.Length = 0 Then
                        Return (a.Value Is Nothing And a.NewValue IsNot Nothing) OrElse
                               (a.Value IsNot Nothing And a.NewValue IsNot Nothing) AndAlso
                               (Not a.Value.Equals(a.NewValue))
                    Else
                        Return ((a.Value Is Nothing And a.NewValue IsNot Nothing) OrElse
                               (a.Value IsNot Nothing And a.NewValue IsNot Nothing) AndAlso
                               (Not a.Value.Equals(a.NewValue))) _
                               AndAlso
                               ((a.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) Or
                               (a.Value IsNot Nothing AndAlso a.Value.ToString.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) Or
                               (a.NewValue IsNot Nothing AndAlso a.NewValue.ToString.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0))
                    End If
                Else ' all
                    If filter.Length = 0 Then
                        Return True
                    Else
                        Return (a.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) Or
                               (a.Value IsNot Nothing AndAlso a.Value.ToString.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) Or
                               (a.NewValue IsNot Nothing AndAlso a.NewValue.ToString.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                    End If

                End If

            End Function)
    End Sub

End Class
