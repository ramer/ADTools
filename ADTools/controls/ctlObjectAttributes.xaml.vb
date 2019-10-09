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
        cvAttributes.SortDescriptions.Add(New SortDescription("lDAPDisplayName", ListSortDirection.Ascending))

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

            Await Task.Run(Sub() _currentobject.RefreshAllAllowedAttributes())

            Dim allattributes As New List(Of clsAttribute)
            Await Task.Run(
                Sub()
                    For Each a In _currentobject.AllowedAttributes
                        allattributes.Add(_currentobject.GetAttribute(a))
                    Next
                End Sub)

            attributes.Clear()
            For Each a In allattributes
                attributes.Add(a)
            Next

            cap.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Sub btnAttributesRefresh_Click(sender As Object, e As RoutedEventArgs) Handles btnAttributesRefresh.Click
        'RefreshAsync()
    End Sub

    'Public Async Sub RefreshAsync()
    '    If _currentobject Is Nothing Then Exit Sub

    '    cap.Visibility = Visibility.Visible
    '    Dim ldapattributes As New ObservableCollection(Of clsAttribute)

    '    Await Task.Run(
    '        Sub()
    '            _currentobject.Refresh()
    '            ldapattributes = _currentobject.AllAttributes
    '        End Sub)

    '    For I = 0 To ldapattributes.Count - 1
    '        For J = 0 To attributes.Count - 1
    '            If ldapattributes(I).Name = attributes(J).Name Then
    '                'magic
    '                If ldapattributes(I).Value Is Nothing And attributes(J).Value Is Nothing Then Continue For

    '                If (ldapattributes(I).Value IsNot Nothing And attributes(J).Value Is Nothing) OrElse
    '                   (ldapattributes(I).Value Is Nothing And attributes(J).Value IsNot Nothing) OrElse
    '                   (Not ldapattributes(I).Value.ToString = attributes(J).Value.ToString) Then
    '                    attributes(J).NewValue = ldapattributes(I).Value
    '                End If

    '            End If
    '        Next
    '    Next

    '    ApplyFilter()

    '    cap.Visibility = Visibility.Hidden
    'End Sub

    Private Sub rbAttributesWithValue_Checked(sender As Object, e As RoutedEventArgs) Handles rbAttributesWithValue.Checked, rbAttributesAll.Checked
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

        For Each a As String In dgAttributes.SelectedItems
            Dim val = _currentobject.GetValue(a)
            If TypeOf val Is String Then str &= a & vbTab & val & vbCrLf
            If TypeOf val Is String() Then str &= a & vbTab & Join(CType(val, String()), ", ") & vbCrLf
        Next

        Clipboard.SetText(str)
    End Sub

    Private Sub ctxmnuCopyValues_Click(sender As Object, e As RoutedEventArgs) Handles ctxmnuCopyValues.Click
        If dgAttributes.SelectedItems Is Nothing Then Exit Sub

        Dim str As String = ""

        For Each a As clsAttribute In dgAttributes.SelectedItems
            Dim val = a.Value
            If TypeOf val Is String Then str &= val & vbCrLf
            If TypeOf val Is String() Then str &= Join(CType(val, String()), ", ") & vbCrLf
        Next

        Clipboard.SetText(str)
    End Sub

    Private Sub ApplyFilter()
        cvAttributes.Filter = New Predicate(Of Object)(
            Function(a As clsAttribute)
                Dim filter As String = tbAttributesFilter.Text

                If rbAttributesWithValue.IsChecked Then ' with values
                    If filter.Length = 0 Then
                        Return a.Value IsNot Nothing
                    Else
                        Return a.Value IsNot Nothing AndAlso a.lDAPDisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                    End If
                Else ' all
                    If filter.Length = 0 Then
                        Return True
                    Else
                        Return a.lDAPDisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                    End If
                End If

            End Function)
    End Sub

End Class
