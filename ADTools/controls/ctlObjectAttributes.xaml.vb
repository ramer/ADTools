Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Threading.Tasks

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
        InitializeComponent()

        cvsAttributes = New CollectionViewSource() With {.Source = attributes}
        cvAttributes = cvsAttributes.View
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
    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing Then Exit Sub

        cap.Visibility = Visibility.Visible
        Dim attrs As New ObservableCollection(Of clsAttribute)

        Await Task.Run(
            Sub()
                attrs = _currentobject.AllAttributes
            End Sub)

        attributes.Clear()

        For Each a In attrs
            attributes.Add(a)
        Next

        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub tbAttributesFilter_TextChanged(sender As Object, e As TextChangedEventArgs) Handles tbAttributesFilter.TextChanged
        cvAttributes.Filter = New Predicate(Of Object)(
            Function(a As clsAttribute)
                Dim filter As String = tbAttributesFilter.Text

                If tbAttributesFilter.Text = "*" Then
                    Return a.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(a.Value.ToString)
                Else
                    Return a.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 Or (a.Value IsNot Nothing AndAlso a.Value.ToString.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                End If
            End Function)
    End Sub
End Class
