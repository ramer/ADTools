Imports System.ComponentModel
Imports System.Net.NetworkInformation
Imports System.Windows.Threading
Imports IPrompt.VisualBasic

Public Class pgComputer
    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(pgComputer),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomainobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgComputer = CType(d, pgComputer)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Public Property events As New clsThreadSafeObservableCollection(Of clsEvent)
    Public WithEvents wmisearcher As New clsWmi

    Private Sub dtpPeriodFrom_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles dtpPeriodFrom.ValueChanged
        If dtpPeriodTo.Value < dtpPeriodFrom.Value Then dtpPeriodTo.Value = dtpPeriodFrom.Value
    End Sub

    Private Sub dtpPeriodTo_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles dtpPeriodTo.ValueChanged
        If dtpPeriodTo.Value < dtpPeriodFrom.Value Then dtpPeriodFrom.Value = dtpPeriodTo.Value
    End Sub

    Private Async Sub btnEventsSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnEventsSearch.Click
        pbSearch.Visibility = Visibility.Visible

        Await wmisearcher.BasicSearchWmiAsync(events, currentobject, dtpPeriodFrom.Value, dtpPeriodTo.Value, If(rbEventAll.IsChecked, 0, If(rbEventSuccess.IsChecked, 1, 2)))

        pbSearch.Visibility = Visibility.Hidden
    End Sub

    Private Sub dgEvents_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles dgEvents.MouseDoubleClick
        If dgEvents.SelectedItem Is Nothing Then Exit Sub
        IMsgBox(CType(dgEvents.SelectedItem, clsEvent).Message, vbOKOnly + vbInformation, CType(dgEvents.SelectedItem, clsEvent).CategoryString)
    End Sub

    Private Sub hlManagedBy_Click(sender As Object, e As RoutedEventArgs) Handles hlManagedBy.Click
        ShowDirectoryObjectProperties(CurrentObject.managedBy, Window.GetWindow(Me))
    End Sub

    Private Sub pgComputer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dgEvents.ItemsSource = events

        dtpPeriodTo.Value = Now
        dtpPeriodFrom.Value = Now.AddDays(-1)
    End Sub

End Class
