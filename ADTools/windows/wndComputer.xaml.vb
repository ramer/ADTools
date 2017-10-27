Imports System.ComponentModel
Imports System.Net.NetworkInformation
Imports System.Windows.Threading
Imports IPrompt.VisualBasic

Public Class wndComputer

    Public Property currentobject As clsDirectoryObject

    Public Property events As New clsThreadSafeObservableCollection(Of clsEvent)
    Public WithEvents wmisearcher As New clsWmi

    Private Sub wndComputer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.DataContext = currentobject
        ctlMemberOf.CurrentObject = currentobject
        ctlNet.CurrentObject = currentobject
        ctlAttributes.CurrentObject = currentobject

        dgEvents.ItemsSource = events

        dtpPeriodTo.Value = Now
        dtpPeriodFrom.Value = Now.AddDays(-1)
    End Sub

    Private Sub wndComputer_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub tabctlComputer_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tabctlComputer.SelectionChanged
        If tabctlComputer.SelectedIndex = 1 Then
            ctlMemberOf.InitializeAsync()
        ElseIf tabctlComputer.SelectedIndex = 2 Then
            ctlNet.InitializeAsync()
        ElseIf tabctlComputer.SelectedIndex = 4 Then
            ctlAttributes.InitializeAsync()
        End If
    End Sub

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


End Class
