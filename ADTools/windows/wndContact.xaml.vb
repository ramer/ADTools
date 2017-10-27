Imports System.ComponentModel

Public Class wndContact

    Public Property currentobject As clsDirectoryObject

    Private Sub wndContact_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.DataContext = currentobject
        ctlMemberOf.CurrentObject = currentobject
        ctlContact.CurrentObject = currentobject
        ctlAttributes.CurrentObject = currentobject

        tabctlContactExchange.IsEnabled = currentobject.Domain.UseExchange
    End Sub

    Private Sub wnd_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndContact(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub сmboTelephoneNumber_DropDownOpened(sender As Object, e As EventArgs) Handles сmboTelephoneNumber.DropDownOpened
        сmboTelephoneNumber.ItemsSource = GetNextDomainTelephoneNumbers(currentobject.Domain)
    End Sub

    Private Sub сmboTelephoneNumber_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles сmboTelephoneNumber.SelectionChanged
        e.Handled = True
    End Sub

    Private Sub Manager_hyperlink_click(sender As Object, e As RequestNavigateEventArgs)
        ShowDirectoryObjectProperties(currentobject.manager, Me)
    End Sub

    Private Sub tabctlContact_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tabctlContact.SelectionChanged
        If tabctlContact.SelectedIndex = 1 Then
            ctlMemberOf.InitializeAsync()
        ElseIf tabctlContact.SelectedIndex = 2 Then
            ctlContact.InitializeAsync()
        ElseIf tabctlContact.SelectedIndex = 3 Then
            ctlAttributes.InitializeAsync()
        End If
    End Sub

End Class
