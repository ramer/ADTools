Imports System.ComponentModel
Imports System.Threading.Tasks

Public Class wndContact

    Public Property currentobject As clsDirectoryObject
    'Public Property contact As clsContact

    Private Sub wndContact_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.DataContext = currentobject
        ctlMemberOf.CurrentObject = currentobject
        ctlAttributes.CurrentObject = currentobject

        tabctlContactExchange.IsEnabled = currentobject.Domain.UseExchange
    End Sub

    Private Sub wnd_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        'If contact IsNot Nothing Then contact.Close()
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
            'ElseIf tabctlContact.SelectedIndex = 2 Then
            '    tbMailbox.Focus()

            '    If contact Is Nothing AndAlso currentobject.Domain.UseExchange AndAlso currentobject.Domain.ExchangeServer IsNot Nothing Then
            '        cap.Visibility = Visibility.Visible
            '        contact = Await Task.Run(Function() New clsContact(currentobject))
            '        tabctlContactExchange.DataContext = contact
            '        cap.Visibility = Visibility.Hidden
            '    End If

        ElseIf tabctlContact.SelectedIndex = 3 Then
            ctlAttributes.InitializeAsync()
        End If
    End Sub

    'Private Sub lvEmailAddresses_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lvEmailAddresses.SelectionChanged
    '    e.Handled = True

    '    If lvEmailAddresses.SelectedItem Is Nothing Then
    '        tbMailbox.Text = ""
    '        cmboMailboxDomain.SelectedItem = Nothing
    '        Exit Sub
    '    End If

    '    Dim a As String() = CType(lvEmailAddresses.SelectedItem, clsEmailAddress).Address.Split({"@"}, StringSplitOptions.RemoveEmptyEntries)
    '    If a.Count < 2 Then
    '        tbMailbox.Text = ""
    '        cmboMailboxDomain.SelectedItem = Nothing
    '    Else
    '        tbMailbox.Text = a(0)
    '        cmboMailboxDomain.SelectedItem = a(1)
    '    End If
    'End Sub

    'Private Async Sub btnMailboxAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxAdd.Click
    '    cap.Visibility = Visibility.Visible

    '    Dim name As String = tbMailbox.Text
    '    Dim domain As String = cmboMailboxDomain.Text
    '    Await Task.Run(Sub() contact.Add(name, domain))

    '    tbMailbox.Text = ""
    '    cap.Visibility = Visibility.Hidden
    'End Sub

    'Private Async Sub btnMailboxEdit_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxEdit.Click
    '    cap.Visibility = Visibility.Visible

    '    Dim newname As String = tbMailbox.Text
    '    Dim newdomain As String = cmboMailboxDomain.Text
    '    Dim oldaddress As clsEmailAddress = lvEmailAddresses.SelectedItem
    '    Await Task.Run(Sub() contact.Edit(newname, newdomain, oldaddress))

    '    tbMailbox.Text = ""
    '    cap.Visibility = Visibility.Hidden
    'End Sub

    'Private Async Sub btnMailboxRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxRemove.Click
    '    cap.Visibility = Visibility.Visible

    '    Dim oldaddress As clsEmailAddress = lvEmailAddresses.SelectedItem
    '    Await Task.Run(Sub() contact.Remove(oldaddress))

    '    tbMailbox.Text = ""
    '    cap.Visibility = Visibility.Hidden
    'End Sub

    'Private Async Sub btnMailboxSetPrimary_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxSetPrimary.Click
    '    If CType(lvEmailAddresses.SelectedItem, clsEmailAddress).IsPrimary = False Then currentobject.mail = CType(lvEmailAddresses.SelectedItem, clsEmailAddress).Address

    '    cap.Visibility = Visibility.Visible

    '    Dim mail As clsEmailAddress = lvEmailAddresses.SelectedItem
    '    Await Task.Run(Sub() contact.SetPrimary(mail))

    '    cap.Visibility = Visibility.Hidden
    'End Sub

End Class
