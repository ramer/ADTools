
Imports System.ComponentModel
Imports IRegisty

Public Class wndDomains

    Private passwordchanged As Boolean

    Private Sub wndDomains_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        lvDomains.ItemsSource = domains
    End Sub

    Private Sub btnDomainsAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnDomainsAdd.Click
        Dim newdomain = New clsDomain()
        domains.Add(newdomain)
        lvDomains.SelectedItem = newdomain
        tabctlDomain.SelectedIndex = 0
        passwordchanged = False
        pbPassword.Password = ""
        tbDomainName.Focus()
    End Sub

    Private Sub btnDomainsRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnDomainsRemove.Click
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        If domains.Contains(lvDomains.SelectedItem) Then domains.Remove(lvDomains.SelectedItem)
    End Sub

    Private Sub pbPassword_LostFocus(sender As Object, e As RoutedEventArgs) Handles pbPassword.LostFocus
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        If passwordchanged Then
            CType(lvDomains.SelectedItem, clsDomain).Password = CType(sender, PasswordBox).Password
        Else
            If Not String.IsNullOrEmpty(CType(lvDomains.SelectedItem, clsDomain).Password) Then tblckPassword.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub lvDomains_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lvDomains.SelectionChanged
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        tblckPassword.Visibility = If(String.IsNullOrEmpty(CType(lvDomains.SelectedItem, clsDomain).Password), Visibility.Collapsed, Visibility.Visible)
        passwordchanged = False
        pbPassword.Password = ""
    End Sub

    Private Sub pbPassword_GotFocus(sender As Object, e As RoutedEventArgs) Handles pbPassword.GotFocus
        tblckPassword.Visibility = Visibility.Collapsed
    End Sub

    Private Sub pbPassword_PasswordChanged(sender As Object, e As RoutedEventArgs) Handles pbPassword.PasswordChanged
        passwordchanged = True
    End Sub

    Private Async Sub btnConnect_Click(sender As Object, e As RoutedEventArgs) Handles btnConnect.Click
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        cap.Visibility = Visibility.Visible
        Await CType(lvDomains.SelectedItem, clsDomain).ConnectAsync()
        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub wndDomains_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Array.ForEach(Of String)(regDomains.GetSubKeyNames, New Action(Of String)(Sub(p) regDomains.DeleteSubKeyTree(p, False)))
        IRegistrySerializer.Serialize(domains, regDomains)
    End Sub

End Class
