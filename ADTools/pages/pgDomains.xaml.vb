Imports IRegisty
Imports IPrompt.VisualBasic

Public Class pgDomains

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

    Private Sub wndDomains_Unloaded() Handles Me.Unloaded
        Array.ForEach(Of String)(regDomains.GetSubKeyNames, New Action(Of String)(Sub(p) regDomains.DeleteSubKeyTree(p, False)))
        IRegistrySerializer.Serialize(domains, regDomains)
    End Sub

    Private Sub btnSearchRootBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnSearchRootBrowse.Click
        If lvDomains.SelectedItem Is Nothing Then Exit Sub

        Dim domainbrowser As New pgDomainBrowser
        Dim domain As clsDomain = CType(lvDomains.SelectedItem, clsDomain)

        domainbrowser.rootobject = New clsDirectoryObject(domain.DefaultNamingContext, domain)
        ShowPage(domainbrowser, True, Window.GetWindow(Me), True)

        'TODO If domainbrowser.DialogResult = True AndAlso domainbrowser.currentobject IsNot Nothing Then
        domain.SearchRoot = domainbrowser.currentobject.distinguishedName
        'End If
    End Sub

    Private Sub hlTemplateHelp_Click(sender As Object, e As RoutedEventArgs) Handles hlTemplateHelp.Click
        IMsgBox(My.Resources.wndDomains_lbl_PatternsExampleDefinition, vbOKOnly + vbInformation, My.Resources.wndDomains_lbl_PatternsExampleDefinitionHandlebars)
    End Sub

    Private Sub tabctlDomain_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tabctlDomain.SelectionChanged
        If tabctlDomain.SelectedIndex = 3 Then
            ctlMemberOf.InitializeAsync()
        End If
    End Sub
End Class
