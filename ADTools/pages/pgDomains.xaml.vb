Imports IRegisty
Imports IPrompt.VisualBasic
Imports System.ComponentModel
Imports System.Collections.ObjectModel

Public Class pgDomains

    Private model As pgDomainsVM

    Sub New()
        InitializeComponent()
        model = DataContext
    End Sub

    Private Sub btnDomainsAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnDomainsAdd.Click
        Dim newdomain = New clsDomain()
        domains.Add(newdomain)
        lvDomains.SelectedItem = newdomain
        tabctlDomain.SelectedIndex = 0
        tbDomainName.Focus()
    End Sub

    Private Sub btnDomainsRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnDomainsRemove.Click
        If lvDomains.SelectedItem Is Nothing Then Exit Sub
        If domains.Contains(lvDomains.SelectedItem) Then domains.Remove(lvDomains.SelectedItem)
    End Sub

    'Private Sub pbPassword_LostFocus(sender As Object, e As RoutedEventArgs) Handles pbPassword.LostFocus
    '    If lvDomains.SelectedItem Is Nothing Then Exit Sub
    '    If passwordchanged Then
    '        CType(lvDomains.SelectedItem, clsDomain).Password = CType(sender, PasswordBox).Password
    '    Else
    '        If Not String.IsNullOrEmpty(CType(lvDomains.SelectedItem, clsDomain).Password) Then tblckPassword.Visibility = Visibility.Visible
    '    End If
    'End Sub

    Private passwordchanged As Boolean
    Private Sub lvDomains_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lvDomains.SelectionChanged
        pbPassword.Password = ""
        passwordchanged = False
    End Sub

    Private Sub pbPassword_PasswordChanged(sender As Object, e As RoutedEventArgs) Handles pbPassword.PasswordChanged
        passwordchanged = True
    End Sub

    Private Async Sub btnConnect_Click(sender As Object, e As RoutedEventArgs) Handles btnConnect.Click
        If model.SelectedDomain Is Nothing Then Exit Sub
        If passwordchanged Then model.SelectedDomain.Password = pbPassword.Password

        cap.Visibility = Visibility.Visible
        Await CType(lvDomains.SelectedItem, clsDomain).ConnectAsync()
        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub str_Domains_Unloaded() Handles Me.Unloaded
        Array.ForEach(regDomains.GetSubKeyNames, New Action(Of String)(Sub(p) regDomains.DeleteSubKeyTree(p, False)))
        IRegistrySerializer.Serialize(domains, regDomains)
    End Sub

    Private Sub btnSearchRootBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnSearchRootBrowse.Click
        If model.SelectedDomain Is Nothing Then Exit Sub

        Dim domain As clsDomain = CType(lvDomains.SelectedItem, clsDomain)
        Dim domainbrowser As New pgDomainBrowser(model.SelectedDomain.UnderlyingObject)
        AddHandler domainbrowser.Return, AddressOf domainbrowserReturn
        NavigationService.Navigate(domainbrowser)
    End Sub

    Public Sub domainbrowserReturn(sender As Object, e As ReturnEventArgs(Of clsDirectoryObject))
        If model.SelectedDomain Is Nothing OrElse e.Result Is Nothing Then Exit Sub
        model.SelectedDomain.SearchRoot = e.Result.distinguishedName
    End Sub

    Private Sub hlTemplateHelp_Click(sender As Object, e As RoutedEventArgs) Handles hlTemplateHelp.Click
        IMsgBox(My.Resources.str_PatternsExampleDefinition2, vbOKOnly + vbInformation, My.Resources.str_PatternsExampleDefinitionHandlebars)
    End Sub

    Private Sub tabctlDomain_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tabctlDomain.SelectionChanged
        If tabctlDomain.SelectedIndex = 3 Then
            ctlMemberOf.InitializeAsync()
        End If
    End Sub
End Class

Public Class pgDomainsVM
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New()

    End Sub

    Private _selecteddomain As clsDomain
    Public Property SelectedDomain() As clsDomain
        Get
            Return _selecteddomain
        End Get
        Set(ByVal value As clsDomain)
            _selecteddomain = value
            NotifyPropertyChanged("SelectedDomain")
        End Set
    End Property

    Public Property Domains() As ObservableCollection(Of clsDomain)
        Get
            Return mdlTools.domains
        End Get
        Set(ByVal value As ObservableCollection(Of clsDomain))
            mdlTools.domains = value
        End Set
    End Property


End Class
