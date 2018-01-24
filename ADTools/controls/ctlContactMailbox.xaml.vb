Imports IPrompt.VisualBasic

Public Class ctlContactMailbox

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlContactMailbox),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject

    Public Property contact As clsContact

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
    End Sub

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlContactMailbox = CType(d, ctlContactMailbox)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Sub ctlContact_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbContact.Focus()

        InitializeAsync()
    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing Then Exit Sub

        If contact Is Nothing AndAlso _currentobject.Domain.UseExchange AndAlso _currentobject.Domain.ExchangeServer IsNot Nothing Then
            capexchange.Visibility = Visibility.Visible

            tbContact.Text = GetNextUserMailbox(_currentobject)

            contact = Await Task.Run(Function() New clsContact(_currentobject))

            Me.DataContext = contact

            capexchange.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Sub hlState_Click(sender As Object, e As RoutedEventArgs) Handles hlState.Click
        If contact IsNot Nothing AndAlso
           contact.ExchangeConnection IsNot Nothing AndAlso
           contact.ExchangeConnection.State IsNot Nothing AndAlso
           contact.ExchangeConnection.State.Reason IsNot Nothing Then

            IMsgBox(contact.ExchangeConnection.State.Reason.Message, vbExclamation + vbOKOnly, My.Resources.str_ConnectionError, Window.GetWindow(Me))
        End If
    End Sub

    Private Sub lvEmailAddresses_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lvEmailAddresses.SelectionChanged
        e.Handled = True

        If lvEmailAddresses.SelectedItem Is Nothing Then
            tbContact.Text = ""
            cmboContactDomain.SelectedItem = Nothing
            Exit Sub
        End If

        Dim a As String() = CType(lvEmailAddresses.SelectedItem, clsEmailAddress).Address.Split({"@"}, StringSplitOptions.RemoveEmptyEntries)
        If a.Count < 2 Then
            tbContact.Text = ""
            cmboContactDomain.SelectedItem = Nothing
        Else
            tbContact.Text = a(0)
            cmboContactDomain.SelectedItem = a(1)
        End If
    End Sub

    Private Async Sub btnContactAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnContactAdd.Click
        If contact Is Nothing Then Exit Sub
        capexchange.Visibility = Visibility.Visible

        Dim name As String = tbContact.Text
        Dim domain As String = cmboContactDomain.Text
        Await Task.Run(Sub() contact.Add(name, domain))

        tbContact.Text = ""
        capexchange.Visibility = Visibility.Hidden
    End Sub

    Private Async Sub btnContactEdit_Click(sender As Object, e As RoutedEventArgs) Handles btnContactEdit.Click
        If contact Is Nothing Then Exit Sub
        capexchange.Visibility = Visibility.Visible

        Dim newname As String = tbContact.Text
        Dim newdomain As String = cmboContactDomain.Text
        Dim oldaddress As clsEmailAddress = lvEmailAddresses.SelectedItem
        Await Task.Run(Sub() contact.Edit(newname, newdomain, oldaddress))

        tbContact.Text = ""
        capexchange.Visibility = Visibility.Hidden
    End Sub

    Private Async Sub btnContactRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnContactRemove.Click
        If contact Is Nothing Then Exit Sub
        capexchange.Visibility = Visibility.Visible

        Dim oldaddress As clsEmailAddress = lvEmailAddresses.SelectedItem

        If oldaddress.IsPrimary AndAlso IMsgBox(My.Resources.str_RemovingPrimaryAddress, vbYesNo + vbQuestion, My.Resources.str_RemovingPrimaryAddressTitle) = vbNo Then
            capexchange.Visibility = Visibility.Hidden
            Exit Sub
        End If

        Await Task.Run(Sub() contact.Remove(oldaddress))

        tbContact.Text = ""
        capexchange.Visibility = Visibility.Hidden
    End Sub

    Private Async Sub btnContactSetPrimary_Click(sender As Object, e As RoutedEventArgs) Handles btnContactSetPrimary.Click
        If contact Is Nothing Then Exit Sub
        If CType(lvEmailAddresses.SelectedItem, clsEmailAddress).IsPrimary = False Then CurrentObject.mail = CType(lvEmailAddresses.SelectedItem, clsEmailAddress).Address

        capexchange.Visibility = Visibility.Visible

        Dim mail As clsEmailAddress = lvEmailAddresses.SelectedItem
        Await Task.Run(Sub() contact.SetPrimary(mail))

        capexchange.Visibility = Visibility.Hidden
    End Sub

    Private Sub ctlContact_Unloaded(sender As Object, e As RoutedEventArgs) Handles Me.Unloaded
        If contact Is Nothing Then Exit Sub

        contact.Close()
    End Sub

End Class
