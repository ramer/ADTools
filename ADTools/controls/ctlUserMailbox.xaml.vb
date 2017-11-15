Imports IPrompt.VisualBasic

Public Class ctlUserMailbox

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlUserMailbox),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject

    Public Property mailbox As clsMailbox

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
        Dim instance As ctlUserMailbox = CType(d, ctlUserMailbox)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Sub ctlMailbox_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        tbMailbox.Focus()

        InitializeAsync()
    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing OrElse _currentobject.Entry Is Nothing Then Exit Sub

        If mailbox Is Nothing AndAlso _currentobject.Domain.UseExchange AndAlso _currentobject.Domain.ExchangeServer IsNot Nothing Then
            capexchange.Visibility = Visibility.Visible

            tbMailbox.Text = GetNextUserMailbox(_currentobject)

            mailbox = Await Task.Run(Function() New clsMailbox(_currentobject))

            Me.DataContext = mailbox

            capexchange.Visibility = Visibility.Hidden
            End If
    End Sub

    Private Sub hlState_Click(sender As Object, e As RoutedEventArgs) Handles hlState.Click
        If mailbox IsNot Nothing AndAlso
           mailbox.ExchangeConnection IsNot Nothing AndAlso
           mailbox.ExchangeConnection.State IsNot Nothing AndAlso
           mailbox.ExchangeConnection.State.Reason IsNot Nothing Then

            IMsgBox(mailbox.ExchangeConnection.State.Reason.Message, vbExclamation + vbOKOnly, My.Resources.ctlMailbox_msg_ConnectionError, Window.GetWindow(Me))
        End If
    End Sub

    Private Sub lvEmailAddresses_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lvEmailAddresses.SelectionChanged
        e.Handled = True

        If lvEmailAddresses.SelectedItem Is Nothing Then
            tbMailbox.Text = ""
            cmboMailboxDomain.SelectedItem = Nothing
            Exit Sub
        End If

        Dim a As String() = CType(lvEmailAddresses.SelectedItem, clsEmailAddress).Address.Split({"@"}, StringSplitOptions.RemoveEmptyEntries)
        If a.Count < 2 Then
            tbMailbox.Text = ""
            cmboMailboxDomain.SelectedItem = Nothing
        Else
            tbMailbox.Text = a(0)
            cmboMailboxDomain.SelectedItem = a(1)
        End If
    End Sub

    Private Async Sub btnMailboxAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxAdd.Click
        If mailbox Is Nothing Then Exit Sub
        capexchange.Visibility = Visibility.Visible

        Dim name As String = tbMailbox.Text
        Dim domain As String = cmboMailboxDomain.Text
        Await Task.Run(Sub() mailbox.Add(name, domain))

        tbMailbox.Text = ""
        capexchange.Visibility = Visibility.Hidden
    End Sub

    Private Async Sub btnMailboxEdit_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxEdit.Click
        If mailbox Is Nothing Then Exit Sub
        capexchange.Visibility = Visibility.Visible

        Dim newname As String = tbMailbox.Text
        Dim newdomain As String = cmboMailboxDomain.Text
        Dim oldaddress As clsEmailAddress = lvEmailAddresses.SelectedItem
        Await Task.Run(Sub() mailbox.Edit(newname, newdomain, oldaddress))

        tbMailbox.Text = ""
        capexchange.Visibility = Visibility.Hidden
    End Sub

    Private Async Sub btnMailboxRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxRemove.Click
        If mailbox Is Nothing Then Exit Sub
        capexchange.Visibility = Visibility.Visible

        Dim oldaddress As clsEmailAddress = lvEmailAddresses.SelectedItem

        If oldaddress.IsPrimary AndAlso IMsgBox(My.Resources.ctlMailbox_msg_RemovingPrimaryAddress, vbYesNo + vbQuestion, My.Resources.ctlMailbox_msg_RemovingPrimaryAddressTitle) = vbNo Then
            capexchange.Visibility = Visibility.Hidden
            Exit Sub
        End If

        Await Task.Run(Sub() mailbox.Remove(oldaddress))

        tbMailbox.Text = ""
        capexchange.Visibility = Visibility.Hidden
    End Sub

    Private Async Sub btnMailboxSetPrimary_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxSetPrimary.Click
        If mailbox Is Nothing Then Exit Sub
        If CType(lvEmailAddresses.SelectedItem, clsEmailAddress).IsPrimary = False Then CurrentObject.mail = CType(lvEmailAddresses.SelectedItem, clsEmailAddress).Address

        capexchange.Visibility = Visibility.Visible

        Dim mail As clsEmailAddress = lvEmailAddresses.SelectedItem
        Await Task.Run(Sub() mailbox.SetPrimary(mail))

        capexchange.Visibility = Visibility.Hidden
    End Sub

    Private Sub btnMailboxQuota_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxQuota.Click
        If mailbox Is Nothing Then Exit Sub
        Dim w As New wndMailboxQuota
        w.mailbox = mailbox
        ShowWindow(w, True, Window.GetWindow(Me), True)
    End Sub

    Private Sub btnMailboxShare_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxShare.Click
        If mailbox Is Nothing Then Exit Sub
        Dim w As New wndMailboxShare
        w.currentuser = CurrentObject
        w.mailbox = mailbox
        ShowWindow(w, True, Window.GetWindow(Me), True)
    End Sub

    Private Sub ctlMailbox_Unloaded(sender As Object, e As RoutedEventArgs) Handles Me.Unloaded
        If mailbox Is Nothing Then Exit Sub

        mailbox.Close()
    End Sub

End Class
