Imports System.ComponentModel
Imports System.Threading.Tasks
Imports HandlebarsDotNet
Imports IPrompt.VisualBasic

Public Class wndUser

    Public Property currentobject As clsDirectoryObject
    'Public Property mailbox As clsMailbox

    Private Sub wndUser_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.DataContext = currentobject

        ctlUserWorkstations.DataContext = currentobject
        ctlUserWorkstations.CurrentObject = currentobject

        ctlMemberOf.DataContext = currentobject
        ctlMemberOf.CurrentObject = currentobject

        'tbMailbox.Text = GetNextUserMailbox(currentobject)

        'tabctlUserExchange.DataContext = mailbox
        tabctlUserExchange.IsEnabled = currentobject.Domain.UseExchange

        ctlAttributes.DataContext = currentobject
        ctlAttributes.CurrentObject = currentobject
    End Sub

    Private Sub wnd_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        'If mailbox IsNot Nothing Then mailbox.Close()
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub сmboTelephoneNumber_DropDownOpened(sender As Object, e As EventArgs) Handles сmboTelephoneNumber.DropDownOpened
        сmboTelephoneNumber.ItemsSource = GetNextDomainTelephoneNumbers(currentobject.Domain)
    End Sub

    Private Sub сmboTelephoneNumber_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles сmboTelephoneNumber.SelectionChanged
        e.Handled = True
    End Sub

    Private Sub btnResolveDepartment_Click(sender As Object, e As RoutedEventArgs) Handles btnResolveDepartment.Click
        If currentobject Is Nothing Then Exit Sub
        tbDepartment.Text = Replace(currentobject.Entry.Parent.Name, "OU=", "")
    End Sub

    Private Sub Manager_hyperlink_click(sender As Object, e As RequestNavigateEventArgs)
        ShowDirectoryObjectProperties(currentobject.manager, Me)
    End Sub

    Private Sub btnClearPhoto_Click(sender As Object, e As RoutedEventArgs) Handles btnClearPhoto.Click
        If IMsgBox("Вы уверены?", vbYesNo + vbQuestion, "Удаление фото") = MsgBoxResult.Yes Then currentobject.thumbnailPhoto = Nothing
    End Sub

    Private Sub imgPhoto_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgPhoto.MouseDown
        Dim w As wndUserPhoto

        For Each wnd As Window In Me.OwnedWindows
            If GetType(wndUserPhoto) Is wnd.GetType Then
                w = wnd

                w.Show()
                w.Activate()

                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal

                w.Topmost = True
                w.Topmost = False

                Exit Sub
            End If
        Next

        w = New wndUserPhoto With {.Owner = Me}

        w.imgPhoto.Source = imgPhoto.Source
        w.Show()
    End Sub

    Private Async Sub tabctlUser_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tabctlUser.SelectionChanged
        If tabctlUser.SelectedIndex = 0 Then
            tbGivenName.Focus()
        ElseIf tabctlUser.SelectedIndex = 1 Then
            tbUserPrincipalNameName.Focus()
        ElseIf tabctlUser.SelectedIndex = 2 Then
            ctlMemberOf.Focus()
        ElseIf tabctlUser.SelectedIndex = 3 Then
            'tbMailbox.Focus()

            'If mailbox Is Nothing AndAlso currentobject.Domain.UseExchange AndAlso currentobject.Domain.ExchangeServer IsNot Nothing Then
            '    capexchange.Visibility = Visibility.Visible
            '    mailbox = Await Task.Run(Function() New clsMailbox(currentobject))
            '    tabctlUserExchange.DataContext = mailbox
            '    capexchange.Visibility = Visibility.Hidden
            'End If

        ElseIf tabctlUser.SelectedIndex = 4 Then

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

    Private Sub btnResetPassword_Click(sender As Object, e As RoutedEventArgs) Handles btnResetPassword.Click
        Try
            If currentobject Is Nothing Then Exit Sub
            If IMsgBox("Вы уверены?", vbYesNo + vbQuestion, "Сброс пароля") = vbYes Then
                currentobject.ResetPassword()
                currentobject.passwordNeverExpires = False
                IMsgBox("Пароль сброшен.", vbOKOnly + vbInformation, "Сброс пароля")
            End If
        Catch ex As Exception
            ThrowException(ex, "btnResetPassword_Click")
        End Try
    End Sub

    Private Sub btnSetPassword_Click(sender As Object, e As RoutedEventArgs) Handles btnSetPassword.Click
        Try
            If currentobject Is Nothing Then Exit Sub
            Dim newpassword As String
            newpassword = IPasswordBox("Введите новый пароль:", "Смена пароля", vbQuestion)
            If String.IsNullOrEmpty(newpassword) Then Exit Sub
            currentobject.SetPassword(newpassword)
            IMsgBox("Пароль сброшен.", vbOKOnly + vbInformation, "Сброс пароля")
        Catch ex As Exception
            ThrowException(ex, "btnSetPassword_Click")
        End Try
    End Sub

    'Private Async Sub btnMailboxAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxAdd.Click
    '    capexchange.Visibility = Visibility.Visible

    '    Dim name As String = tbMailbox.Text
    '    Dim domain As String = cmboMailboxDomain.Text
    '    Await Task.Run(Sub() mailbox.Add(name, domain))

    '    tbMailbox.Text = ""
    '    capexchange.Visibility = Visibility.Hidden
    'End Sub

    'Private Async Sub btnMailboxEdit_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxEdit.Click
    '    capexchange.Visibility = Visibility.Visible

    '    Dim newname As String = tbMailbox.Text
    '    Dim newdomain As String = cmboMailboxDomain.Text
    '    Dim oldaddress As clsEmailAddress = lvEmailAddresses.SelectedItem
    '    Await Task.Run(Sub() mailbox.Edit(newname, newdomain, oldaddress))

    '    tbMailbox.Text = ""
    '    capexchange.Visibility = Visibility.Hidden
    'End Sub

    'Private Async Sub btnMailboxRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxRemove.Click
    '    capexchange.Visibility = Visibility.Visible

    '    Dim oldaddress As clsEmailAddress = lvEmailAddresses.SelectedItem

    '    If oldaddress.IsPrimary AndAlso IMsgBox("Удаление основного адреса подразумевает полное удаление почтового ящика" & vbCrLf & vbCrLf & "Вы уверены?", "Удаление почтового ящика", vbYesNo, vbQuestion) = vbNo Then
    '        capexchange.Visibility = Visibility.Hidden
    '        Exit Sub
    '    End If

    '    Await Task.Run(Sub() mailbox.Remove(oldaddress))

    '    tbMailbox.Text = ""
    '    capexchange.Visibility = Visibility.Hidden
    'End Sub

    'Private Async Sub btnMailboxSetPrimary_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxSetPrimary.Click
    '    If CType(lvEmailAddresses.SelectedItem, clsEmailAddress).IsPrimary = False Then currentobject.mail = CType(lvEmailAddresses.SelectedItem, clsEmailAddress).Address

    '    capexchange.Visibility = Visibility.Visible

    '    Dim mail As clsEmailAddress = lvEmailAddresses.SelectedItem
    '    Await Task.Run(Sub() mailbox.SetPrimary(mail))

    '    capexchange.Visibility = Visibility.Hidden
    'End Sub

    'Private Sub btnMailboxQuota_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxQuota.Click
    '    Dim w As wndMailboxQuota

    '    For Each wnd As Window In Me.OwnedWindows
    '        If GetType(wndMailboxQuota) Is wnd.GetType Then
    '            w = wnd

    '            w.Show()
    '            w.Activate()

    '            If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal

    '            w.Topmost = True
    '            w.Topmost = False

    '            Exit Sub
    '        End If
    '    Next

    '    w = New wndMailboxQuota With {.Owner = Me}

    '    w.mailbox = mailbox
    '    w.Show()
    'End Sub

    'Private Sub btnMailboxShare_Click(sender As Object, e As RoutedEventArgs) Handles btnMailboxShare.Click
    '    Dim w As wndMailboxShare

    '    For Each wnd As Window In Me.OwnedWindows
    '        If GetType(wndMailboxShare) Is wnd.GetType Then
    '            w = wnd

    '            w.Show()
    '            w.Activate()

    '            If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal

    '            w.Topmost = True
    '            w.Topmost = False

    '            Exit Sub
    '        End If
    '    Next

    '    w = New wndMailboxShare With {.Owner = Me}

    '    w.currentuser = currentobject
    '    w.mailbox = mailbox
    '    w.Show()
    'End Sub


End Class
