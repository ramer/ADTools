Imports System.ComponentModel

Public Class wndMailboxQuota

    Public Property mailbox As clsMailbox

    Private Sub wndMailboxQuota_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        chbUseDatabaseQuotaDefaults.IsChecked = mailbox.UseDatabaseQuotaDefaults

        chbIssueWarningQuota.IsChecked = mailbox.IssueWarningQuota > 0
        chbProhibitSendQuota.IsChecked = mailbox.ProhibitSendQuota > 0
        chbProhibitSendReceiveQuota.IsChecked = mailbox.ProhibitSendReceiveQuota > 0

        tbIssueWarningQuota.Text = If(mailbox.IssueWarningQuota > 0, Int(mailbox.IssueWarningQuota / 1024 / 1024), "")
        tbProhibitSendQuota.Text = If(mailbox.ProhibitSendQuota > 0, Int(mailbox.ProhibitSendQuota / 1024 / 1024), "")
        tbProhibitSendReceiveQuota.Text = If(mailbox.ProhibitSendReceiveQuota > 0, Int(mailbox.ProhibitSendReceiveQuota / 1024 / 1024), "")

    End Sub

    Private Sub tbIssueWarningQuota_LostFocus(sender As Object, e As RoutedEventArgs) Handles tbIssueWarningQuota.LostFocus
        ValidateIssueWarningQuota()
    End Sub

    Private Sub tbProhibitSendQuota_LostFocus(sender As Object, e As RoutedEventArgs) Handles tbProhibitSendQuota.LostFocus
        ValidateProhibitSendQuota()
    End Sub

    Private Sub tbProhibitSendReceiveQuota_LostFocus(sender As Object, e As RoutedEventArgs) Handles tbProhibitSendReceiveQuota.LostFocus
        ValidateProhibitSendReceiveQuota()
    End Sub

    Private Sub chbIssueWarningQuota_Checked(sender As Object, e As RoutedEventArgs) Handles chbIssueWarningQuota.Checked
        If StringToLong(tbIssueWarningQuota.Text) = 0 Then tbIssueWarningQuota.Text = mailbox.DatabaseIssueWarningQuota / 1024 / 1024
        ValidateIssueWarningQuota()
    End Sub

    Private Sub chbProhibitSendQuota_Checked(sender As Object, e As RoutedEventArgs) Handles chbProhibitSendQuota.Checked
        If StringToLong(tbProhibitSendQuota.Text) = 0 Then tbProhibitSendQuota.Text = mailbox.DatabaseProhibitSendQuota / 1024 / 1024
        ValidateProhibitSendQuota()
    End Sub

    Private Sub chbProhibitSendReceiveQuota_Checked(sender As Object, e As RoutedEventArgs) Handles chbProhibitSendReceiveQuota.Checked
        If StringToLong(tbProhibitSendReceiveQuota.Text) = 0 Then tbProhibitSendReceiveQuota.Text = mailbox.DatabaseProhibitSendReceiveQuota / 1024 / 1024
        ValidateProhibitSendReceiveQuota()
    End Sub

    Private Sub ValidateIssueWarningQuota()
        If Int(StringToLong(tbIssueWarningQuota.Text) < 512) Then tbIssueWarningQuota.Text = 512
        If Int(StringToLong(tbIssueWarningQuota.Text) + 512) > StringToLong(tbProhibitSendQuota.Text) And chbProhibitSendQuota.IsChecked Then
            tbProhibitSendQuota.Text = Int(StringToLong(tbIssueWarningQuota.Text) + 512)
        End If
        If Int(StringToLong(tbIssueWarningQuota.Text) + 1024) > StringToLong(tbProhibitSendReceiveQuota.Text) And chbProhibitSendReceiveQuota.IsChecked Then
            tbProhibitSendReceiveQuota.Text = Int(StringToLong(tbIssueWarningQuota.Text) + 1024)
        End If
    End Sub

    Private Sub ValidateProhibitSendQuota()
        If Int(StringToLong(tbProhibitSendQuota.Text) < 1024) Then tbProhibitSendQuota.Text = 1024
        If Int(StringToLong(tbProhibitSendQuota.Text) - 512) < StringToLong(tbIssueWarningQuota.Text) And chbIssueWarningQuota.IsChecked And Int(StringToLong(tbProhibitSendQuota.Text) - 512) > 0 Then
            tbIssueWarningQuota.Text = Int(StringToLong(tbProhibitSendQuota.Text) - 512)
        End If
        If Int(StringToLong(tbProhibitSendQuota.Text) + 512) > StringToLong(tbProhibitSendReceiveQuota.Text) And chbProhibitSendReceiveQuota.IsChecked Then
            tbProhibitSendReceiveQuota.Text = Int(StringToLong(tbProhibitSendQuota.Text) + 512)
        End If
    End Sub

    Private Sub ValidateProhibitSendReceiveQuota()
        If Int(StringToLong(tbProhibitSendReceiveQuota.Text) < 1536) Then tbProhibitSendReceiveQuota.Text = 1536
        If Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 512) < StringToLong(tbProhibitSendQuota.Text) And chbProhibitSendQuota.IsChecked And Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 512) > 0 Then
            tbProhibitSendQuota.Text = Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 512)
        End If
        If Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 1024) < StringToLong(tbIssueWarningQuota.Text) And chbIssueWarningQuota.IsChecked And Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 1024) > 0 Then
            tbIssueWarningQuota.Text = Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 1024)
        End If
    End Sub

    Private Async Function SaveIssueWarningQuota() As Task
        If chbIssueWarningQuota.IsChecked Then
            Dim q = StringToLong(tbIssueWarningQuota.Text) * 1024 * 1024
            Await Task.Run(Sub() mailbox.IssueWarningQuota = q)
        Else
            Await Task.Run(Sub() mailbox.IssueWarningQuota = -1) 'unlimited
        End If
    End Function

    Private Async Function SaveProhibitSendQuota() As Task
        If chbProhibitSendQuota.IsChecked Then
            Dim q = StringToLong(tbProhibitSendQuota.Text) * 1024 * 1024
            Await Task.Run(Sub() mailbox.ProhibitSendQuota = q)
        Else
            Await Task.Run(Sub() mailbox.ProhibitSendQuota = -1) 'unlimited
        End If
    End Function

    Private Async Function SaveProhibitSendReceiveQuota() As Task
        If chbProhibitSendReceiveQuota.IsChecked Then
            Dim q = StringToLong(tbProhibitSendReceiveQuota.Text) * 1024 * 1024
            Await Task.Run(Sub() mailbox.ProhibitSendReceiveQuota = q)
        Else
            Await Task.Run(Sub() mailbox.ProhibitSendReceiveQuota = -1) 'unlimited
        End If
    End Function

    Private Sub btnDecreaseAll_Click(sender As Object, e As RoutedEventArgs) Handles btnDecreaseAll.Click
        If chbIssueWarningQuota.IsChecked Then
            tbIssueWarningQuota.Text = Int(StringToLong(tbIssueWarningQuota.Text) - 512)
            ValidateIssueWarningQuota()
        End If

        If chbProhibitSendQuota.IsChecked Then
            tbProhibitSendQuota.Text = Int(StringToLong(tbProhibitSendQuota.Text) - 512)
            ValidateProhibitSendQuota()
        End If

        If chbProhibitSendReceiveQuota.IsChecked Then
            tbProhibitSendReceiveQuota.Text = Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 512)
            ValidateProhibitSendReceiveQuota()
        End If
    End Sub

    Private Sub btnIncreaseAll_Click(sender As Object, e As RoutedEventArgs) Handles btnIncreaseAll.Click
        If chbIssueWarningQuota.IsChecked Then
            tbIssueWarningQuota.Text = Int(StringToLong(tbIssueWarningQuota.Text) + 512)
            ValidateIssueWarningQuota()
        End If

        If chbProhibitSendQuota.IsChecked Then
            tbProhibitSendQuota.Text = Int(StringToLong(tbProhibitSendQuota.Text) + 512)
            ValidateProhibitSendQuota()
        End If

        If chbProhibitSendReceiveQuota.IsChecked Then
            tbProhibitSendReceiveQuota.Text = Int(StringToLong(tbProhibitSendReceiveQuota.Text) + 512)
            ValidateProhibitSendReceiveQuota()
        End If
    End Sub

    Private Sub btnIssueWarningQuotaDecrease_Click(sender As Object, e As RoutedEventArgs) Handles btnIssueWarningQuotaDecrease.Click
        tbIssueWarningQuota.Text = Int(StringToLong(tbIssueWarningQuota.Text) - 512)
        ValidateIssueWarningQuota()
    End Sub

    Private Sub btnIssueWarningQuotaIncrease_Click(sender As Object, e As RoutedEventArgs) Handles btnIssueWarningQuotaIncrease.Click
        tbIssueWarningQuota.Text = Int(StringToLong(tbIssueWarningQuota.Text) + 512)
        ValidateIssueWarningQuota()
    End Sub

    Private Sub btnProhibitSendQuotaDecrease_Click(sender As Object, e As RoutedEventArgs) Handles btnProhibitSendQuotaDecrease.Click
        tbProhibitSendQuota.Text = Int(StringToLong(tbProhibitSendQuota.Text) - 512)
        ValidateProhibitSendQuota()
    End Sub

    Private Sub btnProhibitSendQuotaIncrease_Click(sender As Object, e As RoutedEventArgs) Handles btnProhibitSendQuotaIncrease.Click
        tbProhibitSendQuota.Text = Int(StringToLong(tbProhibitSendQuota.Text) + 512)
        ValidateProhibitSendQuota()
    End Sub

    Private Sub btnProhibitSendReceiveQuotaDecrease_Click(sender As Object, e As RoutedEventArgs) Handles btnProhibitSendReceiveQuotaDecrease.Click
        tbProhibitSendReceiveQuota.Text = Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 512)
        ValidateProhibitSendReceiveQuota()
    End Sub

    Private Sub btnProhibitSendReceiveQuotaIncrease_Click(sender As Object, e As RoutedEventArgs) Handles btnProhibitSendReceiveQuotaIncrease.Click
        tbProhibitSendReceiveQuota.Text = Int(StringToLong(tbProhibitSendReceiveQuota.Text) + 512)
        ValidateProhibitSendReceiveQuota()
    End Sub

    Private Sub wndMailboxQuota_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Async Sub btnOK_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
        cap.Visibility = Visibility.Visible

        ValidateIssueWarningQuota()
        ValidateProhibitSendQuota()
        ValidateProhibitSendQuota()

        Dim d = chbUseDatabaseQuotaDefaults.IsChecked
        Await Task.Run(Sub() mailbox.UseDatabaseQuotaDefaults = d)

        Await SaveProhibitSendReceiveQuota()
        Await SaveProhibitSendQuota()
        Await SaveIssueWarningQuota()

        cap.Visibility = Visibility.Hidden

        Me.Close()
    End Sub

    Private Function StringToLong(str) As Long
        Dim lng As Double
        If Not Double.TryParse(str, lng) Then lng = 0
        Return Int(lng)
    End Function

End Class
