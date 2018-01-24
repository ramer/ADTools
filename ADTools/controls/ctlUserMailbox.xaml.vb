Imports System.Collections.ObjectModel
Imports IPrompt.VisualBasic

Public Class ctlUserMailbox

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlUserMailbox),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject

    Private Property mailbox As clsMailbox

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

    Private Sub ctlMailbox_Unloaded(sender As Object, e As RoutedEventArgs) Handles Me.Unloaded
        If mailbox Is Nothing Then Exit Sub
        mailbox.Close()
    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing Then Exit Sub

        If mailbox Is Nothing AndAlso _currentobject.Domain.UseExchange AndAlso _currentobject.Domain.ExchangeServer IsNot Nothing Then
            capexchange.Visibility = Visibility.Visible

            tbMailbox.Text = GetNextUserMailbox(_currentobject)

            mailbox = Await Task.Run(Function() New clsMailbox(_currentobject))

            Me.DataContext = mailbox

            If mailbox IsNot Nothing Then
                InitializeQuotas()
                InitializeShares()
            End If

            capexchange.Visibility = Visibility.Hidden
        End If
    End Sub

#Region "Adresses"

    Private Sub hlState_Click(sender As Object, e As RoutedEventArgs) Handles hlState.Click
        If mailbox IsNot Nothing AndAlso
           mailbox.ExchangeConnection IsNot Nothing AndAlso
           mailbox.ExchangeConnection.State IsNot Nothing AndAlso
           mailbox.ExchangeConnection.State.Reason IsNot Nothing Then

            IMsgBox(mailbox.ExchangeConnection.State.Reason.Message, vbExclamation + vbOKOnly, My.Resources.str_ConnectionError, Window.GetWindow(Me))
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

        If oldaddress.IsPrimary AndAlso IMsgBox(My.Resources.str_RemovingPrimaryAddress, vbYesNo + vbQuestion, My.Resources.str_RemovingPrimaryAddressTitle) = vbNo Then
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

#End Region

#Region "Share"

    Public WithEvents searcherusers As New clsSearcher

    Public Property users As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Private sourceobject As Object
    Private allowdrag As Boolean

    Private Sub InitializeShares()
        lvUsers.ItemsSource = users
    End Sub

    Private Sub tbSearchPattern_KeyDown(sender As Object, e As KeyEventArgs) Handles tbSearchPattern.KeyDown
        If e.Key = Key.Enter Then
            searcherusers.SearchAsync(
                users,
                Nothing,
                New clsFilter(CType(sender, TextBox).Text, attributesForSearchExchangePermissionTarget, New clsSearchObjectClasses(True, False, False, False, False)),
                New ObservableCollection(Of clsDomain)({_currentobject.Domain}))
        End If
    End Sub

    Private Sub lv_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles lvUsers.PreviewMouseLeftButtonDown,
                                                                                                   lvSendAs.PreviewMouseLeftButtonDown,
                                                                                                   lvFullAccess.PreviewMouseLeftButtonDown,
                                                                                                   lvSendOnBehalf.PreviewMouseLeftButtonDown
        Dim listView As ListView = TryCast(sender, ListView)
        allowdrag = e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth And e.GetPosition(sender).Y < listView.ActualHeight - SystemParameters.HorizontalScrollBarHeight
    End Sub

    Private Sub lv_MouseMove(sender As Object, e As MouseEventArgs) Handles lvUsers.MouseMove,
                                                                            lvSendAs.MouseMove,
                                                                            lvFullAccess.MouseMove,
                                                                            lvSendOnBehalf.MouseMove
        Dim listView As ListView = TryCast(sender, ListView)

        If e.LeftButton = MouseButtonState.Pressed And listView.SelectedItem IsNot Nothing And allowdrag Then
            sourceobject = listView

            Dim obj As clsDirectoryObject = CType(listView.SelectedItem, clsDirectoryObject)
            Dim dragData As New DataObject("clsDirectoryObject", obj)

            DragDrop.DoDragDrop(listView, dragData, DragDropEffects.Move)
        End If
    End Sub

    Private Sub lv_DragEnter(sender As Object, e As DragEventArgs) Handles lvUsers.DragEnter,
                                                                            lvSendAs.DragEnter,
                                                                            lvFullAccess.DragEnter,
                                                                            lvSendOnBehalf.DragEnter

        If Not e.Data.GetDataPresent("clsDirectoryObject") OrElse sender Is sourceobject Then
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Async Sub lv_Drop(sender As Object, e As DragEventArgs) Handles lvUsers.Drop,
                                                                        lvSendAs.Drop,
                                                                        lvFullAccess.Drop,
                                                                        lvSendOnBehalf.Drop

        If e.Data.GetDataPresent("clsDirectoryObject") And sender IsNot sourceobject Then

            capexchange.Visibility = Visibility.Visible

            Dim obj As clsDirectoryObject = TryCast(e.Data.GetData("clsDirectoryObject"), clsDirectoryObject)

            If sender Is lvSendAs Then
                If Not mailbox.PermissionSendAs.Contains(obj) Then
                    If lvSendAs.Items.Count = 0 AndAlso Not chbSentItemsConfigurationSendAs.IsChecked = True AndAlso IMsgBox(My.Resources.str_SendAsCopyToSentQuestion, vbYesNo + vbQuestion, My.Resources.str_SendAction) = MsgBoxResult.Yes Then
                        Await Task.Run(Sub() mailbox.SentItemsConfigurationSendAs = True)
                    End If
                    Await Task.Run(Sub() mailbox.AddPermissionSendAs(obj))
                End If
            ElseIf sender Is lvFullAccess Then
                If Not mailbox.PermissionFullAccess.Contains(obj) Then
                    Await Task.Run(Sub() mailbox.AddPermissionFullAccess(obj))
                End If
            ElseIf sender Is lvSendOnBehalf Then
                If Not mailbox.PermissionSendOnBehalf.Contains(obj) Then
                    If lvSendAs.Items.Count = 0 AndAlso Not chbSentItemsConfigurationSendOnBehalf.IsChecked = True AndAlso IMsgBox(My.Resources.str_SendOnBehalfCopyToSentQuestion, vbYesNo + vbQuestion, My.Resources.str_SendAction) = MsgBoxResult.Yes Then
                        Await Task.Run(Sub() mailbox.SentItemsConfigurationSendOnBehalf = True)
                    End If
                    Await Task.Run(Sub() mailbox.AddPermissionSendOnBehalf(obj))
                End If
            Else
                If sourceobject Is lvSendAs Then
                    Await Task.Run(Sub() mailbox.RemovePermissionSendAs(obj))
                ElseIf sourceobject Is lvFullAccess Then
                    Await Task.Run(Sub() mailbox.RemovePermissionFullAccess(obj))
                ElseIf sourceobject Is lvSendOnBehalf Then
                    Await Task.Run(Sub() mailbox.RemovePermissionSendOnBehalf(obj))
                End If
            End If

            capexchange.Visibility = Visibility.Hidden

        End If
    End Sub

    Private Sub lvUsers_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvUsers.MouseDoubleClick
        If lvUsers.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvUsers.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lvFullAccess_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvFullAccess.MouseDoubleClick
        If lvFullAccess.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvFullAccess.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lvSendAs_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvSendAs.MouseDoubleClick
        If lvSendAs.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvSendAs.SelectedItem, Window.GetWindow(Me))
    End Sub

    Private Sub lvSendOnBehalf_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvSendOnBehalf.MouseDoubleClick
        If lvSendOnBehalf.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvSendOnBehalf.SelectedItem, Window.GetWindow(Me))
    End Sub

#End Region

#Region "Quotas"

    Private Sub InitializeQuotas()
        chbUseDatabaseQuotaDefaults.IsChecked = _mailbox.UseDatabaseQuotaDefaults

        chbIssueWarningQuota.IsChecked = _mailbox.IssueWarningQuota > 0
        chbProhibitSendQuota.IsChecked = _mailbox.ProhibitSendQuota > 0
        chbProhibitSendReceiveQuota.IsChecked = _mailbox.ProhibitSendReceiveQuota > 0

        tbIssueWarningQuota.Text = If(_mailbox.IssueWarningQuota > 0, Int(_mailbox.IssueWarningQuota / 1024 / 1024), "")
        tbProhibitSendQuota.Text = If(_mailbox.ProhibitSendQuota > 0, Int(_mailbox.ProhibitSendQuota / 1024 / 1024), "")
        tbProhibitSendReceiveQuota.Text = If(_mailbox.ProhibitSendReceiveQuota > 0, Int(_mailbox.ProhibitSendReceiveQuota / 1024 / 1024), "")
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
        If StringToLong(tbIssueWarningQuota.Text) = 0 Then tbIssueWarningQuota.Text = _mailbox.DatabaseIssueWarningQuota / 1024 / 1024
        ValidateIssueWarningQuota()
    End Sub

    Private Sub chbProhibitSendQuota_Checked(sender As Object, e As RoutedEventArgs) Handles chbProhibitSendQuota.Checked
        If StringToLong(tbProhibitSendQuota.Text) = 0 Then tbProhibitSendQuota.Text = _mailbox.DatabaseProhibitSendQuota / 1024 / 1024
        ValidateProhibitSendQuota()
    End Sub

    Private Sub chbProhibitSendReceiveQuota_Checked(sender As Object, e As RoutedEventArgs) Handles chbProhibitSendReceiveQuota.Checked
        If StringToLong(tbProhibitSendReceiveQuota.Text) = 0 Then tbProhibitSendReceiveQuota.Text = _mailbox.DatabaseProhibitSendReceiveQuota / 1024 / 1024
        ValidateProhibitSendReceiveQuota()
    End Sub

    Private Sub ValidateIssueWarningQuota()
        If chbIssueWarningQuota.IsChecked = False Then Exit Sub

        If Int(StringToLong(tbIssueWarningQuota.Text) < 512) Then tbIssueWarningQuota.Text = 512
        If Int(StringToLong(tbIssueWarningQuota.Text) + 512) > StringToLong(tbProhibitSendQuota.Text) And chbProhibitSendQuota.IsChecked Then
            tbProhibitSendQuota.Text = Int(StringToLong(tbIssueWarningQuota.Text) + 512)
        End If
        If Int(StringToLong(tbIssueWarningQuota.Text) + 1024) > StringToLong(tbProhibitSendReceiveQuota.Text) And chbProhibitSendReceiveQuota.IsChecked Then
            tbProhibitSendReceiveQuota.Text = Int(StringToLong(tbIssueWarningQuota.Text) + 1024)
        End If
        btnApply.IsEnabled = True
    End Sub

    Private Sub ValidateProhibitSendQuota()
        If chbProhibitSendQuota.IsChecked = False Then Exit Sub

        If Int(StringToLong(tbProhibitSendQuota.Text) < 1024) Then tbProhibitSendQuota.Text = 1024
        If Int(StringToLong(tbProhibitSendQuota.Text) - 512) < StringToLong(tbIssueWarningQuota.Text) And chbIssueWarningQuota.IsChecked And Int(StringToLong(tbProhibitSendQuota.Text) - 512) > 0 Then
            tbIssueWarningQuota.Text = Int(StringToLong(tbProhibitSendQuota.Text) - 512)
        End If
        If Int(StringToLong(tbProhibitSendQuota.Text) + 512) > StringToLong(tbProhibitSendReceiveQuota.Text) And chbProhibitSendReceiveQuota.IsChecked Then
            tbProhibitSendReceiveQuota.Text = Int(StringToLong(tbProhibitSendQuota.Text) + 512)
        End If
        btnApply.IsEnabled = True
    End Sub

    Private Sub ValidateProhibitSendReceiveQuota()
        If chbProhibitSendReceiveQuota.IsChecked = False Then Exit Sub

        If Int(StringToLong(tbProhibitSendReceiveQuota.Text) < 1536) Then tbProhibitSendReceiveQuota.Text = 1536
        If Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 512) < StringToLong(tbProhibitSendQuota.Text) And chbProhibitSendQuota.IsChecked And Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 512) > 0 Then
            tbProhibitSendQuota.Text = Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 512)
        End If
        If Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 1024) < StringToLong(tbIssueWarningQuota.Text) And chbIssueWarningQuota.IsChecked And Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 1024) > 0 Then
            tbIssueWarningQuota.Text = Int(StringToLong(tbProhibitSendReceiveQuota.Text) - 1024)
        End If
        btnApply.IsEnabled = True
    End Sub

    Private Async Function SaveIssueWarningQuota() As Task
        If chbIssueWarningQuota.IsChecked Then
            Dim q = StringToLong(tbIssueWarningQuota.Text) * 1024 * 1024
            Await Task.Run(Sub() _mailbox.IssueWarningQuota = q)
        Else
            Await Task.Run(Sub() _mailbox.IssueWarningQuota = -1) 'unlimited
        End If
    End Function

    Private Async Function SaveProhibitSendQuota() As Task
        If chbProhibitSendQuota.IsChecked Then
            Dim q = StringToLong(tbProhibitSendQuota.Text) * 1024 * 1024
            Await Task.Run(Sub() _mailbox.ProhibitSendQuota = q)
        Else
            Await Task.Run(Sub() _mailbox.ProhibitSendQuota = -1) 'unlimited
        End If
    End Function

    Private Async Function SaveProhibitSendReceiveQuota() As Task
        If chbProhibitSendReceiveQuota.IsChecked Then
            Dim q = StringToLong(tbProhibitSendReceiveQuota.Text) * 1024 * 1024
            Await Task.Run(Sub() _mailbox.ProhibitSendReceiveQuota = q)
        Else
            Await Task.Run(Sub() _mailbox.ProhibitSendReceiveQuota = -1) 'unlimited
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

    Private Async Sub btnApply_Click(sender As Object, e As RoutedEventArgs) Handles btnApply.Click
        capexchange.Visibility = Visibility.Visible

        ValidateIssueWarningQuota()
        ValidateProhibitSendQuota()
        ValidateProhibitSendReceiveQuota()

        Dim d = chbUseDatabaseQuotaDefaults.IsChecked
        Await Task.Run(Sub() _mailbox.UseDatabaseQuotaDefaults = d)

        Await SaveProhibitSendReceiveQuota()
        Await SaveProhibitSendQuota()
        Await SaveIssueWarningQuota()

        capexchange.Visibility = Visibility.Hidden

        btnApply.IsEnabled = False
    End Sub

    Private Function StringToLong(str) As Long
        Dim lng As Double
        If Not Long.TryParse(str, lng) Then lng = 0
        Return lng
    End Function

#End Region

End Class
