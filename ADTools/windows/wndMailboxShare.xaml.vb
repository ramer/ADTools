Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports IPrompt.VisualBasic

Public Class wndMailboxShare

    Public WithEvents searcherusers As New clsSearcher

    Public Property users As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Private sourceobject As Object
    Private allowdrag As Boolean

    Public Property currentuser As clsDirectoryObject
    Public Property mailbox As clsMailbox

    Private Sub wndMailboxShare_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        lvSendAs.DataContext = mailbox
        lvFullAccess.DataContext = mailbox
        lvSendOnBehalf.DataContext = mailbox
        chbSentItemsConfigurationSendAs.DataContext = mailbox
        chbSentItemsConfigurationSendOnBehalf.DataContext = mailbox
        tbSearchPattern.Focus()
    End Sub

    Private Async Sub tbSearchPattern_KeyDown(sender As Object, e As KeyEventArgs) Handles tbSearchPattern.KeyDown
        If e.Key = Key.Enter Then
            Await searcherusers.BasicSearchAsync(users, Nothing, New clsFilter(CType(sender, TextBox).Text, attributesForSearchExchangePermissionTarget, New clsSearchObjectClasses(True, False, False, False, False)), New ObservableCollection(Of clsDomain)({currentuser.Domain}))
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

            cap.Visibility = Visibility.Visible

            Dim obj As clsDirectoryObject = TryCast(e.Data.GetData("clsDirectoryObject"), clsDirectoryObject)

            If sender Is lvSendAs Then
                If Not mailbox.PermissionSendAs.Contains(obj) Then
                    If lvSendAs.Items.Count = 0 AndAlso Not chbSentItemsConfigurationSendAs.IsChecked = True AndAlso IMsgBox(My.Resources.wndMailboxShare_msg_SendAsCopyToSent, vbYesNo + vbQuestion, My.Resources.wndMailboxShare_msg_SendAction) = MsgBoxResult.Yes Then
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
                    If lvSendAs.Items.Count = 0 AndAlso Not chbSentItemsConfigurationSendOnBehalf.IsChecked = True AndAlso IMsgBox(My.Resources.wndMailboxShare_msg_SendOnBehalfCopyToSent, vbYesNo + vbQuestion, My.Resources.wndMailboxShare_msg_SendAction) = MsgBoxResult.Yes Then
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

            cap.Visibility = Visibility.Hidden

        End If
    End Sub

    Private Sub wndMailboxShare_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub lvUsers_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvUsers.MouseDoubleClick
        If lvUsers.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvUsers.SelectedItem, Me)
    End Sub

    Private Sub lvFullAccess_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvFullAccess.MouseDoubleClick
        If lvFullAccess.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvFullAccess.SelectedItem, Me)
    End Sub

    Private Sub lvSendAs_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvSendAs.MouseDoubleClick
        If lvSendAs.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvSendAs.SelectedItem, Me)
    End Sub

    Private Sub lvSendOnBehalf_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvSendOnBehalf.MouseDoubleClick
        If lvSendOnBehalf.SelectedItem Is Nothing Then Exit Sub
        ShowDirectoryObjectProperties(lvSendOnBehalf.SelectedItem, Me)
    End Sub

End Class
