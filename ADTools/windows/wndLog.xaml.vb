
Public Class wndLog

    Private Sub wndLog_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dgLog.ItemsSource = ADToolsApplication.tsocLog
    End Sub

    Private Sub ctxmnuErrorCopy_Click() Handles ctxmnuErrorCopy.Click
        If dgLog.SelectedItems.Count = 0 Then Exit Sub

        Dim msg As String = ""
        For Each log As clsLog In dgLog.SelectedItems
            msg &= log.TimeStamp & vbTab & log.Message & vbCrLf
        Next
        Clipboard.SetText(msg)
    End Sub

End Class
