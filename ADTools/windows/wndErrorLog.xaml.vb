Public Class wndErrorLog

    Private Sub wndErrorLog_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        lvErrorLog.ItemsSource = ADToolsApplication.tsocErrorLog
    End Sub

    Private Sub ctxmnuErrorCopy_Click() Handles ctxmnuErrorCopy.Click
        If lvErrorLog.SelectedItems.Count = 0 Then Exit Sub

        Dim msg As String = ""
        For Each log As clsErrorLog In lvErrorLog.SelectedItems
            msg &= log.TimeStamp & vbTab & log.Command & vbTab & log.Obj & vbTab & log.Message & vbTab & log.StackTrace & vbCrLf
        Next
        Clipboard.SetText(msg)
    End Sub

End Class
