Class wndMain

    Private Sub wndMain_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles MyBase.Closing, MyBase.Closing
        Windows.Application.Current.Shutdown()
    End Sub

End Class
