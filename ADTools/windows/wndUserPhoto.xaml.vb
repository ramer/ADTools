Public Class wndUserPhoto

    Private Sub imgPhoto_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgPhoto.MouseDown
        Me.Close()
    End Sub

    Private Sub wndUserPhoto_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDown
        Me.Close()
    End Sub

End Class
