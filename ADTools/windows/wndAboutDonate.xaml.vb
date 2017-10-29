Public Class wndAboutDonate

    Public WithEvents closeTimer As New Threading.DispatcherTimer()
    Private counter As Integer

    Private Sub wndAboutDonate_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        closeTimer.Interval = New TimeSpan(0, 0, 1)
        closeTimer.Start()
        counter = 7
    End Sub

    Private Sub imgDonate_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgDonate.MouseDown
        Donate()
    End Sub

    Private Sub closeTimer_Tick(sender As Object, e As EventArgs) Handles closeTimer.Tick
        counter -= 1
        If counter <= 0 Then Application.Current.Shutdown()
    End Sub

End Class
