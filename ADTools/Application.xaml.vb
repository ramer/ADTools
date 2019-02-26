Class Application

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Dim w As NavigationWindow = FindVisualParent(Of NavigationWindow)(sender)
        If w IsNot Nothing Then w.Close()
    End Sub

    Private Sub btnMinimize_Click(sender As Object, e As RoutedEventArgs)
        Dim w As NavigationWindow = FindVisualParent(Of NavigationWindow)(sender)
        If w IsNot Nothing AndAlso w.WindowState <> WindowState.Minimized Then w.WindowState = WindowState.Minimized
    End Sub

    Private Sub btnMaximize_Click(sender As Object, e As RoutedEventArgs)
        Dim w As NavigationWindow = FindVisualParent(Of NavigationWindow)(sender)
        If w IsNot Nothing Then
            If w.WindowState = WindowState.Normal Then
                w.WindowState = WindowState.Maximized
            Else
                w.WindowState = WindowState.Normal
            End If
        End If
    End Sub

    Private Sub ScrollViewer_ManipulationBoundaryFeedback(sender As Object, e As ManipulationBoundaryFeedbackEventArgs)
        e.Handled = True
    End Sub

End Class
