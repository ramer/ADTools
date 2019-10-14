Class Application

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Dim w = Window.GetWindow(sender)
        If w IsNot Nothing Then w.Close()
    End Sub

    Private Sub btnMinimize_Click(sender As Object, e As RoutedEventArgs)
        Dim w = Window.GetWindow(sender)
        If w IsNot Nothing AndAlso w.WindowState <> WindowState.Minimized Then w.WindowState = WindowState.Minimized
    End Sub

    Private Sub btnMaximize_Click(sender As Object, e As RoutedEventArgs)
        Dim w = Window.GetWindow(sender)
        If w IsNot Nothing Then w.WindowState = If(w.WindowState = WindowState.Normal, WindowState.Maximized, WindowState.Normal)
    End Sub

    Private Sub ScrollViewer_ManipulationBoundaryFeedback(sender As Object, e As ManipulationBoundaryFeedbackEventArgs)
        e.Handled = True
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Dim wnd As Window = sender
        AddHandler wnd.Closed, Sub(s As Window, evt As EventArgs) If s.Owner IsNot Nothing Then s.Owner.Activate()
    End Sub

    Private Sub DragDropHelper_PreviewDragEnter(sender As Object, e As DragEventArgs)
        DragDropHelper.SetIsDragOver(sender, True)
    End Sub

    Private Sub DragDropHelper_PreviewDragLeave(sender As Object, e As DragEventArgs)
        DragDropHelper.SetIsDragOver(sender, False)
    End Sub

End Class
