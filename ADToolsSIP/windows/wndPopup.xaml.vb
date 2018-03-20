Public Class wndPopup

    Private Sub lbCalls_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles lbCalls.MouseDoubleClick
        Dim obj = FindVisualParent(Of ListBoxItem)(e.OriginalSource)
        If obj Is Nothing OrElse obj.DataContext Is Nothing Then Exit Sub
        Process.Start("..\..\ADTools.exe", "-search """ & CType(obj.DataContext, clsCall).Filter & """")
    End Sub

    Private Sub wndPopup_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.Left = Forms.Screen.PrimaryScreen.WorkingArea.Right - 5 - e.NewSize.Width
        Me.Top = Forms.Screen.PrimaryScreen.WorkingArea.Bottom - 5 - e.NewSize.Height
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub

End Class
