
Class wndMain

    Private Sub wndMain_Closing() Handles MyBase.Closing
        Dim count As Integer = 0

        For Each wnd As Window In ADToolsApplication.Current.Windows
            If GetType(wndMain) Is wnd.GetType Then count += 1
        Next

        If preferences.CloseOnXButton AndAlso count <= 1 Then ApplicationDeactivate()
    End Sub

End Class
