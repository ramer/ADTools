
Public Class wndDomainBrowser

    Public Property currentobject As clsDirectoryObject
    Public Property rootobject As clsDirectoryObject

    Private Async Sub wndDomainBrowser_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If rootobject Is Nothing Then Exit Sub
        cap.Visibility = Visibility.Visible

        tvObjects.ItemsSource = Await Task.Run(Function() {rootobject})

        cap.Visibility = Visibility.Hidden
    End Sub

    Private Sub tviDomains_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)

        If TypeOf sp.Tag Is clsDirectoryObject Then
            currentobject = (CType(sp.Tag, clsDirectoryObject))
            tbCurrentObject.Text = currentobject.distinguishedName
        End If
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
        DialogResult = True
    End Sub
End Class
