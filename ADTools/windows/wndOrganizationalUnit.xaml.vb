Imports System.ComponentModel

Public Class wndOrganizationalUnit

    Public Property currentobject As clsDirectoryObject

    Private Sub wndComputer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.DataContext = currentobject
        ctlAttributes.CurrentObject = currentobject
    End Sub

    Private Sub wndComputer_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub tabctlComputer_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tabctlOrganizationalUnit.SelectionChanged
        If tabctlOrganizationalUnit.SelectedIndex = 1 Then
            ctlAttributes.InitializeAsync()
        End If
    End Sub

End Class
