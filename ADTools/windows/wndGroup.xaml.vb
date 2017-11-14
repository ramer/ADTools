Imports System.ComponentModel

Public Class wndGroup

    Public Property currentobject As clsDirectoryObject

    Private Sub wndComputer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.DataContext = currentobject
        ctlMember.CurrentObject = currentobject
        ctlMemberOf.CurrentObject = currentobject
        ctlManagedObj.CurrentObject = currentobject
        ctlAttributes.CurrentObject = currentobject
    End Sub

    Private Sub wndComputer_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub tabctlGroup_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tabctlGroup.SelectionChanged
        If tabctlGroup.SelectedIndex = 1 Then
            ctlMember.InitializeAsync()
        ElseIf tabctlGroup.SelectedIndex = 2 Then
            ctlMemberOf.InitializeAsync()
        ElseIf tabctlGroup.SelectedIndex = 3 Then
            ctlManagedObj.InitializeAsync()
        ElseIf tabctlGroup.SelectedIndex = 4 Then
            ctlAttributes.InitializeAsync()
        End If
    End Sub

    Private Sub hlManagedBy_Click(sender As Object, e As RoutedEventArgs) Handles hlManagedBy.Click
        ShowDirectoryObjectProperties(currentobject.managedBy, Me)
    End Sub
End Class

'Глобальная группа может быть членом другой глобальной группы, универсальной группы или локальной группы домена.
'Универсальная группа может быть членом другой универсальной группы или локальной группы домена, но не может быть членом глобальной группы.
'Локальная группа домена может быть членом только другой локальной группы домена.
'Локальную группу домена можно преобразовать в универсальную группу лишь в том случае, если эта локальная группа домена не содержит 
'    других членов локальной группы домена. Локальная группа домена не может быть членом универсальной группы.
'Глобальную группу можно преобразовать в универсальную лишь в том случае, если эта глобальная группа не входит в состав другой глобальной группы.
'    Универсальная группа не может быть членом глобальной группы.