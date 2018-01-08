Imports IPrompt.VisualBasic

Public Class pgUser
    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(pgUser),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomainobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgUser = CType(d, pgUser)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    WithEvents searcher As New clsSearcher

    Private Sub cmboTelephoneNumber_DropDownOpened(sender As Object, e As EventArgs) Handles cmboTelephoneNumber.DropDownOpened
        cmboTelephoneNumber.ItemsSource = GetNextDomainTelephoneNumbers(CurrentObject.Domain)
    End Sub

    Private Sub cmboTelephoneNumber_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmboTelephoneNumber.SelectionChanged
        e.Handled = True
    End Sub

    Private Sub btnResolveDepartment_Click(sender As Object, e As RoutedEventArgs) Handles btnResolveDepartment.Click
        If CurrentObject Is Nothing Then Exit Sub
        tbDepartment.Text = Replace(CurrentObject.Parent.name, "OU=", "")
    End Sub

    Private Sub hlManager_Click(sender As Object, e As RoutedEventArgs) Handles hlManager.Click
        ShowDirectoryObjectProperties(CurrentObject.manager, Window.GetWindow(Me))
    End Sub

    Private Sub btnClearPhoto_Click(sender As Object, e As RoutedEventArgs) Handles btnClearPhoto.Click
        If IMsgBox(My.Resources.wndObject_msg_AreYouSure, vbYesNo + vbQuestion, My.Resources.wndObject_msg_ClearPhoto, Window.GetWindow(Me)) = MsgBoxResult.Yes Then CurrentObject.thumbnailPhoto = Nothing
    End Sub

    Private Sub imgPhoto_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgPhoto.MouseDown
        Dim w As New pgUserPhoto
        w.CurrentObject = CurrentObject
        'TODO ShowPage(w, True, Me, False)
    End Sub

    Private Sub btnResetPassword_Click(sender As Object, e As RoutedEventArgs) Handles btnResetPassword.Click
        Try
            If CurrentObject Is Nothing Then Exit Sub
            If IMsgBox(My.Resources.wndObject_msg_AreYouSure, vbYesNo + vbQuestion, My.Resources.wndObject_msg_ResetPassword) = vbYes Then
                CurrentObject.ResetPassword()
                CurrentObject.passwordNeverExpires = False
                IMsgBox(My.Resources.wndObject_msg_PasswordChanged, vbOKOnly + vbInformation, My.Resources.wndObject_msg_ResetPassword, Window.GetWindow(Me))
            End If
        Catch ex As Exception
            ThrowException(ex, "btnResetPassword_Click")
        End Try
    End Sub

    Private Sub btnSetPassword_Click(sender As Object, e As RoutedEventArgs) Handles btnSetPassword.Click
        Try
            If CurrentObject Is Nothing Then Exit Sub
            Dim newpassword As String
            newpassword = IPasswordBox(My.Resources.wndObject_msg_EnterNewPassword, My.Resources.wndObject_msg_ChangePassword,, vbQuestion, Window.GetWindow(Me))
            If String.IsNullOrEmpty(newpassword) Then Exit Sub
            CurrentObject.SetPassword(newpassword)
            IMsgBox(My.Resources.wndObject_lbl_PasswordLastSet, vbOKOnly + vbInformation, My.Resources.wndObject_msg_ChangePassword)
        Catch ex As Exception
            ThrowException(ex, "btnSetPassword_Click")
        End Try
    End Sub

    Private Sub pgUser_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
    End Sub

End Class