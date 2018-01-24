Imports IPrompt.VisualBasic

Class pgUserObject

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                        GetType(clsDirectoryObject),
                                                        GetType(pgUserObject),
                                                        New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgUserObject = CType(d, pgUserObject)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Sub New(obj As clsDirectoryObject)
        InitializeComponent()
        CurrentObject = obj
    End Sub

    Private Sub btnResetPassword_Click(sender As Object, e As RoutedEventArgs) Handles btnResetPassword.Click
        Try
            If CurrentObject Is Nothing Then Exit Sub
            If IMsgBox(My.Resources.str_AreYouSure, vbYesNo + vbQuestion, My.Resources.str_PasswordReset) = vbYes Then
                CurrentObject.ResetPassword()
                CurrentObject.passwordNeverExpires = False
                IMsgBox(My.Resources.str_PasswordChanged, vbOKOnly + vbInformation, My.Resources.str_PasswordReset, Window.GetWindow(Me))
            End If
        Catch ex As Exception
            ThrowException(ex, "btnResetPassword_Click")
        End Try
    End Sub

    Private Sub btnSetPassword_Click(sender As Object, e As RoutedEventArgs) Handles btnSetPassword.Click
        Try
            If CurrentObject Is Nothing Then Exit Sub
            Dim newpassword As String
            newpassword = IPasswordBox(My.Resources.str_EnterNewPassword, My.Resources.str_PasswordChange,, vbQuestion, Window.GetWindow(Me))
            If String.IsNullOrEmpty(newpassword) Then Exit Sub
            CurrentObject.SetPassword(newpassword)
            IMsgBox(My.Resources.str_PasswordLastSet, vbOKOnly + vbInformation, My.Resources.str_PasswordChange)
        Catch ex As Exception
            ThrowException(ex, "btnSetPassword_Click")
        End Try
    End Sub

End Class
