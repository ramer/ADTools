Imports System.ComponentModel
Imports IPrompt.VisualBasic

Public Class wndUser
    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(wndUser),
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
        Dim instance As wndUser = CType(d, wndUser)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
            ._currentobject.RefreshAllAllowedAttributes()
        End With
    End Sub

    WithEvents searcher As New clsSearcher

    Sub New()
        InitializeComponent()
    End Sub

    Private Sub wndUser_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

    End Sub

    Private Sub wnd_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub cmboTelephoneNumber_DropDownOpened(sender As Object, e As EventArgs) Handles cmboTelephoneNumber.DropDownOpened
        cmboTelephoneNumber.ItemsSource = GetNextDomainTelephoneNumbers(currentobject.Domain)
    End Sub

    Private Sub cmboTelephoneNumber_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmboTelephoneNumber.SelectionChanged
        e.Handled = True
    End Sub

    Private Sub btnResolveDepartment_Click(sender As Object, e As RoutedEventArgs) Handles btnResolveDepartment.Click
        If currentobject Is Nothing Then Exit Sub
        'tbDepartment.Text = Replace(currentobject.Entry.Parent.Name, "OU=", "")
    End Sub

    Private Sub hlManager_Click(sender As Object, e As RoutedEventArgs) Handles hlManager.Click
        ShowDirectoryObjectProperties(currentobject.manager, Me)
    End Sub

    Private Sub btnClearPhoto_Click(sender As Object, e As RoutedEventArgs) Handles btnClearPhoto.Click
        If IMsgBox(My.Resources.wndObject_msg_AreYouSure, vbYesNo + vbQuestion, My.Resources.wndObject_msg_ClearPhoto) = MsgBoxResult.Yes Then currentobject.thumbnailPhoto = Nothing
    End Sub

    Private Sub imgPhoto_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgPhoto.MouseDown
        Dim w As wndUserPhoto

        For Each wnd As Window In Me.OwnedWindows
            If GetType(wndUserPhoto) Is wnd.GetType Then
                w = wnd

                w.Show()
                w.Activate()

                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal

                w.Topmost = True
                w.Topmost = False

                Exit Sub
            End If
        Next

        w = New wndUserPhoto With {.Owner = Me}

        w.imgPhoto.Source = imgPhoto.Source
        w.Show()
    End Sub

    Private Sub btnResetPassword_Click(sender As Object, e As RoutedEventArgs) Handles btnResetPassword.Click
        Try
            If currentobject Is Nothing Then Exit Sub
            If IMsgBox(My.Resources.wndObject_msg_AreYouSure, vbYesNo + vbQuestion, My.Resources.wndObject_msg_ResetPassword) = vbYes Then
                currentobject.ResetPassword()
                currentobject.passwordNeverExpires = False
                IMsgBox(My.Resources.wndObject_msg_PasswordChanged, vbOKOnly + vbInformation, My.Resources.wndObject_msg_ResetPassword)
            End If
        Catch ex As Exception
            ThrowException(ex, "btnResetPassword_Click")
        End Try
    End Sub

    Private Sub btnSetPassword_Click(sender As Object, e As RoutedEventArgs) Handles btnSetPassword.Click
        Try
            If currentobject Is Nothing Then Exit Sub
            Dim newpassword As String
            newpassword = IPasswordBox(My.Resources.wndObject_msg_EnterNewPassword, My.Resources.wndObject_msg_ChangePassword,, vbQuestion, Me)
            If String.IsNullOrEmpty(newpassword) Then Exit Sub
            currentobject.SetPassword(newpassword)
            IMsgBox(My.Resources.wndObject_lbl_PasswordLastSet, vbOKOnly + vbInformation, My.Resources.wndObject_msg_ChangePassword)
        Catch ex As Exception
            ThrowException(ex, "btnSetPassword_Click")
        End Try
    End Sub

End Class
