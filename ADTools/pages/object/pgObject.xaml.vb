Imports IPrompt.VisualBasic

Class pgObject

    Private DualPanelMode As Boolean
    Private frmNav As Frame

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(pgObject),
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
        Dim instance As pgObject = CType(d, pgObject)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Sub New(obj As clsDirectoryObject)
        InitializeComponent()
        CurrentObject = obj
    End Sub

    Private Sub pgObject_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        For Each uie In spNav.Children
            If TypeOf uie Is RadioButton Then CType(uie, RadioButton).IsChecked = False
        Next
    End Sub

    Private Sub pgObject_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        DualPanelMode = e.NewSize.Width > OBJECT_DUALPANEL_MINWIDTH
        If DualPanelMode Then
            If dpWrapper.Children.Count = 1 Then
                frmNav = New Frame With {.Style = TryFindResource("FrameWithoutNavigationUI")}
                dpWrapper.Children.Add(frmNav)
                If frmNav.Content Is Nothing Then NavigateFirst()
            End If
        Else
            If dpWrapper.Children.Count = 2 Then
                dpWrapper.Children.Remove(frmNav)
            End If
        End If
    End Sub

    Private Sub rbUserBasicInformation_Click(sender As Object, e As RoutedEventArgs) Handles rbUserBasicInformation.Click
        Navigate(New pgUserBasicInformation(CurrentObject))
    End Sub

    Private Sub rbContactBasicInformation_Click(sender As Object, e As RoutedEventArgs) Handles rbContactBasicInformation.Click
        Navigate(New pgContactBasicInformation(CurrentObject))
    End Sub

    Private Sub rbComputerBasicInformation_Click(sender As Object, e As RoutedEventArgs) Handles rbComputerBasicInformation.Click
        Navigate(New pgComputerBasicInformation(CurrentObject))
    End Sub

    Private Sub rbGroupBasicInformation_Click(sender As Object, e As RoutedEventArgs) Handles rbGroupBasicInformation.Click
        Navigate(New pgGroupBasicInformation(CurrentObject))
    End Sub

    Private Sub rbOrganizationalUnitBasicInformation_Click(sender As Object, e As RoutedEventArgs) Handles rbGroupBasicInformation.Click
        Navigate(New pgOrganizationalUnitBasicInformation(CurrentObject))
    End Sub

    Private Sub rbObject_Click(sender As Object, e As RoutedEventArgs) Handles rbObject.Click
        Navigate(New pgUserObject(CurrentObject))
    End Sub

    Private Sub rbMember_Click(sender As Object, e As RoutedEventArgs) Handles rbMember.Click
        Navigate(New pgObjectMember(CurrentObject))
    End Sub

    Private Sub rbMemberOf_Click(sender As Object, e As RoutedEventArgs) Handles rbUserMemberOf.Click, rbContactMemberOf.Click, rbComputerMemberOf.Click, rbGroupMemberOf.Click
        Navigate(New pgObjectMemberOf(CurrentObject))
    End Sub

    Private Sub rbComputerNetwork_Click(sender As Object, e As RoutedEventArgs) Handles rbComputerNetwork.Click
        Navigate(New pgComputerNetwork(CurrentObject))
    End Sub

    Private Sub rbComputerLoginEventLog_Click(sender As Object, e As RoutedEventArgs) Handles rbComputerLoginEventLog.Click
        Navigate(New pgComputerLoginEventLog(CurrentObject))
    End Sub

    Private Sub rbUserOrganization_Click(sender As Object, e As RoutedEventArgs) Handles rbUserOrganization.Click
        Navigate(New pgUserOrganization(CurrentObject))
    End Sub

    Private Sub rbGroupOrganization_Click(sender As Object, e As RoutedEventArgs) Handles rbGroupOrganization.Click
        Navigate(New pgGroupOrganization(CurrentObject))
    End Sub

    Private Sub rbUserExchange_Click(sender As Object, e As RoutedEventArgs) Handles rbUserExchange.Click
        Navigate(New pgUserExchange(CurrentObject))
    End Sub

    Private Sub rbContactExchange_Click(sender As Object, e As RoutedEventArgs) Handles rbContactExchange.Click
        Navigate(New pgContactExchange(CurrentObject))
    End Sub

    Private Sub rbAllAttributes_Click(sender As Object, e As RoutedEventArgs) Handles rbAllAttributes.Click
        Navigate(New pgObjectAllAttributes(CurrentObject))
    End Sub

    Private Sub imgPhoto_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgPhoto.MouseDown
        Navigate(New pgUserPhoto(CurrentObject))
    End Sub

    Public Sub Navigate(pg As Page)
        If DualPanelMode Then
            frmNav.Navigate(pg)
        Else
            NavigationService.Navigate(pg)
        End If
    End Sub

    Public Sub NavigateFirst()
        Dim pg As Page = Nothing
        Select Case CurrentObject.SchemaClass
            Case clsDirectoryObject.enmSchemaClass.User
                pg = New pgUserBasicInformation(CurrentObject)
            Case clsDirectoryObject.enmSchemaClass.Computer
                pg = New pgComputerBasicInformation(CurrentObject)
            Case clsDirectoryObject.enmSchemaClass.Group
                pg = New pgGroupBasicInformation(CurrentObject)
            Case clsDirectoryObject.enmSchemaClass.Contact
                pg = New pgContactBasicInformation(CurrentObject)
            Case clsDirectoryObject.enmSchemaClass.OrganizationalUnit
                pg = New pgOrganizationalUnitBasicInformation(CurrentObject)
            Case Else
                pg = New pgObjectAllAttributes(CurrentObject)
        End Select
        Navigate(pg)
    End Sub

    Private Sub btnClearPhoto_Click(sender As Object, e As RoutedEventArgs) Handles btnClearPhoto.Click
        If IMsgBox(My.Resources.wndObject_msg_AreYouSure, vbYesNo + vbQuestion, My.Resources.wndObject_msg_ClearPhoto, Window.GetWindow(Me)) = MsgBoxResult.Yes Then CurrentObject.thumbnailPhoto = Nothing
    End Sub

End Class
