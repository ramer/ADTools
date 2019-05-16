Imports System.DirectoryServices.Protocols
Imports IPrompt.VisualBasic

Public Class pgCreateObject
    Public Property destinationdomain As clsDomain
    Public Property destinationcontainer As clsDirectoryObject
    Public Property copyingobject As clsDirectoryObject = Nothing

    Private Property newobjectissharedmailbox As Boolean
    Private Property newobjectdisplayname As String
    Private Property newobjectuserprincipalname As String
    Private Property newobjectuserprincipalnamedomain As String
    Private Property newobjectname As String
    Private Property newobjectsamaccountname As String

    Sub New(destinationcontainer, destinationdomain)
        InitializeComponent()

        Me.destinationcontainer = destinationcontainer
        Me.destinationdomain = destinationdomain
    End Sub

    Private Sub wndCreateObject_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        cmboDomain.ItemsSource = domains
        chbCloseWindow.IsEnabled = Not preferences.PageInterface

        cmboDomain.SelectedItem = destinationdomain
        tbContainer.Text = If(destinationcontainer IsNot Nothing, destinationcontainer.distinguishedNameFormated, "")
    End Sub

    Private Sub btnContainerBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnContainerBrowse.Click
        If cmboDomain.SelectedItem Is Nothing Then Exit Sub

        Dim domain As clsDomain = CType(cmboDomain.SelectedItem, clsDomain)
        Dim domainbrowser As New pgDomainBrowser(New clsDirectoryObject(domain.DefaultNamingContext, domain))
        AddHandler domainbrowser.Return, AddressOf domainbrowserReturn

        NavigationService.Navigate(domainbrowser)
    End Sub

    Public Sub domainbrowserReturn(snd As Object, ea As ReturnEventArgs(Of clsDirectoryObject))
        If ea.Result IsNot Nothing Then
            destinationcontainer = ea.Result
            destinationdomain = CType(cmboDomain.SelectedItem, clsDomain)
            tbContainer.Text = destinationcontainer.distinguishedNameFormated
        End If
    End Sub

    Private Async Sub cmboUserUserPrincipalName_DropDownOpened(sender As Object, e As EventArgs) Handles cmboUserUserPrincipalName.DropDownOpened
        cmboUserUserPrincipalName.ItemsSource = Await GetNextDomainUserAsync(CType(cmboDomain.SelectedValue, clsDomain))
    End Sub

    Private Sub chbUserSharedMailbox_CheckedUnchecked(sender As Object, e As RoutedEventArgs) Handles chbUserSharedMailbox.Checked, chbUserSharedMailbox.Unchecked
        tbUserObjectName.Text = If(chbUserSharedMailbox.IsChecked, "SharedMailbox_" & cmboUserUserPrincipalName.Text, tbUserDisplayname.Text)
    End Sub

    Private Async Sub cmboComputerObjectName_DropDownOpened(sender As Object, e As EventArgs) Handles cmboComputerObjectName.DropDownOpened
        cmboComputerObjectName.ItemsSource = Await GetNextDomainComputerAsync(CType(cmboDomain.SelectedValue, clsDomain))
    End Sub

    Private Async Sub btnCreate_Click(sender As Object, e As RoutedEventArgs) Handles btnCreate.Click
        Dim obj As clsDirectoryObject = Nothing

        cap.Visibility = Visibility.Visible

        If tabctlObject.SelectedIndex = 0 Then
            obj = Await CreateUser()
        ElseIf tabctlObject.SelectedIndex = 1 Then
            obj = Await CreateComputer()
        ElseIf tabctlObject.SelectedIndex = 2 Then
            obj = Await CreateGroup()
        ElseIf tabctlObject.SelectedIndex = 3 Then
            obj = Await CreateContact()
        ElseIf tabctlObject.SelectedIndex = 4 Then
            obj = Await CreateOrganizationaUnit()
        End If

        cap.Visibility = Visibility.Hidden

        If obj IsNot Nothing Then
            If preferences.PageInterface Then
                If chbOpenObject.IsChecked Then ShowDirectoryObjectProperties(obj, Window.GetWindow(Me))
            Else
                If chbOpenObject.IsChecked Then ShowDirectoryObjectProperties(obj, Window.GetWindow(Me).Owner)
                If chbCloseWindow.IsChecked Then Window.GetWindow(Me).Close()
            End If
        End If
    End Sub

    Private Async Function CreateUser() As Task(Of clsDirectoryObject)
        Dim newobject As clsDirectoryObject = Nothing

        newobjectdisplayname = tbUserDisplayname.Text
        newobjectuserprincipalname = cmboUserUserPrincipalName.Text
        newobjectuserprincipalnamedomain = cmboUserUserPrincipalNameDomain.Text
        newobjectname = tbUserObjectName.Text
        newobjectsamaccountname = tbUserSamAccountName.Text

        If destinationdomain Is Nothing Or
           destinationcontainer Is Nothing Or
           newobjectdisplayname = "" Or
           newobjectuserprincipalname = "" Or
           newobjectuserprincipalnamedomain = "" Or
           newobjectname = "" Or
        newobjectsamaccountname = "" Then
            IMsgBox(My.Resources.str_MissedRequiredFields, vbOK + vbExclamation, My.Resources.str_CreateObject)
            Return Nothing
        End If

        newobjectissharedmailbox = chbUserSharedMailbox.IsChecked

        If destinationdomain.DefaultPassword = "" Then IMsgBox(My.Resources.str_DefaultPasswordIsNotSet, vbOK + vbExclamation, My.Resources.str_CreateObject) : Return Nothing

        Await Task.Run(
            Sub()

                Try
                    Dim newobjectDN As String = String.Format("CN={0},{1}", newobjectname, destinationcontainer.distinguishedName)
                    Dim addRequest As New AddRequest(newobjectDN, "user")
                    Dim addResponse As AddResponse = DirectCast(destinationdomain.Connection.SendRequest(addRequest), AddResponse)

                    newobject = New clsDirectoryObject(newobjectDN, destinationdomain)

                    newobject.sAMAccountName = newobjectsamaccountname
                    newobject.userPrincipalName = newobjectuserprincipalname & "@" & newobjectuserprincipalnamedomain
                Catch ex As Exception
                    ThrowException(ex, "Create User")
                End Try

                Try
                    newobject.ResetPassword()
                Catch ex As Exception
                    ThrowException(ex, "Set User Password")
                End Try

                Try
                    newobject.displayName = newobjectdisplayname

                    If Not newobjectissharedmailbox Then ' user
                        newobject.givenName = If(Split(newobjectdisplayname, " ").Count >= 2, Split(newobjectdisplayname, " ")(1), "")
                        newobject.sn = If(Split(newobjectdisplayname, " ").Count >= 1, Split(newobjectdisplayname, " ")(0), "")
                        newobject.userAccountControl = ADS_UF_NORMAL_ACCOUNT
                        newobject.userMustChangePasswordNextLogon = True
                    Else  ' sharedmailbox
                        newobject.userAccountControl = ADS_UF_NORMAL_ACCOUNT + ADS_UF_ACCOUNTDISABLE
                        newobject.userMustChangePasswordNextLogon = True
                    End If

                Catch ex As Exception
                    ThrowException(ex, "Set User attributes")
                End Try

                If Not newobjectissharedmailbox Then ' user
                    Try
                        For Each group As clsDirectoryObject In destinationdomain.DefaultGroups
                            group.UpdateAttribute(DirectoryAttributeOperation.Add, "member", newobject.distinguishedName)
                        Next
                    Catch ex As Exception
                        ThrowException(ex, "Set User memberof attributes")
                    End Try
                End If

            End Sub)

        newobject.Refresh()

        Return newobject
    End Function

    Private Async Function CreateComputer() As Task(Of clsDirectoryObject)
        Dim newobject As clsDirectoryObject = Nothing

        newobjectname = cmboComputerObjectName.Text
        newobjectsamaccountname = tbComputerSamAccountName.Text

        If destinationdomain Is Nothing Or
           destinationcontainer Is Nothing Or
           newobjectname = "" Or
           newobjectsamaccountname = "" Then
            IMsgBox(My.Resources.str_MissedRequiredFields, vbOK + vbExclamation, My.Resources.str_CreateObject)
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    Dim newobjectDN As String = String.Format("CN={0},{1}", newobjectname, destinationcontainer.distinguishedName)
                    Dim addRequest As New AddRequest(newobjectDN, "computer")
                    Dim addResponse As AddResponse = DirectCast(destinationdomain.Connection.SendRequest(addRequest), AddResponse)

                    newobject = New clsDirectoryObject(newobjectDN, destinationdomain)

                    newobject.sAMAccountName = newobjectsamaccountname
                Catch ex As Exception
                    ThrowException(ex, "Create Computer")
                End Try

                Try
                    newobject.ResetPassword()
                Catch ex As Exception
                    ThrowException(ex, "Set Computer Password")
                End Try

                Try
                    newobject.userAccountControl = ADS_UF_WORKSTATION_TRUST_ACCOUNT + ADS_UF_PASSWD_NOTREQD
                Catch ex As Exception
                    ThrowException(ex, "Set Computer attributes")
                End Try

            End Sub)

        Return newobject
    End Function

    Private Async Function CreateGroup() As Task(Of clsDirectoryObject)
        Dim newobject As clsDirectoryObject = Nothing

        newobjectname = tbGroupObjectName.Text
        newobjectsamaccountname = tbGroupSamAccountName.Text

        Dim _groupscopedomainlocal As Boolean = rbGroupScopeDomainLocal.IsChecked
        Dim _groupscopeglobal As Boolean = rbGroupScopeGlobal.IsChecked
        Dim _groupscopeuniversal As Boolean = rbGroupScopeUniversal.IsChecked
        Dim _grouptypesecurity As Boolean = rbGroupTypeSecurity.IsChecked

        If destinationdomain Is Nothing Or
           destinationcontainer Is Nothing Or
           newobjectname = "" Or
           newobjectsamaccountname = "" Then
            IMsgBox(My.Resources.str_MissedRequiredFields, vbOK + vbExclamation, My.Resources.str_CreateObject)
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    Dim newobjectDN As String = String.Format("CN={0},{1}", newobjectname, destinationcontainer.distinguishedName)
                    Dim addRequest As New AddRequest(newobjectDN, "Group")
                    Dim addResponse As AddResponse = DirectCast(destinationdomain.Connection.SendRequest(addRequest), AddResponse)

                    newobject = New clsDirectoryObject(newobjectDN, destinationdomain)

                    newobject.sAMAccountName = newobjectsamaccountname
                Catch ex As Exception
                    ThrowException(ex, "Create Group")
                End Try

                Dim grouptype As Long = 0
                grouptype += If(_groupscopedomainlocal, ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP, 0)
                grouptype += If(_groupscopeglobal, ADS_GROUP_TYPE_GLOBAL_GROUP, 0)
                grouptype += If(_groupscopeuniversal, ADS_GROUP_TYPE_UNIVERSAL_GROUP, 0)
                grouptype += If(_grouptypesecurity, ADS_GROUP_TYPE_SECURITY_ENABLED, 0)

                If _groupscopedomainlocal Then ' domain local group, but unversal first
                    Try
                        newobject.groupType = ADS_GROUP_TYPE_UNIVERSAL_GROUP

                        newobject.groupType = grouptype
                    Catch ex As Exception
                        ThrowException(ex, "Set Group attributes")
                    End Try
                Else
                    Try
                        newobject.groupType = grouptype
                    Catch ex As Exception
                        ThrowException(ex, "Set Group attributes")
                    End Try
                End If

                Threading.Thread.Sleep(500)

            End Sub)

        Return newobject
    End Function

    Private Async Function CreateContact() As Task(Of clsDirectoryObject)
        Dim newobject As clsDirectoryObject = Nothing

        newobjectdisplayname = tbContactDisplayname.Text
        newobjectname = tbContactObjectName.Text

        If destinationdomain Is Nothing Or
           destinationcontainer Is Nothing Or
           newobjectdisplayname = "" Or
           newobjectname = "" Then
            IMsgBox(My.Resources.str_MissedRequiredFields, vbOK + vbExclamation, My.Resources.str_CreateObject)
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    Dim newobjectDN As String = String.Format("CN={0},{1}", newobjectname, destinationcontainer.distinguishedName)
                    Dim addRequest As New AddRequest(newobjectDN, "contact")
                    Dim addResponse As AddResponse = DirectCast(destinationdomain.Connection.SendRequest(addRequest), AddResponse)

                    newobject = New clsDirectoryObject(newobjectDN, destinationdomain)
                Catch ex As Exception
                    ThrowException(ex, "Create Contact")
                End Try

                Try
                    newobject.displayName = newobjectdisplayname
                    newobject.givenName = If(Split(newobjectdisplayname, " ").Count >= 2, Split(newobjectdisplayname, " ")(1), "")
                    newobject.sn = If(Split(newobjectdisplayname, " ").Count >= 1, Split(newobjectdisplayname, " ")(0), "")
                Catch ex As Exception
                    ThrowException(ex, "Set Contact attributes")
                End Try

            End Sub)

        newobject.Refresh()

        Return newobject
    End Function

    Private Async Function CreateOrganizationaUnit() As Task(Of clsDirectoryObject)
        Dim newobject As clsDirectoryObject = Nothing

        newobjectname = tbOrganizationalUnitObjectName.Text

        If destinationdomain Is Nothing Or
           destinationcontainer Is Nothing Or
           newobjectname = "" Then
            IMsgBox(My.Resources.str_MissedRequiredFields, vbOK + vbExclamation, My.Resources.str_CreateObject)
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    Dim newobjectDN As String = String.Format("OU={0},{1}", newobjectname, destinationcontainer.distinguishedName)
                    Dim addRequest As New AddRequest(newobjectDN, "organizationalUnit")
                    Dim addResponse As AddResponse = DirectCast(destinationdomain.Connection.SendRequest(addRequest), AddResponse)

                    newobject = New clsDirectoryObject(newobjectDN, destinationdomain)
                Catch ex As Exception
                    ThrowException(ex, "Create OrganizationalUnit")
                End Try

            End Sub)

        Return newobject
    End Function
End Class
