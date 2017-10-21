Imports System.ComponentModel

Public Class wndCreateObject
    Public Property objectdomain As clsDomain
    Public Property objectcontainer As clsDirectoryObject

    Private Property _objectissharedmailbox As Boolean
    Private Property _objectdisplayname As String
    Private Property _objectuserprincipalname As String
    Private Property _objectuserprincipalnamedomain As String
    Private Property _objectname As String
    Private Property _objectsamaccountname As String

    Private Sub wndCreateObject_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DataContext = Me
        cmboDomain.ItemsSource = domains

        cmboDomain.SelectedItem = objectdomain
        tbContainer.Text = objectcontainer.Entry.Path

    End Sub

    Private Sub btnContainerBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnContainerBrowse.Click
        If cmboDomain.SelectedItem Is Nothing Then Exit Sub

        Dim domainbrowser As New wndDomainBrowser
        Dim domain As clsDomain = CType(cmboDomain.SelectedItem, clsDomain)

        domainbrowser.rootobject = New clsDirectoryObject(domain.DefaultNamingContext, domain)
        ShowWindow(domainbrowser, True, Me, True)

        If domainbrowser.DialogResult = True AndAlso domainbrowser.currentobject IsNot Nothing Then
            domain.SearchRoot = domainbrowser.currentobject.Entry
            objectdomain = CType(cmboDomain.SelectedItem, clsDomain)
            objectcontainer = domainbrowser.currentobject
            tbContainer.Text = objectcontainer.Entry.Path
        End If
    End Sub

    Private Sub tbUserDisplayname_TextChanged(sender As Object, e As TextChangedEventArgs) Handles tbUserDisplayname.TextChanged
        tbUserObjectName.Text = If(chbUserSharedMailbox.IsChecked, "SharedMailbox_" & cmboUserUserPrincipalName.Text, tbUserDisplayname.Text)
    End Sub

    Private Sub cmboUserUserPrincipalName_TextChanged(sender As Object, e As TextChangedEventArgs)
        tbUserObjectName.Text = If(chbUserSharedMailbox.IsChecked, "SharedMailbox_" & cmboUserUserPrincipalName.Text, tbUserDisplayname.Text)
    End Sub

    Private Sub cmboUserUserPrincipalName_DropDownOpened(sender As Object, e As EventArgs) Handles cmboUserUserPrincipalName.DropDownOpened
        cmboUserUserPrincipalName.ItemsSource = GetNextDomainUsers(CType(cmboDomain.SelectedValue, clsDomain), tbUserDisplayname.Text)
    End Sub

    Private Sub chbUserSharedMailbox_CheckedUnchecked(sender As Object, e As RoutedEventArgs) Handles chbUserSharedMailbox.Checked, chbUserSharedMailbox.Unchecked
        tbUserObjectName.Text = If(chbUserSharedMailbox.IsChecked, "SharedMailbox_" & cmboUserUserPrincipalName.Text, tbUserDisplayname.Text)
    End Sub

    Private Sub cmboComputerObjectName_DropDownOpened(sender As Object, e As EventArgs) Handles cmboComputerObjectName.DropDownOpened
        cmboComputerObjectName.ItemsSource = GetNextDomainComputers(CType(cmboDomain.SelectedValue, clsDomain))
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
            If TypeOf Me.Owner Is wndMain Then
                If obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Then CType(Me.Owner, wndMain).DomainTreeUpdate()
                CType(Me.Owner, wndMain).Refresh()
            End If
            If chbOpenObject.IsChecked Then ShowDirectoryObjectProperties(obj, Me.Owner)
            If chbCloseWindow.IsChecked = True Then Me.Close()
        End If
    End Sub

    Private Sub wndCreateObject_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Async Function CreateUser() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectdisplayname = tbUserDisplayname.Text
        _objectuserprincipalname = cmboUserUserPrincipalName.Text
        _objectuserprincipalnamedomain = cmboUserUserPrincipalNameDomain.Text
        _objectname = tbUserObjectName.Text
        _objectsamaccountname = tbUserSamAccountName.Text

        If objectdomain Is Nothing Or
           objectcontainer Is Nothing Or
           _objectdisplayname = "" Or
           _objectuserprincipalname = "" Or
           _objectuserprincipalnamedomain = "" Or
           _objectname = "" Or
        _objectsamaccountname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        _objectissharedmailbox = chbUserSharedMailbox.IsChecked

        If objectdomain.DefaultPassword = "" Then ThrowCustomException("Стандартный пароль не указан") : Return Nothing

        Await Task.Run(
            Sub()

                Try
                    Dim entry As DirectoryServices.DirectoryEntry = objectcontainer.Entry.Children.Add("CN=" & _objectname, "user")
                    entry.CommitChanges()
                    _currentobject = New clsDirectoryObject(entry, objectdomain)

                    _currentobject.sAMAccountName = _objectsamaccountname
                    _currentobject.userPrincipalName = _objectuserprincipalname & "@" & _objectuserprincipalnamedomain
                Catch ex As Exception
                    ThrowException(ex, "Create User")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.ResetPassword()
                Catch ex As Exception
                    ThrowException(ex, "Set User Password")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.displayName = _objectdisplayname

                    If Not _objectissharedmailbox Then ' user
                        _currentobject.givenName = If(Split(_objectdisplayname, " ").Count >= 2, Split(_objectdisplayname, " ")(1), "")
                        _currentobject.sn = If(Split(_objectdisplayname, " ").Count >= 1, Split(_objectdisplayname, " ")(0), "")
                        _currentobject.userAccountControl = ADS_UF_NORMAL_ACCOUNT
                        _currentobject.userMustChangePasswordNextLogon = True
                    Else                                       ' sharedmailbox
                        _currentobject.userAccountControl = ADS_UF_NORMAL_ACCOUNT + ADS_UF_ACCOUNTDISABLE
                        _currentobject.userMustChangePasswordNextLogon = True
                    End If

                Catch ex As Exception
                    ThrowException(ex, "Set User attributes")
                End Try

                Threading.Thread.Sleep(500)

                If Not _objectissharedmailbox Then ' user
                    Try
                        For Each group As clsDirectoryObject In objectdomain.DefaultGroups
                            group.Entry.Invoke("Add", _currentobject.Entry.Path)
                            group.Entry.CommitChanges()
                            Threading.Thread.Sleep(500)
                        Next
                    Catch ex As Exception
                        ThrowException(ex, "Set User memberof attributes")
                    End Try
                End If

                Threading.Thread.Sleep(500)

            End Sub)

        _currentobject.Refresh()

        Return _currentobject
    End Function

    Private Async Function CreateComputer() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectname = cmboComputerObjectName.Text
        _objectsamaccountname = tbComputerSamAccountName.Text

        If objectdomain Is Nothing Or
           objectcontainer Is Nothing Or
           _objectname = "" Or
           _objectsamaccountname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    Dim entry As DirectoryServices.DirectoryEntry = objectcontainer.Entry.Children.Add("CN=" & _objectname, "computer")
                    entry.CommitChanges()
                    _currentobject = New clsDirectoryObject(entry, objectdomain)

                    _currentobject.sAMAccountName = _objectsamaccountname
                Catch ex As Exception
                    ThrowException(ex, "Create Computer")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.ResetPassword()
                Catch ex As Exception
                    ThrowException(ex, "Set Computer Password")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.userAccountControl = ADS_UF_WORKSTATION_TRUST_ACCOUNT + ADS_UF_PASSWD_NOTREQD
                Catch ex As Exception
                    ThrowException(ex, "Set Computer attributes")
                End Try

                Threading.Thread.Sleep(500)

            End Sub)

        Return _currentobject
    End Function

    Private Async Function CreateGroup() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectname = tbGroupObjectName.Text
        _objectsamaccountname = tbGroupSamAccountName.Text

        Dim _groupscopedomainlocal As Boolean = rbGroupScopeDomainLocal.IsChecked
        Dim _groupscopeglobal As Boolean = rbGroupScopeGlobal.IsChecked
        Dim _groupscopeuniversal As Boolean = rbGroupScopeUniversal.IsChecked
        Dim _grouptypesecurity As Boolean = rbGroupTypeSecurity.IsChecked

        If objectdomain Is Nothing Or
           objectcontainer Is Nothing Or
           _objectname = "" Or
           _objectsamaccountname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    Dim entry As DirectoryServices.DirectoryEntry = objectcontainer.Entry.Children.Add("CN=" & _objectname, "Group")
                    entry.CommitChanges()
                    _currentobject = New clsDirectoryObject(entry, objectdomain)

                    _currentobject.sAMAccountName = _objectsamaccountname
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
                        _currentobject.groupType = ADS_GROUP_TYPE_UNIVERSAL_GROUP

                        _currentobject.groupType = grouptype
                    Catch ex As Exception
                        ThrowException(ex, "Set Group attributes")
                    End Try
                Else
                    Try
                        _currentobject.groupType = grouptype
                    Catch ex As Exception
                        ThrowException(ex, "Set Group attributes")
                    End Try
                End If

                Threading.Thread.Sleep(500)

            End Sub)

        Return _currentobject
    End Function

    Private Async Function CreateContact() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectdisplayname = tbContactDisplayname.Text
        _objectname = tbContactObjectName.Text

        If objectdomain Is Nothing Or
           objectcontainer Is Nothing Or
           _objectdisplayname = "" Or
           _objectname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    Dim entry As DirectoryServices.DirectoryEntry = objectcontainer.Entry.Children.Add("CN=" & _objectname, "contact")
                    entry.CommitChanges()
                    _currentobject = New clsDirectoryObject(entry, objectdomain)
                Catch ex As Exception
                    ThrowException(ex, "Create Contact")
                End Try

                Threading.Thread.Sleep(500)

                Try
                    _currentobject.displayName = _objectdisplayname
                    _currentobject.givenName = If(Split(_objectdisplayname, " ").Count >= 2, Split(_objectdisplayname, " ")(1), "")
                    _currentobject.sn = If(Split(_objectdisplayname, " ").Count >= 1, Split(_objectdisplayname, " ")(0), "")
                Catch ex As Exception
                    ThrowException(ex, "Set Contact attributes")
                End Try

                Threading.Thread.Sleep(500)

            End Sub)

        _currentobject.Refresh()

        Return _currentobject
    End Function

    Private Async Function CreateOrganizationaUnit() As Task(Of clsDirectoryObject)
        Dim _currentobject As clsDirectoryObject = Nothing

        _objectname = tbOrganizationalUnitObjectName.Text

        If objectdomain Is Nothing Or
           objectcontainer Is Nothing Or
           _objectname = "" Then
            ThrowCustomException("Не заполнены необходимые поля")
            Return Nothing
        End If

        Await Task.Run(
            Sub()

                Try
                    Dim entry As DirectoryServices.DirectoryEntry = objectcontainer.Entry.Children.Add("OU=" & _objectname, "organizationalUnit")
                    entry.CommitChanges()
                    _currentobject = New clsDirectoryObject(entry, objectdomain)
                Catch ex As Exception
                    ThrowException(ex, "Create OrganizationalUnit")
                End Try

                Threading.Thread.Sleep(500)

            End Sub)

        Return _currentobject
    End Function
End Class
