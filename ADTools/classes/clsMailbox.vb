Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks

Public Class clsMailbox
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _currentobject As clsDirectoryObject
    Private exchangeconnector As clsPowerShell = Nothing

    Private _exchangeserver As PSObject
    Private _mailbox As PSObject = Nothing
    Private _mailboxaccepteddomain As Collection(Of PSObject) = Nothing
    Private _mailboxstatistics As PSObject = Nothing
    Private _mailboxdatabase As PSObject = Nothing
    Private _mailboxsentitemsconfiguration As PSObject = Nothing
    Private _mailboxadpermission As Collection(Of PSObject) = Nothing 'for SendAs
    Private _mailboxpermission As Collection(Of PSObject) = Nothing ' for FullAccess

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Sub New()

    End Sub

    Sub New(ByRef currentobject As clsDirectoryObject)
        _currentobject = currentobject

        If exchangeconnector IsNot Nothing Then Exit Sub

        Try
            exchangeconnector = New clsPowerShell(_currentobject.Domain.Username, _currentobject.Domain.Password, _currentobject.Domain.ExchangeServer & "." & _currentobject.Domain.Name)
            GetExchangeInfo()
        Catch ex As Exception
            exchangeconnector = Nothing
            ThrowException(ex, "New clsExchange")
        End Try
    End Sub

    Public Sub GetExchangeInfo()
        If Not exchangeconnector.State.State = RunspaceState.Opened Then Exit Sub

        Try
            _exchangeserver = GetExchangeServer(exchangeconnector, _currentobject.Domain.ExchangeServer & "." & _currentobject.Domain.Name)
            NotifyPropertyChanged("Version")
            NotifyPropertyChanged("VersionFormatted")
            NotifyPropertyChanged("State")
        Catch ex As Exception
            _exchangeserver = Nothing
            ThrowException(ex, "GetExchangeServer")
        End Try

        Try
            _mailboxaccepteddomain = GetAcceptedDomain(exchangeconnector)
            NotifyPropertyChanged("AcceptedDomain")
        Catch ex As Exception
            _mailboxaccepteddomain = Nothing
            ThrowException(ex, "AcceptedDomain")
        End Try

        UpdateMailboxInfo()
    End Sub

    Public Sub Close()
        If exchangeconnector IsNot Nothing Then
            exchangeconnector.Dispose()
            exchangeconnector = Nothing
            _mailbox = Nothing
        End If
    End Sub

    Private Sub UpdateMailboxInfo()
        Try
            _mailboxstatistics = Nothing
            _mailboxdatabase = Nothing
            _mailboxsentitemsconfiguration = Nothing
            _mailboxadpermission = Nothing
            _mailboxpermission = Nothing

            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("Exist")
            NotifyPropertyChanged("CurrentProhibitSendQuotaFormatted")
            NotifyPropertyChanged("CurrentProhibitSendQuota")
            NotifyPropertyChanged("IssueWarningQuota")
            NotifyPropertyChanged("ProhibitSendQuota")
            NotifyPropertyChanged("ProhibitSendReceiveQuota")
            NotifyPropertyChanged("UseDatabaseQuotaDefaults")
            NotifyPropertyChanged("EmailAddresses")
            NotifyPropertyChanged("HiddenFromAddressListsEnabled")
            NotifyPropertyChanged("Type")
            NotifyPropertyChanged("PermissionSendOnBehalf")
        Catch ex As Exception
            _mailbox = Nothing
            ThrowException(ex, "GetMailbox")
        End Try

        If _mailbox IsNot Nothing Then

            Try
                _mailboxstatistics = GetMailboxStatistics(exchangeconnector, _currentobject.objectGUID)
                NotifyPropertyChanged("Size")
                NotifyPropertyChanged("SizeFormatted")
            Catch ex As Exception
                _mailboxstatistics = Nothing
                ThrowException(ex, "GetMailboxStatistics")
            End Try

            Try
                _mailboxdatabase = GetMailboxDatabase(exchangeconnector)
                NotifyPropertyChanged("DatabaseIssueWarningQuota")
                NotifyPropertyChanged("DatabaseProhibitSendQuota")
                NotifyPropertyChanged("DatabaseProhibitSendReceiveQuota")
                NotifyPropertyChanged("CurrentProhibitSendQuotaFormatted")
                NotifyPropertyChanged("CurrentProhibitSendQuota")
            Catch ex As Exception
                _mailboxdatabase = Nothing
                ThrowException(ex, "GetMailboxDatabase")
            End Try

            Try
                _mailboxsentitemsconfiguration = GetMailboxSentItemsConfiguration(exchangeconnector, _currentobject.objectGUID)
                NotifyPropertyChanged("SentItemsConfigurationSendAs")
                NotifyPropertyChanged("SentItemsConfigurationSendOnBehalf")
            Catch ex As Exception
                _mailboxsentitemsconfiguration = Nothing
                ThrowException(ex, "GetMailboxSentItemsConfiguration")
            End Try

            Try
                _mailboxadpermission = GetMailboxADPermission(exchangeconnector, _currentobject.objectGUID)
                NotifyPropertyChanged("PermissionSendAs")
            Catch ex As Exception
                _mailboxadpermission = Nothing
                ThrowException(ex, "GetMailboxADPermission")
            End Try

            Try
                _mailboxpermission = GetMailboxPermission(exchangeconnector, _currentobject.objectGUID)
                NotifyPropertyChanged("PermissionFullAccess")
            Catch ex As Exception
                _mailboxpermission = Nothing
                ThrowException(ex, "GetMailboxPermission")
            End Try

        End If
    End Sub

    Private Function GetExchangeServer(exch As clsPowerShell, server As String) As PSObject
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-ExchangeServer")

        Dim obj As Collection(Of PSObject) = exch.Command(cmd)
        If obj Is Nothing Then Return Nothing

        For Each obj1 As PSObject In obj
            If LCase(obj1.Properties("Name").Value) = LCase(Split(server, ".")(0)) Then
                Return obj1
            End If
        Next

        Return Nothing
    End Function

    Private Function GetMailbox(exch As clsPowerShell, objectGUID As String) As PSObject
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-Mailbox")
        cmd.AddParameter("Identity", objectGUID)

        Dim obj As Collection(Of PSObject) = exch.Command(cmd)
        If obj Is Nothing OrElse (obj.Count <> 1) Then Return Nothing

        Return obj(0)
    End Function

    Private Function GetAcceptedDomain(exch As clsPowerShell) As Collection(Of PSObject)
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-AcceptedDomain")

        Dim obj As Collection(Of PSObject) = exchangeconnector.Command(cmd)
        If obj Is Nothing Then Return Nothing

        Return obj
    End Function

    Private Function GetMailboxStatistics(exch As clsPowerShell, objectGUID As String) As PSObject
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-MailboxStatistics")
        cmd.AddParameter("Identity", objectGUID)

        Dim obj As Collection(Of PSObject) = exchangeconnector.Command(cmd)
        If obj Is Nothing OrElse (obj.Count <> 1) Then Return Nothing

        Return obj(0)
    End Function

    Private Function GetMailboxDatabase(exch As clsPowerShell) As PSObject
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-MailboxDatabase")
        cmd.AddParameter("Identity", _mailbox.Properties("Database").Value)

        Dim obj As Collection(Of PSObject) = exchangeconnector.Command(cmd)
        If obj Is Nothing OrElse (obj.Count <> 1) Then Return Nothing

        Return obj(0)
    End Function

    Private Function GetMailboxSentItemsConfiguration(exch As clsPowerShell, objectGUID As String) As PSObject
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-MailboxSentItemsConfiguration")
        cmd.AddParameter("Identity", objectGUID)

        Dim obj As Collection(Of PSObject) = exch.Command(cmd)
        If obj Is Nothing OrElse (obj.Count <> 1) Then Return Nothing

        Return obj(0)
    End Function

    Public Function GetMailboxADPermission(exch As clsPowerShell, objectGUID As String) As Collection(Of PSObject)
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-User")
        cmd.AddParameter("Identity", objectGUID)
        cmd.AddCommand("Get-ADPermission")

        Dim obj1 As Collection(Of PSObject) = exch.Command(cmd) 'список прав
        If obj1 Is Nothing Then Return Nothing

        Return obj1
    End Function

    Public Function GetMailboxPermission(exch As clsPowerShell, objectGUID As String) As Collection(Of PSObject)
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-MailboxPermission")
        cmd.AddParameter("Identity", objectGUID)

        Dim obj1 As Collection(Of PSObject) = exch.Command(cmd) 'список прав
        If obj1 Is Nothing Then Return Nothing

        Return obj1
    End Function

    Public ReadOnly Property Exist As Boolean
        Get
            Return _mailbox IsNot Nothing
        End Get
    End Property

    Public ReadOnly Property State
        Get
            If exchangeconnector IsNot Nothing Then
                Return exchangeconnector.State
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property VersionFormatted As String
        Get
            If _exchangeserver Is Nothing Then Return Nothing

            Dim obj As String = _exchangeserver.Properties("AdminDisplayVersion").Value

            Return obj
        End Get
    End Property

    Public ReadOnly Property Version As Integer
        Get
            Dim rg As New Regex("\d+")
            Dim m As Match = rg.Match(VersionFormatted)
            Return Int(m.Value)
        End Get
    End Property

    Public ReadOnly Property AcceptedDomain As String()
        Get
            If _mailboxaccepteddomain Is Nothing Then Return Nothing

            Return _mailboxaccepteddomain.ToArray().Select(Function(x As PSObject) x.Properties("DomainName").Value.ToString).ToArray
        End Get
    End Property

    Public ReadOnly Property Size As Long
        Get
            If _mailboxstatistics Is Nothing Then Return 0

            Dim obj As String = _mailboxstatistics.Properties("TotalitemSize").Value

            Return GetSizeFromString(obj.ToString)
        End Get
    End Property

    Public ReadOnly Property SizeFormatted As String
        Get
            Return SizeFormat(Size)
        End Get
    End Property

    Public ReadOnly Property CurrentProhibitSendQuotaFormatted As String
        Get
            Return SizeFormat(CurrentProhibitSendQuota)
        End Get
    End Property

    Public Function GetSizeFromString(str As String) As Long
        '2.011 GB (2,159,225,856 bytes)
        Try
            Return Long.Parse(str.Split({"(", ")"}, StringSplitOptions.RemoveEmptyEntries)(1).Replace(",", "").Split({" "}, StringSplitOptions.RemoveEmptyEntries)(0))
        Catch
            Return 0
        End Try
    End Function

    Public Function SizeFormat(size As Long) As String
        Dim suf As String() = {" B", " KB", " MB", " GB", " TB", " PB", " EB"}
        If size = 0 Then
            Return "0" + suf(0)
        End If
        Dim bytes As Long = Math.Abs(size)
        Dim place As Integer = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)))
        Dim num As Double = Math.Round(bytes / Math.Pow(1024, place), 1)
        Return (Math.Sign(size) * num).ToString() + suf(place)
    End Function

    Public ReadOnly Property CurrentProhibitSendQuota As Long
        Get
            Return If(UseDatabaseQuotaDefaults, DatabaseProhibitSendQuota, If(ProhibitSendQuota <= 0, DatabaseProhibitSendQuota, ProhibitSendQuota))
        End Get
    End Property

    Public Property IssueWarningQuota As Long
        Get
            If _mailbox Is Nothing Then Return 0
            Dim obj As String = _mailbox.Properties("IssueWarningQuota").Value
            If obj = "unlimited" Then
                Return -1
            Else
                Return GetSizeFromString(obj)
            End If
        End Get
        Set(value As Long)
            Dim cmd As New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("IssueWarningQuota", If(value < 0, "unlimited", value))
            Dim obj As Object = exchangeconnector.Command(cmd)
            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("IssueWarningQuota")
        End Set
    End Property

    Public Property ProhibitSendQuota As Long
        Get
            If _mailbox Is Nothing Then Return 0
            Dim obj As String = _mailbox.Properties("ProhibitSendQuota").Value
            If obj = "unlimited" Then
                Return -1
            Else
                Return GetSizeFromString(obj)
            End If
        End Get
        Set(value As Long)
            Dim cmd As New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("ProhibitSendQuota", If(value < 0, "unlimited", value))
            Dim obj As Object = exchangeconnector.Command(cmd)
            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("CurrentProhibitSendQuotaFormatted")
            NotifyPropertyChanged("CurrentProhibitSendQuota")
            NotifyPropertyChanged("ProhibitSendQuota")
        End Set
    End Property

    Public Property ProhibitSendReceiveQuota As Long
        Get
            If _mailbox Is Nothing Then Return 0
            Dim obj As String = _mailbox.Properties("ProhibitSendReceiveQuota").Value
            If obj = "unlimited" Then
                Return -1
            Else
                Return GetSizeFromString(obj)
            End If
        End Get
        Set(value As Long)
            Dim cmd As New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("ProhibitSendReceiveQuota", If(value < 0, "unlimited", value))
            Dim obj As Object = exchangeconnector.Command(cmd)
            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("ProhibitSendReceiveQuota")
        End Set
    End Property

    Public Property UseDatabaseQuotaDefaults As Boolean
        Get
            If _mailbox Is Nothing Then Return True
            Return _mailbox.Properties("UseDatabaseQuotaDefaults").Value
        End Get
        Set(value As Boolean)
            Dim cmd As New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("UseDatabaseQuotaDefaults", value)
            Dim obj As Object = exchangeconnector.Command(cmd)
            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("UseDatabaseQuotaDefaults")
            NotifyPropertyChanged("CurrentProhibitSendReceiveQuotaFormatted")
            NotifyPropertyChanged("CurrentProhibitSendReceiveQuota")
        End Set
    End Property

    Public ReadOnly Property DatabaseIssueWarningQuota As Long
        Get
            If _mailboxdatabase Is Nothing Then Return 0
            Dim obj As String = _mailboxdatabase.Properties("IssueWarningQuota").Value
            If obj = "unlimited" Then
                Return -1
            Else
                Return GetSizeFromString(obj)
            End If
        End Get
    End Property

    Public ReadOnly Property DatabaseProhibitSendQuota As Long
        Get
            If _mailboxdatabase Is Nothing Then Return 0
            Dim obj As String = _mailboxdatabase.Properties("ProhibitSendQuota").Value
            If obj = "unlimited" Then
                Return -1
            Else
                Return GetSizeFromString(obj)
            End If
        End Get
    End Property

    Public ReadOnly Property DatabaseProhibitSendReceiveQuota As Long
        Get
            If _mailboxdatabase Is Nothing Then Return 0
            Dim obj As String = _mailboxdatabase.Properties("ProhibitSendReceiveQuota").Value
            If obj = "unlimited" Then
                Return -1
            Else
                Return GetSizeFromString(obj)
            End If
        End Get
    End Property

    Public ReadOnly Property EmailAddresses As ObservableCollection(Of clsEmailAddress)
        Get
            If _mailbox Is Nothing Then Return Nothing

            Dim obj As PSObject = _mailbox.Properties("EmailAddresses").Value
            If obj Is Nothing Then Return Nothing

            Dim _primary As String = _mailbox.Properties("PrimarySmtpAddress").Value
            If _primary Is Nothing Then Return Nothing

            Dim e As New ObservableCollection(Of clsEmailAddress)(
                CType(obj.BaseObject, ArrayList).ToArray().Select(Function(x As Object)
                                                                      Return New clsEmailAddress(
                                                                                    Replace(LCase(x.ToString), "smtp:", ""),
                                                                                    LCase(x.ToString),
                                                                                    Replace(LCase(x.ToString), "smtp:", "").Equals(LCase(_primary)))
                                                                  End Function).ToList)
            If e Is Nothing Then Return Nothing

            Return e
        End Get
    End Property

    Public Property HiddenFromAddressListsEnabled As Boolean
        Get
            If _mailbox Is Nothing Then Return False

            Dim obj As Boolean = _mailbox.Properties("HiddenFromAddressListsEnabled").Value

            Return obj
        End Get
        Set(value As Boolean)
            If _mailbox Is Nothing Then Exit Property

            Dim cmd As New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("HiddenFromAddressListsEnabled", value)

            Dim dummy = exchangeconnector.Command(cmd)

            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)

            NotifyPropertyChanged("HiddenFromAddressListsEnabled")
        End Set
    End Property

    Public Property Type As Boolean
        Get
            If _mailbox Is Nothing Then Return False

            Dim obj As String = _mailbox.Properties("RecipientTypeDetails").Value

            If obj = "SharedMailbox" Then
                Return True
            ElseIf obj = "UserMailbox" Then
                Return False
            Else
                Return False
            End If
        End Get
        Set(value As Boolean)
            If _mailbox Is Nothing Then Exit Property

            Dim cmd As New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("Type", If(value, "Shared", "Regular"))

            Dim dummy = exchangeconnector.Command(cmd)

            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)

            NotifyPropertyChanged("Type")
        End Set
    End Property


    Public Sub Add(name As String, domain As String)
        If _mailbox IsNot Nothing Then

            Dim obj As PSObject = _mailbox.Properties("EmailAddresses").Value
            If obj Is Nothing Then Exit Sub

            obj.Methods.Item("Add").Invoke("smtp:" & name & "@" & domain)

            Dim cmd As New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("EmailAddresses", obj)

            Dim dummy = exchangeconnector.Command(cmd)

            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("EmailAddresses")

        Else

            Dim cmd As New PSCommand
            cmd = New PSCommand
            cmd.AddCommand("Enable-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("Alias", name)

            Dim dummy = exchangeconnector.Command(cmd)

            cmd = New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("Customattribute1", domain)

            Dim dummy2 = exchangeconnector.Command(cmd)

            For I As Integer = 0 To 30
                Threading.Thread.Sleep(1000)

                cmd = New PSCommand
                cmd.AddCommand("Get-Mailbox")
                cmd.AddParameter("Identity", _currentobject.objectGUID)

                Dim obj As Collection(Of PSObject) = exchangeconnector.Command(cmd)
                If obj IsNot Nothing AndAlso (obj.Count = 1) Then Exit For
            Next

            UpdateMailboxInfo()

        End If
    End Sub

    Public Sub Edit(newname As String, newdomain As String, oldaddress As clsEmailAddress)
        If _mailbox Is Nothing Then Exit Sub

        Dim cmd As New PSCommand

        If oldaddress.IsPrimary Then

            cmd = New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("Alias", newname)

            Dim dummy = exchangeconnector.Command(cmd)

            cmd = New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("Customattribute1", newdomain)

            Dim dummy2 = exchangeconnector.Command(cmd)

            Dim obj As PSObject = _mailbox.Properties("EmailAddresses").Value
            If obj Is Nothing Then Exit Sub

            For I As Integer = 0 To CType(obj.BaseObject, ArrayList).Count - 1
                If LCase(CType(obj.BaseObject, ArrayList)(I)) = LCase(oldaddress.AddressFull) Then
                    obj.Methods.Item("Remove").Invoke(CType(obj.BaseObject, ArrayList)(I))
                    Exit For
                End If
            Next
            obj.Methods.Item("Add").Invoke("smtp:" & newname & "@" & newdomain)

            cmd = New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("EmailAddresses", obj)

            Dim dummy3 = exchangeconnector.Command(cmd)

            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("EmailAddresses")
        Else

            Dim obj As PSObject = _mailbox.Properties("EmailAddresses").Value
            If obj Is Nothing Then Exit Sub

            For I As Integer = 0 To CType(obj.BaseObject, ArrayList).Count - 1
                If LCase(CType(obj.BaseObject, ArrayList)(I)) = LCase(oldaddress.AddressFull) Then
                    obj.Methods.Item("Remove").Invoke(CType(obj.BaseObject, ArrayList)(I))
                    Exit For
                End If
            Next
            obj.Methods.Item("Add").Invoke("smtp:" & newname & "@" & newdomain)

            cmd = New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("EmailAddresses", obj)

            Dim dummy = exchangeconnector.Command(cmd)

            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)

            NotifyPropertyChanged("EmailAddresses")
        End If
    End Sub

    Public Sub Remove(oldaddress As clsEmailAddress)
        If _mailbox Is Nothing Then Exit Sub

        Dim cmd As New PSCommand

        If oldaddress.IsPrimary Then

            cmd = New PSCommand
            cmd.AddCommand("Disable-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("Confirm", False)

            Dim dummy = exchangeconnector.Command(cmd)

            For I As Integer = 0 To 30
                Threading.Thread.Sleep(1000)

                cmd = New PSCommand
                cmd.AddCommand("Get-Mailbox")
                cmd.AddParameter("Identity", _currentobject.objectGUID)

                Dim obj As Collection(Of PSObject) = exchangeconnector.Command(cmd)
                If obj Is Nothing OrElse (obj.Count <> 1) Then Exit For
            Next

            UpdateMailboxInfo()

        Else

            Dim obj As PSObject = _mailbox.Properties("EmailAddresses").Value
            If obj Is Nothing Then Exit Sub

            For I As Integer = 0 To CType(obj.BaseObject, ArrayList).Count - 1
                If LCase(CType(obj.BaseObject, ArrayList)(I)) = LCase(oldaddress.AddressFull) Then
                    obj.Methods.Item("Remove").Invoke(CType(obj.BaseObject, ArrayList)(I))
                    Exit For
                End If
            Next

            cmd = New PSCommand
            cmd.AddCommand("Set-Mailbox")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("EmailAddresses", obj)

            Dim dummy = exchangeconnector.Command(cmd)

            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("EmailAddresses")
        End If
    End Sub

    Public Sub SetPrimary(mail As clsEmailAddress)
        If _mailbox Is Nothing Then Exit Sub

        Dim cmd As New PSCommand

        Dim _primary As String = _mailbox.Properties("PrimarySmtpAddress").Value
        If _primary Is Nothing Then Exit Sub

        If mail.IsPrimary Then

            ThrowCustomException("Выбранный адрес уже используется как основной")

        Else

            Dim a As String()

            Dim oldname As String = Nothing
            Dim olddomain As String = Nothing
            a = LCase(_primary).Split({"@"}, StringSplitOptions.RemoveEmptyEntries)
            If a.Count >= 2 Then
                oldname = a(0)
                olddomain = a(1)
            End If

            Dim newname As String = Nothing
            Dim newdomain As String = Nothing
            a = LCase(mail.Address).Split({"@"}, StringSplitOptions.RemoveEmptyEntries)
            If a.Count >= 2 Then
                newname = a(0)
                newdomain = a(1)
            End If

            If oldname <> newname Then
                cmd = New PSCommand
                cmd.AddCommand("Set-Mailbox")
                cmd.AddParameter("Identity", _currentobject.objectGUID)
                cmd.AddParameter("Alias", newname)

                Dim dummy = exchangeconnector.Command(cmd)
            End If

            If olddomain <> newdomain Then
                cmd = New PSCommand
                cmd.AddCommand("Set-Mailbox")
                cmd.AddParameter("Identity", _currentobject.objectGUID)
                cmd.AddParameter("Customattribute1", newdomain)

                Dim dummy = exchangeconnector.Command(cmd)
            End If

            _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("EmailAddresses")
        End If
    End Sub


    Public Property SentItemsConfigurationSendAs() As Boolean
        Get
            If _mailboxsentitemsconfiguration Is Nothing Then Return Nothing

            Dim obj As String = _mailboxsentitemsconfiguration.Properties("SendAsItemsCopiedTo").Value

            If obj = "SenderAndFrom" Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(value As Boolean)
            Dim cmd As New PSCommand
            cmd.AddCommand("Set-MailboxSentItemsConfiguration")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("SendAsItemsCopiedTo", If(value, "SenderAndFrom", "Sender"))

            Dim dummy = exchangeconnector.Command(cmd)

            _mailboxsentitemsconfiguration = GetMailboxSentItemsConfiguration(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("SentItemsConfigurationSendAs")
        End Set
    End Property

    Public Property SentItemsConfigurationSendOnBehalf() As Boolean
        Get
            If _mailboxsentitemsconfiguration Is Nothing Then Return Nothing

            Dim obj As String = _mailboxsentitemsconfiguration.Properties("SendOnBehalfOfItemsCopiedTo").Value

            If obj = "SenderAndFrom" Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(value As Boolean)
            Dim cmd As New PSCommand
            cmd.AddCommand("Set-MailboxSentItemsConfiguration")
            cmd.AddParameter("Identity", _currentobject.objectGUID)
            cmd.AddParameter("SendOnBehalfOfItemsCopiedTo", If(value, "SenderAndFrom", "Sender"))

            Dim dummy = exchangeconnector.Command(cmd)

            _mailboxsentitemsconfiguration = GetMailboxSentItemsConfiguration(exchangeconnector, _currentobject.objectGUID)
            NotifyPropertyChanged("SentItemsConfigurationSendOnBehalf")
        End Set
    End Property

    Public ReadOnly Property PermissionSendAs() As ObservableCollection(Of clsDirectoryObject)
        Get
            Dim searcher As New clsSearcher

            If _mailboxadpermission Is Nothing Then Return Nothing

            Dim rights As String() = _mailboxadpermission.Where(Function(x As PSObject) x.Properties("ExtendedRights").Value IsNot Nothing AndAlso
                                                                                        x.Properties("ExtendedRights").Value.ToString = "Send-As" AndAlso
                                                                                        x.Properties("User").Value IsNot Nothing AndAlso
                                                                                        Not UCase(x.Properties("User").Value.ToString) = "NT AUTHORITY\SELF"
                                                        ).Select(Function(x As PSObject) GetUserPartFromExchangeUsername(x.Properties("User").Value.ToString
                                                       )).ToArray

            If rights.Count = 0 Then Return New ObservableCollection(Of clsDirectoryObject)

            Return searcher.BasicSearchSync(
                    New clsDirectoryObject(_currentobject.Domain.DefaultNamingContext, _currentobject.Domain),
                    New clsFilter(Join(rights, "/"), attributesForSearchExchangePermissionSendAs, New clsSearchObjectClasses(True, True, True, True, False), False))
        End Get
    End Property

    Public Sub AddPermissionSendAs(user As clsDirectoryObject)
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-User")
        cmd.AddParameter("Identity", _currentobject.objectGUID)
        cmd.AddCommand("Add-ADPermission")
        cmd.AddParameter("User", user.userPrincipalNameName)
        cmd.AddParameter("Extendedrights", "Send-As")
        cmd.AddParameter("Confirm", False)

        Dim dummy = exchangeconnector.Command(cmd)

        _mailboxadpermission = GetMailboxADPermission(exchangeconnector, _currentobject.objectGUID)
        NotifyPropertyChanged("PermissionSendAs")
    End Sub

    Public Sub RemovePermissionSendAs(user As clsDirectoryObject)
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-User")
        cmd.AddParameter("Identity", _currentobject.objectGUID)
        cmd.AddCommand("Remove-ADPermission")
        cmd.AddParameter("User", user.userPrincipalNameName)
        cmd.AddParameter("Extendedrights", "Send-As")
        cmd.AddParameter("Confirm", False)

        Dim dummy = exchangeconnector.Command(cmd)

        _mailboxadpermission = GetMailboxADPermission(exchangeconnector, _currentobject.objectGUID)
        NotifyPropertyChanged("PermissionSendAs")
    End Sub


    Public ReadOnly Property PermissionFullAccess() As ObservableCollection(Of clsDirectoryObject)
        Get
            Dim searcher As New clsSearcher

            If _mailboxpermission Is Nothing Then Return Nothing

            Dim rights As String() = _mailboxpermission.Where(Function(x As PSObject) x.Properties("AccessRights").Value IsNot Nothing AndAlso
                                                                                      x.Properties("AccessRights").Value.ToString.Contains("FullAccess") AndAlso
                                                                                      x.Properties("IsInherited").Value = False AndAlso
                                                                                      x.Properties("User").Value IsNot Nothing AndAlso
                                                                                      Not UCase(x.Properties("User").Value.ToString) = "NT AUTHORITY\SELF"
                                                       ).Select(Function(x As PSObject) GetUserPartFromExchangeUsername(x.Properties("User").Value.ToString
                                                      )).ToArray

            If rights.Count = 0 Then Return New ObservableCollection(Of clsDirectoryObject)

            Return searcher.BasicSearchSync(
                    New clsDirectoryObject(_currentobject.Domain.DefaultNamingContext, _currentobject.Domain),
                    New clsFilter(Join(rights, "/"), attributesForSearchExchangePermissionFullAccess, New clsSearchObjectClasses(True, True, True, True, False), False))
        End Get
    End Property

    Public Sub AddPermissionFullAccess(user As clsDirectoryObject)
        Dim cmd As New PSCommand
        cmd.AddCommand("Add-MailboxPermission")
        cmd.AddParameter("Identity", _currentobject.objectGUID)
        cmd.AddParameter("User", user.userPrincipalNameName)
        cmd.AddParameter("AccessRights", "FullAccess")
        cmd.AddParameter("InheritanceType", "All")
        cmd.AddParameter("Confirm", False)

        Dim dummy = exchangeconnector.Command(cmd)

        _mailboxpermission = GetMailboxPermission(exchangeconnector, _currentobject.objectGUID)
        NotifyPropertyChanged("PermissionFullAccess")
    End Sub

    Public Sub RemovePermissionFullAccess(user As clsDirectoryObject)
        Dim cmd As New PSCommand
        cmd.AddCommand("Remove-MailboxPermission")
        cmd.AddParameter("Identity", _currentobject.objectGUID)
        cmd.AddParameter("User", user.userPrincipalNameName)
        cmd.AddParameter("AccessRights", "FullAccess")
        cmd.AddParameter("InheritanceType", "All")
        cmd.AddParameter("Confirm", False)

        Dim dummy = exchangeconnector.Command(cmd)

        _mailboxpermission = GetMailboxPermission(exchangeconnector, _currentobject.objectGUID)
        NotifyPropertyChanged("PermissionFullAccess")
    End Sub

    Public ReadOnly Property PermissionSendOnBehalf() As ObservableCollection(Of clsDirectoryObject)
        Get
            Dim searcher As New clsSearcher

            If _mailbox Is Nothing Then Return Nothing

            Dim obj1 As PSObject = _mailbox.Properties("GrantSendOnBehalfTo").Value 'список прав
            If obj1 Is Nothing Then Return Nothing

            Dim rights As String() = CType(obj1.BaseObject, ArrayList).ToArray().Select(Function(x) GetUserPartFromExchangeUsername(x.ToString)).ToArray

            If rights.Count = 0 Then Return New ObservableCollection(Of clsDirectoryObject)

            Return searcher.BasicSearchSync(
                    New clsDirectoryObject(_currentobject.Domain.DefaultNamingContext, _currentobject.Domain),
                    New clsFilter(Join(rights, "/"), attributesForSearchExchangePermissionSendOnBehalf, New clsSearchObjectClasses(True, True, True, True, False), False))
        End Get
    End Property

    Public Sub AddPermissionSendOnBehalf(user As clsDirectoryObject)
        If _mailbox Is Nothing Then Exit Sub

        Dim obj As PSObject = _mailbox.Properties("GrantSendOnBehalfTo").Value
        If obj Is Nothing Then Exit Sub

        obj.Methods.Item("Add").Invoke(user.userPrincipalNameName)

        Dim cmd As New PSCommand
        cmd = New PSCommand
        cmd.AddCommand("Set-Mailbox")
        cmd.AddParameter("Identity", _currentobject.objectGUID)
        cmd.AddParameter("GrantSendOnBehalfTo", obj)

        Dim dummy = exchangeconnector.Command(cmd)

        _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
    End Sub

    Public Sub RemovePermissionSendOnBehalf(user As clsDirectoryObject)
        If _mailbox Is Nothing Then Exit Sub

        Dim obj As PSObject = _mailbox.Properties("GrantSendOnBehalfTo").Value
        If obj Is Nothing Then Exit Sub

        For I As Integer = 0 To CType(obj.BaseObject, ArrayList).Count - 1
            If CType(obj.BaseObject, ArrayList)(I).ToString.Contains(user.name) Then
                obj.Methods.Item("Remove").Invoke(CType(obj.BaseObject, ArrayList)(I))
                Exit For
            End If
        Next

        Dim cmd As New PSCommand
        cmd = New PSCommand
        cmd.AddCommand("Set-Mailbox")
        cmd.AddParameter("Identity", _currentobject.objectGUID)
        cmd.AddParameter("GrantSendOnBehalfTo", obj)

        Dim dummy = exchangeconnector.Command(cmd)

        _mailbox = GetMailbox(exchangeconnector, _currentobject.objectGUID)
        NotifyPropertyChanged("PermissionSendOnBehalf")
    End Sub

    Public Function GetUserPartFromExchangeUsername(str) As String
        Dim arr As String() = str.Split({"\", "/"}, StringSplitOptions.RemoveEmptyEntries)
        Return If(arr.Length > 1, arr(arr.Length - 1), Nothing)
    End Function

End Class
