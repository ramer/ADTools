Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text.RegularExpressions
Imports IPrompt.VisualBasic

Public Class clsContact
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _currentobject As clsDirectoryObject
    Private _exchangeconnection As clsPowerShell = Nothing

    Private _exchangeserver As PSObject
    Private _contact As PSObject = Nothing
    Private _contactaccepteddomain As Collection(Of PSObject) = Nothing

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

        If _exchangeconnection IsNot Nothing Then Exit Sub

        _exchangeconnection = New clsPowerShell(_currentobject.Domain.Username, _currentobject.Domain.Password, _currentobject.Domain.ExchangeServer & "." & _currentobject.Domain.Name)
        NotifyPropertyChanged("State")
        NotifyPropertyChanged("ExchangeConnection")
        NotifyPropertyChanged("Connected")

        GetExchangeInfo()
    End Sub

    Public Sub GetExchangeInfo()
        If Not _exchangeconnection.State.State = RunspaceState.Opened Then Exit Sub

        Try
            _exchangeserver = GetExchangeServer(_exchangeconnection, _currentobject.Domain.ExchangeServer & "." & _currentobject.Domain.Name)
            NotifyPropertyChanged("Version")
            NotifyPropertyChanged("VersionFormatted")
        Catch ex As Exception
            _exchangeserver = Nothing
            ThrowException(ex, "GetExchangeServer")
        End Try

        Try
            _contactaccepteddomain = GetAcceptedDomain(_exchangeconnection)
            NotifyPropertyChanged("AcceptedDomain")
        Catch ex As Exception
            _contactaccepteddomain = Nothing
            ThrowException(ex, "AcceptedDomain")
        End Try

        UpdateContactInfo()
    End Sub

    Public Sub Close()
        If _exchangeconnection IsNot Nothing Then
            _exchangeconnection.Close()
            _exchangeconnection = Nothing
            _contact = Nothing
        End If
    End Sub

    Private Sub UpdateContactInfo()
        Try
            _contact = GetContact(_exchangeconnection, _currentobject.objectGUIDFormated)
            NotifyPropertyChanged("Exist")
            NotifyPropertyChanged("EmailAddresses")
            NotifyPropertyChanged("HiddenFromAddressListsEnabled")
        Catch ex As Exception
            _contact = Nothing
            ThrowException(ex, "GetContact")
        End Try
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

    Private Function GetContact(exch As clsPowerShell, objectGUID As String) As PSObject
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-MailContact")
        cmd.AddParameter("Identity", objectGUID)

        Dim obj As Collection(Of PSObject) = exch.Command(cmd)
        If obj Is Nothing OrElse (obj.Count <> 1) Then Return Nothing

        Return obj(0)
    End Function

    Private Function GetAcceptedDomain(exch As clsPowerShell) As Collection(Of PSObject)
        Dim cmd As New PSCommand
        cmd.AddCommand("Get-AcceptedDomain")

        Dim obj As Collection(Of PSObject) = _exchangeconnection.Command(cmd)
        If obj Is Nothing Then Return Nothing

        Return obj
    End Function

    Public ReadOnly Property Exist As Boolean
        Get
            Return _contact IsNot Nothing
        End Get
    End Property

    Public ReadOnly Property State As RunspaceStateInfo
        Get
            If _exchangeconnection IsNot Nothing Then
                Return _exchangeconnection.State
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property ExchangeConnection As clsPowerShell
        Get
            Return _exchangeconnection
        End Get
    End Property

    Public ReadOnly Property Connected As Boolean
        Get
            Return _exchangeconnection IsNot Nothing AndAlso _exchangeconnection.State.State = RunspaceState.Opened
        End Get
    End Property

    Public ReadOnly Property VersionFormatted As String
        Get
            If _exchangeserver Is Nothing Then Return ""

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
            If _contactaccepteddomain Is Nothing Then Return Nothing

            Return _contactaccepteddomain.ToArray().Select(Function(x As PSObject) x.Properties("DomainName").Value.ToString).ToArray
        End Get
    End Property

    Public ReadOnly Property EmailAddresses As ObservableCollection(Of clsEmailAddress)
        Get
            If _contact Is Nothing Then Return Nothing

            Dim obj As PSObject = _contact.Properties("EmailAddresses").Value
            If obj Is Nothing Then Return Nothing

            Dim _primary As String = Replace(LCase(_contact.Properties("ExternalEmailAddress").Value), "smtp:", "")
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
            If _contact Is Nothing Then Return False

            Dim obj As Boolean = _contact.Properties("HiddenFromAddressListsEnabled").Value

            Return obj
        End Get
        Set(value As Boolean)
            If _contact Is Nothing Then Exit Property

            Dim cmd As New PSCommand
            cmd.AddCommand("Set-MailContact")
            cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)
            cmd.AddParameter("HiddenFromAddressListsEnabled", value)

            Dim dummy = _exchangeconnection.Command(cmd)

            _contact = GetContact(_exchangeconnection, _currentobject.objectGUIDFormated)

            NotifyPropertyChanged("HiddenFromAddressListsEnabled")
        End Set
    End Property

    Public Sub Add(name As String, domain As String)
        If _contact IsNot Nothing Then

            Dim obj As PSObject = _contact.Properties("EmailAddresses").Value
            If obj Is Nothing Then Exit Sub

            obj.Methods.Item("Add").Invoke("smtp:" & name & "@" & domain)

            Dim cmd As New PSCommand
            cmd.AddCommand("Set-MailContact")
            cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)
            cmd.AddParameter("EmailAddresses", obj)

            Dim dummy = _exchangeconnection.Command(cmd)

            _contact = GetContact(_exchangeconnection, _currentobject.objectGUIDFormated)
            NotifyPropertyChanged("EmailAddresses")

        Else

            Dim cmd As New PSCommand
            cmd = New PSCommand
            cmd.AddCommand("Enable-MailContact")
            cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)
            cmd.AddParameter("ExternalEmailAddress", "smtp:" & name & "@" & domain)

            Dim dummy = _exchangeconnection.Command(cmd)

            For I As Integer = 0 To 30
                Threading.Thread.Sleep(1000)

                cmd = New PSCommand
                cmd.AddCommand("Get-MailContact")
                cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)

                Dim obj As Collection(Of PSObject) = _exchangeconnection.Command(cmd)
                If obj IsNot Nothing AndAlso (obj.Count = 1) Then Exit For
            Next

            UpdateContactInfo()

        End If
    End Sub

    Public Sub Edit(newname As String, newdomain As String, oldaddress As clsEmailAddress)
        If _contact Is Nothing Then Exit Sub

        Dim cmd As New PSCommand

        If oldaddress.IsPrimary Then

            cmd = New PSCommand
            cmd.AddCommand("Set-MailContact")
            cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)
            cmd.AddParameter("ExternalEmailAddress", "smtp:" & newname & "@" & newdomain)

            Dim dummy = _exchangeconnection.Command(cmd)

            Dim obj As PSObject = _contact.Properties("EmailAddresses").Value
            If obj Is Nothing Then Exit Sub

            For I As Integer = 0 To CType(obj.BaseObject, ArrayList).Count - 1
                If LCase(CType(obj.BaseObject, ArrayList)(I)) = LCase(oldaddress.AddressFull) Then
                    obj.Methods.Item("Remove").Invoke(CType(obj.BaseObject, ArrayList)(I))
                    Exit For
                End If
            Next
            obj.Methods.Item("Add").Invoke("smtp:" & newname & "@" & newdomain)

            cmd = New PSCommand
            cmd.AddCommand("Set-MailContact")
            cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)
            cmd.AddParameter("EmailAddresses", obj)

            Dim dummy3 = _exchangeconnection.Command(cmd)

            _contact = GetContact(_exchangeconnection, _currentobject.objectGUIDFormated)
            NotifyPropertyChanged("EmailAddresses")
        Else

            Dim obj As PSObject = _contact.Properties("EmailAddresses").Value
            If obj Is Nothing Then Exit Sub

            For I As Integer = 0 To CType(obj.BaseObject, ArrayList).Count - 1
                If LCase(CType(obj.BaseObject, ArrayList)(I)) = LCase(oldaddress.AddressFull) Then
                    obj.Methods.Item("Remove").Invoke(CType(obj.BaseObject, ArrayList)(I))
                    Exit For
                End If
            Next
            obj.Methods.Item("Add").Invoke("smtp:" & newname & "@" & newdomain)

            cmd = New PSCommand
            cmd.AddCommand("Set-MailContact")
            cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)
            cmd.AddParameter("EmailAddresses", obj)

            Dim dummy = _exchangeconnection.Command(cmd)

            _contact = GetContact(_exchangeconnection, _currentobject.objectGUIDFormated)

            NotifyPropertyChanged("EmailAddresses")
        End If
    End Sub

    Public Sub Remove(oldaddress As clsEmailAddress)
        If _contact Is Nothing Then Exit Sub

        Dim cmd As New PSCommand

        If oldaddress.IsPrimary Then

            cmd = New PSCommand
            cmd.AddCommand("Disable-MailContact")
            cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)
            cmd.AddParameter("Confirm", False)

            Dim dummy = _exchangeconnection.Command(cmd)

            For I As Integer = 0 To 30
                Threading.Thread.Sleep(1000)

                cmd = New PSCommand
                cmd.AddCommand("Get-MailContact")
                cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)

                Dim obj As Collection(Of PSObject) = _exchangeconnection.Command(cmd)
                If obj Is Nothing OrElse (obj.Count <> 1) Then Exit For
            Next

            UpdateContactInfo()

        Else

            Dim obj As PSObject = _contact.Properties("EmailAddresses").Value
            If obj Is Nothing Then Exit Sub

            For I As Integer = 0 To CType(obj.BaseObject, ArrayList).Count - 1
                If LCase(CType(obj.BaseObject, ArrayList)(I)) = LCase(oldaddress.AddressFull) Then
                    obj.Methods.Item("Remove").Invoke(CType(obj.BaseObject, ArrayList)(I))
                    Exit For
                End If
            Next

            cmd = New PSCommand
            cmd.AddCommand("Set-MailContact")
            cmd.AddParameter("Identity", _currentobject.objectGUIDFormated)
            cmd.AddParameter("EmailAddresses", obj)

            Dim dummy = _exchangeconnection.Command(cmd)

            _contact = GetContact(_exchangeconnection, _currentobject.objectGUIDFormated)
            NotifyPropertyChanged("EmailAddresses")
        End If
    End Sub

    Public Sub SetPrimary(mail As clsEmailAddress)
        If _contact Is Nothing Then Exit Sub

        Dim cmd As New PSCommand

        Dim _primary As String = Replace(LCase(_contact.Properties("ExternalEmailAddress").Value), "smtp:", "")
        If _primary Is Nothing Then Exit Sub

        If mail.IsPrimary Then

            ThrowCustomException(My.Resources.str_SelectedAddressAlreadyPrimary)

        Else

            Dim a As String()

            Dim newname As String = Nothing
            Dim newdomain As String = Nothing
            a = LCase(mail.Address).Split({"@"}, StringSplitOptions.RemoveEmptyEntries)
            If a.Count >= 2 Then
                newname = a(0)
                newdomain = a(1)
            End If

            cmd = New PSCommand
            cmd.AddCommand("Set-MailContact")
            cmd.AddParameter("Identity", _currentobject.name)
            cmd.AddParameter("ExternalEmailAddress", "smtp:" & newname & "@" & newdomain)

            Dim dummy = _exchangeconnection.Command(cmd)

            _contact = GetContact(_exchangeconnection, _currentobject.name)
            NotifyPropertyChanged("EmailAddresses")
        End If
    End Sub
End Class
