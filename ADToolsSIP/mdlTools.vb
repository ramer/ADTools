Imports System.Collections.ObjectModel
Imports System.Reflection
Imports Microsoft.Win32
Imports System.DirectoryServices
Imports IRegisty
Imports CredentialManagement
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Windows.Controls.Primitives
Imports System.Windows.Threading

Module mdlTools

    Public regADToolsApplication As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADTools")
    Public regDomains As RegistryKey = regADToolsApplication.CreateSubKey("Domains")

    Public domains As New ObservableCollection(Of clsDomain)

    Public searcher As New clsSearcher

    Public SIP As clsSIP

    Public Const ADS_UF_SCRIPT = 1 '0x1
    Public Const ADS_UF_ACCOUNTDISABLE = 2 '0x2
    Public Const ADS_UF_HOMEDIR_REQUIRED = 8 '0x8
    Public Const ADS_UF_LOCKOUT = 16 '0x10
    Public Const ADS_UF_PASSWD_NOTREQD = 32 '0x20
    Public Const ADS_UF_PASSWD_CANT_CHANGE = 64 '0x40
    Public Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = 128 '0x80
    Public Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = 256 '0x100
    Public Const ADS_UF_NORMAL_ACCOUNT = 512 '0x200
    Public Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = 2048 '0x800
    Public Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = 4096 '0x1000
    Public Const ADS_UF_SERVER_TRUST_ACCOUNT = 8192 '0x2000
    Public Const ADS_UF_DONT_EXPIRE_PASSWD = 65536 '0x10000
    Public Const ADS_UF_MNS_LOGON_ACCOUNT = 131072 '0x20000
    Public Const ADS_UF_SMARTCARD_REQUIRED = 262144 '0x40000
    Public Const ADS_UF_TRUSTED_FOR_DELEGATION = 524288 '0x80000
    Public Const ADS_UF_NOT_DELEGATED = 1048576 '0x100000
    Public Const ADS_UF_USE_DES_KEY_ONLY = 2097152 '0x200000
    Public Const ADS_UF_DONT_REQUIRE_PREAUTH = 4194304 '0x400000
    Public Const ADS_UF_PASSWORD_EXPIRED = 8388608 '0x800000
    Public Const ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 16777216 '0x1000000

    Public Const ADS_GROUP_TYPE_GLOBAL_GROUP = 2 '0x00000002
    Public Const ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = 4 '0x00000004
    Public Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = 8 '0x00000008
    Public Const ADS_GROUP_TYPE_SECURITY_ENABLED = -2147483648 '0x80000000

    Public attributesDefault As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("accountExpires", "Объект истекает")},
        {New clsAttribute("accountExpiresFormated", "Объект истекает (формат)")},
        {New clsAttribute("badPwdCount", "Ошибок ввода пароля")},
        {New clsAttribute("company", "Компания")},
        {New clsAttribute("department", "Подразделение")},
        {New clsAttribute("description", "Описание")},
        {New clsAttribute("disabled", "Заблокирован")},
        {New clsAttribute("disabledFormated", "Заблокирован (формат)")},
        {New clsAttribute("displayName", "Отображаемое имя")},
        {New clsAttribute("distinguishedName", "LDAP-путь")},
        {New clsAttribute("distinguishedNameFormated", "LDAP-путь (формат)")},
        {New clsAttribute("givenName", "Имя")},
        {New clsAttribute("Image", "Картинка ⬕")},
        {New clsAttribute("initials", "Инициалы")},
        {New clsAttribute("lastLogonDate", "Последний вход")},
        {New clsAttribute("lastLogonFormated", "Последний вход (формат)")},
        {New clsAttribute("location", "Местонахождение")},
        {New clsAttribute("logonCount", "Входов")},
        {New clsAttribute("mail", "Основной адрес")},
        {New clsAttribute("manager", "Руководитель")},
        {New clsAttribute("managedBy", "Управляется")},
        {New clsAttribute("name", "Имя объекта")},
        {New clsAttribute("objectGUID", "Уникальный идентификатор (GUID)")},
        {New clsAttribute("objectSID", "Уникальный идентификатор (SID)")},
        {New clsAttribute("passwordExpiresDate", "Пароль истекает")},
        {New clsAttribute("passwordExpiresFormated", "Пароль истекает (формат)")},
        {New clsAttribute("physicalDeliveryOfficeName", "Офис")},
        {New clsAttribute("pwdLastSetDate", "Пароль изменен")},
        {New clsAttribute("pwdLastSetFormated", "Пароль изменен (формат)")},
        {New clsAttribute("sAMAccountName", "Имя входа (пред-Windows 2000)")},
        {New clsAttribute("SchemaClassName", "Класс")},
        {New clsAttribute("sn", "Фамилия")},
        {New clsAttribute("StatusImage", "Статус ⬕")},
        {New clsAttribute("StatusFormatted", "Статус (формат)")},
        {New clsAttribute("telephoneNumber", "Телефон")},
        {New clsAttribute("thumbnailPhoto", "Фото ⬕")},
        {New clsAttribute("title", "Должность")},
        {New clsAttribute("userPrincipalName", "Имя входа")},
        {New clsAttribute("whenCreated", "Создан")},
        {New clsAttribute("whenCreatedFormated", "Создан (формат)")}
    }
    Public attributesForSearchDefault As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("name", "Имя объекта")},
        {New clsAttribute("displayName", "Отображаемое имя")},
        {New clsAttribute("telephoneNumber", "Телефон")}
    }

    Public attributesToLoadDefault As String() =
        {"accountExpires",
        "company",
        "department",
        "description",
        "displayName",
        "distinguishedName",
        "dNSHostName",
        "givenName",
        "groupType",
        "initials",
        "isDeleted",
        "isRecycled",
        "lastLogon",
        "location",
        "mail",
        "manager",
        "memberOf",
        "name",
        "objectCategory",
        "objectClass",
        "objectGUID",
        "operatingSystem",
        "operatingSystemVersion",
        "physicalDeliveryOfficeName",
        "pwdLastSet",
        "sAMAccountName",
        "sn",
        "telephoneNumber",
        "thumbnailPhoto",
        "title",
        "userAccountControl",
        "userPrincipalName",
        "whenCreated"}

    Public Sub initializeSIP()
        SIP = New clsSIP

        Dim cred As New Credential("", "", "ADToolsSIP", CredentialType.Generic)
        cred.PersistanceType = PersistanceType.Enterprise
        cred.Load()
        SIP.Username = cred.Username
        SIP.Password = cred.Password
        Dim regADToolsSIP As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADToolsSIP")
        SIP.Server = regADToolsSIP.GetValue("Server", "")
        SIP.Protocol = If(regADToolsSIP.GetValue("Protocol", "UDP") = "UDP", LumiSoft.Net.BindInfoProtocol.UDP, LumiSoft.Net.BindInfoProtocol.TCP)
        SIP.RegistrationName = regADToolsSIP.GetValue("RegistrationName", "")
        SIP.Domain = regADToolsSIP.GetValue("Domain", "")

        SIP.Register()
    End Sub

    Public Sub initializeDomains()
        domains = IRegistrySerializer.Deserialize(GetType(ObservableCollection(Of clsDomain)), regDomains)
    End Sub

    Public Function GetLDAPProperty(ByRef Properties As DirectoryServices.ResultPropertyCollection, ByVal Prop As String)
        Try
            If Properties(Prop).Count > 0 Then
                Return Properties(Prop)(0)
            Else
                Return ""
            End If
        Catch
            Return ""
        End Try
    End Function

    Public Function GetLDAPProperty(ByRef Properties As DirectoryServices.PropertyCollection, ByVal Prop As String)
        Try
            If Properties(Prop).Count > 0 Then
                Return Properties(Prop)(0)
            Else
                Return ""
            End If
        Catch
            Return ""
        End Try
    End Function

    Public Function LongFromLargeInteger(largeInteger As Object) As Long
        Dim valBytes(7) As Byte
        Dim result As Long
        Dim type As System.Type = largeInteger.[GetType]()
        Dim highPart As Integer = CInt(type.InvokeMember("HighPart", BindingFlags.GetProperty, Nothing, largeInteger, Nothing))
        Dim lowPart As Integer = CInt(type.InvokeMember("LowPart", BindingFlags.GetProperty, Nothing, largeInteger, Nothing))
        BitConverter.GetBytes(lowPart).CopyTo(valBytes, 0)
        BitConverter.GetBytes(highPart).CopyTo(valBytes, 4)

        result = BitConverter.ToInt64(valBytes, 0)
        If result = 9223372036854775807 Then result = 0

        Return result
    End Function

    Public Sub UnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Dim ex As Exception = DirectCast(e.ExceptionObject, Exception)
        ThrowException(ex, "Необработанное исключение")
    End Sub

    Public Sub ThrowException(ByVal ex As Exception, ByVal Procedure As String)
        MsgBox(ex.Message, vbExclamation, Procedure)
    End Sub

    Public Sub ThrowCustomException(Message As String)
        MsgBox(Message, vbExclamation, "ADToolsSIP")
    End Sub

    Public Sub Log(Text As String)
        Debug.Print(Text)
        'MyLog.WriteEntry("ROInventoryTelegramService", Text, EventLogEntryType.Information)
    End Sub

    Public Function GetLocalIPAddress() As String
        Dim host = Dns.GetHostEntry(Dns.GetHostName())
        For Each ip As IPAddress In host.AddressList
            If ip.AddressFamily = AddressFamily.InterNetwork Then
                Return ip.ToString()
            End If
        Next
        ThrowCustomException("Local IP Address Not Found!")
        Return Nothing
    End Function

    Public Sub ShowWrongMemberMessage()
    End Sub

    Public Sub ThrowSIPInformation(request As LumiSoft.Net.SIP.Stack.SIP_Request)
        Dim displayName As String = Encoding.UTF8.GetString(Encoding.GetEncoding(1251).GetBytes(request.From.Address.DisplayName))
        Dim telephoneNumber As String = request.From.Address.Uri.Value.Split({"@"}, StringSplitOptions.RemoveEmptyEntries).First

        Dim objects As New List(Of clsDirectoryObject)
        For Each dmn In domains
            objects.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter(displayName & "/*" & telephoneNumber, attributesForSearchDefault, New clsSearchObjectClasses(True, False, False, False, False))))
        Next

        ShowPopup(displayName, telephoneNumber, objects)
    End Sub

    Public Sub ShowPopup(displayName As String, telephoneNumber As String, objects As List(Of clsDirectoryObject))
        Dim popHeaderText = New TextBlock With {.Style = Windows.Application.Current.FindResource("PopupHeaderTextStyle")}
        popHeaderText.Text = String.Format("{0} ({1})", displayName, telephoneNumber)

        Dim popContentListBox As New ListBox With {.Style = Windows.Application.Current.FindResource("PopupListBoxStyle")}
        popContentListBox.ItemsSource = objects

        AddHandler popContentListBox.PreviewMouseDown,
            Sub(sender As Object, e As MouseButtonEventArgs)
                Dim obj = FindVisualParent(Of ListBoxItem)(e.OriginalSource)
                If obj Is Nothing OrElse TypeOf obj.DataContext IsNot clsDirectoryObject Then Exit Sub
                Process.Start("..\..\ADTools.exe", """" & CType(obj.DataContext, clsDirectoryObject).objectGUIDFormated & """")
            End Sub

        Dim pop = New Popup With {.Style = Windows.Application.Current.FindResource("PopupStyle"), .HorizontalOffset = Forms.Screen.PrimaryScreen.WorkingArea.Right - 5, .VerticalOffset = Forms.Screen.PrimaryScreen.WorkingArea.Bottom - 5}
        pop.Child = New HeaderedContentControl With {.Style = Windows.Application.Current.FindResource("PopupContentStyle"), .Content = popContentListBox, .Header = popHeaderText}

        Dim popTimer As New DispatcherTimer With {.Interval = TimeSpan.FromSeconds(5)}

        AddHandler pop.Opened,
            Sub()
                popTimer.Start()
                AddHandler popTimer.Tick,
                Sub()
                    pop.IsOpen = False
                    popTimer.Stop()
                    popTimer = Nothing
                    pop = Nothing
                End Sub
            End Sub

        AddHandler pop.MouseMove, Sub() If popTimer IsNot Nothing Then popTimer.Stop() : popTimer.Start()

        pop.IsOpen = True
    End Sub

    Public Function FindVisualParent(Of T As DependencyObject)(ByVal child As Object) As T
        Dim parent As DependencyObject = If(child.Parent IsNot Nothing, child.Parent, VisualTreeHelper.GetParent(child))

        If parent IsNot Nothing Then
            If TypeOf parent Is T Then
                Return parent
            Else
                Return FindVisualParent(Of T)(parent)
            End If
        Else
            Return Nothing
        End If
    End Function

End Module
