Imports System.Collections.ObjectModel
Imports System.Reflection
Imports Microsoft.Win32
Imports System.Windows.Forms
Imports System.DirectoryServices
Imports IRegisty
Imports CredentialManagement

Module mdlTools

    Public regADToolsApplication As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADTools")
    Public regDomains As RegistryKey = regADToolsApplication.CreateSubKey("Domains")
    Public dispatcherTimer As Windows.Threading.DispatcherTimer

    Public domains As New ObservableCollection(Of clsDomain)

    Public TelegramUsername As String
    Public TelegramAPIKey As String

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
        {New clsAttribute("userPrincipalName", "Имя входа")},
        {New clsAttribute("sAMAccountName", "Имя входа (пред-Windows 2000)")}
    }
    Public attributesForSearchExchangePermissionTarget As New ObservableCollection(Of clsAttribute) From { ' 
        {New clsAttribute("name", "Имя объекта")},
        {New clsAttribute("displayName", "Отображаемое имя")},
        {New clsAttribute("userPrincipalName", "Имя входа")}
    }
    Public attributesForSearchExchangePermissionFullAccess As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("sAMAccountName", "Имя входа (пред-Windows 2000)")}
    }
    Public attributesForSearchExchangePermissionSendAs As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("sAMAccountName", "Имя входа (пред-Windows 2000)")}
    }
    Public attributesForSearchExchangePermissionSendOnBehalf As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("name", "Имя объекта")}
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

    Public Sub initializeCredentials()
        Dim cred As New Credential("", "", "ADToolsTelegramBot", CredentialType.Generic)
        cred.PersistanceType = PersistanceType.Enterprise
        cred.Load()
        TelegramUsername = cred.Username
        TelegramAPIKey = cred.Password

        If String.IsNullOrEmpty(TelegramUsername) Or String.IsNullOrEmpty(TelegramAPIKey) Then Application.Current.Shutdown()

        Bot = New TeleBotDotNet.TeleBot(TelegramAPIKey, False)
    End Sub

    Public Sub initializeTimer()
        dispatcherTimer = New Threading.DispatcherTimer()
        AddHandler dispatcherTimer.Tick, AddressOf dispatcherTimer_Tick
        dispatcherTimer.Interval = New TimeSpan(0, 0, 1)
        dispatcherTimer.Start()
    End Sub

    Public Sub initializeDomains()
        domains = IRegistrySerializer.Deserialize(GetType(ObservableCollection(Of clsDomain)), regDomains)
    End Sub

    Private Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        GetTelegramMessages()
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
        MsgBox(Message, vbExclamation, "ADToolsTelegramBot")
    End Sub

    Public Sub ShowWrongMemberMessage()
        MsgBox("Глобальная группа может быть членом другой глобальной группы, универсальной группы или локальной группы домена." & vbCrLf &
               "Универсальная группа может быть членом другой универсальной группы или локальной группы домена, но не может быть членом глобальной группы." & vbCrLf &
               "Локальная группа домена может быть членом только другой локальной группы домена." & vbCrLf & vbCrLf &
               "Локальную группу домена можно преобразовать в универсальную группу лишь в том случае, если эта локальная группа домена не содержит других членов локальной группы домена. Локальная группа домена не может быть членом универсальной группы." & vbCrLf &
               "Глобальную группу можно преобразовать в универсальную лишь в том случае, если эта глобальная группа не входит в состав другой глобальной группы." & vbCrLf &
               "Универсальная группа не может быть членом глобальной группы.", vbOKOnly + vbExclamation, "Неверный тип группы")
    End Sub

    Public Sub Log(Text As String)
        Debug.Print(Text)
        'MyLog.WriteEntry("ROInventoryTelegramService", Text, EventLogEntryType.Information)
    End Sub

    Public Function Encode58(data As Byte()) As String
        Const alphabet As String = "123456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz"
        Dim i As Integer = 0

        ' Decode byte[] to BigInteger
        Dim intData As Numerics.BigInteger = 0
        For i = 0 To data.Length - 1
            intData = intData * 256 + data(i)
        Next

        ' Encode BigInteger to Base58 string
        Dim result As String = ""
        While intData > 0
            Dim remainder As Integer = CInt(intData Mod 58)
            intData /= 58
            result = alphabet(remainder) + result
        End While

        ' Append `1` for each leading 0 byte
        While i < data.Length AndAlso data(i) = 0
            result = Convert.ToString("1"c) & result
            i += 1
        End While

        Return result
    End Function

    Public Function Decode58(s As String) As Byte()
        Const alphabet As String = "123456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz"
        ' Decode Base58 string to BigInteger 
        Dim intData As Numerics.BigInteger = 0
        For i As Integer = 0 To s.Length - 1
            Dim digit As Integer = alphabet.IndexOf(s(i))
            'Slow
            If digit < 0 Then
                Return Nothing
            End If
            intData = intData * 58 + digit
        Next

        ' Encode BigInteger to byte[]
        ' Leading zero bytes get encoded as leading `1` characters
        Dim leadingZeroCount As Integer = s.TakeWhile(Function(c) c = "1"c).Count()
        Dim leadingZeros = Enumerable.Repeat(CByte(0), leadingZeroCount)
        ' to big endian
        Dim bytesWithoutLeadingZeros = intData.ToByteArray().Reverse().SkipWhile(Function(b) b = 0)
        'strip sign byte
        Dim result = leadingZeros.Concat(bytesWithoutLeadingZeros).ToArray()
        Return result
    End Function

End Module
