Imports System.Collections.ObjectModel
Imports System.Reflection
Imports Microsoft.Win32
Imports IRegisty
Imports System.Windows.Forms
Imports IPrompt.VisualBasic
Imports System.DirectoryServices

Module mdlTools

    Public regApplication As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\" & My.Application.Info.AssemblyName)
    Public regDomains As RegistryKey = regApplication.CreateSubKey("Domains")
    Public regPreferences As RegistryKey = regApplication.CreateSubKey("Preferences")

    Public preferences As clsPreferences
    Public domains As New ObservableCollection(Of clsDomain)

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


    Public columnsDefault As New ObservableCollection(Of clsDataGridColumnInfo) From {
        New clsDataGridColumnInfo("⬕", New List(Of clsAttribute) From {New clsAttribute("Image", "⬕")}, 0, 45),
        New clsDataGridColumnInfo("Имя", New List(Of clsAttribute) From {New clsAttribute("name", "Имя объекта"), New clsAttribute("description", "Описание")}, 1, 220),
        New clsDataGridColumnInfo("Имя входа", New List(Of clsAttribute) From {New clsAttribute("userPrincipalName", "Имя входа"), New clsAttribute("distinguishedName", "LDAP-путь")}, 2, 450),
        New clsDataGridColumnInfo("Телефон", New List(Of clsAttribute) From {New clsAttribute("telephoneNumber", "Телефон"), New clsAttribute("physicalDeliveryOfficeName", "Офис")}, 3, 100),
        New clsDataGridColumnInfo("Место работы", New List(Of clsAttribute) From {New clsAttribute("title", "Должность"), New clsAttribute("department", "Подразделение"), New clsAttribute("company", "Компания")}, 4, 300),
        New clsDataGridColumnInfo("Основной адрес", New List(Of clsAttribute) From {New clsAttribute("mail", "Основной адрес")}, 5, 170),
        New clsDataGridColumnInfo("Объект", New List(Of clsAttribute) From {New clsAttribute("whenCreatedFormated", "Создан (формат)"), New clsAttribute("lastLogonFormated", "Последний вход (формат)"), New clsAttribute("accountExpiresFormated", "Объект истекает (формат)")}, 6, 150),
        New clsDataGridColumnInfo("Пароль", New List(Of clsAttribute) From {New clsAttribute("pwdLastSetFormated", "Пароль изменен (формат)"), New clsAttribute("passwordExpiresFormated", "Пароль истекает (формат)")}, 7, 150)}

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
        {New clsAttribute("Image", "⬕")},
        {New clsAttribute("initials", "Инициалы")},
        {New clsAttribute("lastLogonDate", "Последний вход")},
        {New clsAttribute("lastLogonFormated", "Последний вход (формат)")},
        {New clsAttribute("location", "Местонахождение")},
        {New clsAttribute("logonCount", "Входов")},
        {New clsAttribute("mail", "Основной адрес")},
        {New clsAttribute("manager", "Руководитель")},
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
        {New clsAttribute("Status", "Статус")},
        {New clsAttribute("StatusFormated", "Статус (формат)")},
        {New clsAttribute("telephoneNumber", "Телефон")},
        {New clsAttribute("thumbnailPhoto", "Фото")},
        {New clsAttribute("title", "Должность")},
        {New clsAttribute("userPrincipalName", "Имя входа")},
        {New clsAttribute("whenCreated", "Создан")},
        {New clsAttribute("whenCreatedFormated", "Создан (формат)")}
    }
    Public attributesForSearchDefault As New ObservableCollection(Of clsAttribute) From {
        {New clsAttribute("name", "Имя объекта")},
        {New clsAttribute("displayName", "Отображаемое имя")},
        {New clsAttribute("givenName", "Имя")},
        {New clsAttribute("sn", "Фамилия")},
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

    Public propertiesToLoadDefault As String() =
        {"objectGUID",
        "userAccountControl",
        "accountExpires",
        "name",
        "description",
        "userPrincipalName",
        "distinguishedName",
        "telephoneNumber",
        "physicalDeliveryOfficeName",
        "title",
        "department",
        "company",
        "mail",
        "whenCreated",
        "lastLogon",
        "pwdLastSet",
        "thumbnailPhoto",
        "memberOf",
        "givenName",
        "sn",
        "initials",
        "displayName",
        "manager",
        "sAMAccountName",
        "groupType",
        "dNSHostName",
        "location",
        "operatingSystem",
        "operatingSystemVersion"}

    Public Sub initializePreferences()
        'Handlebars.Configuration.TextEncoder = Nothing
        'Handlebars.RegisterHelper("lz",
        '    Sub(writer, context, parameters)
        '        Try
        '            If parameters(1).ToString = "*" Then
        '                writer.WriteSafeString("*")
        '            Else
        '                Dim int As String = Format(parameters(1), New String("0", parameters(0)))
        '                writer.WriteSafeString(int)
        '            End If
        '        Catch ex As Exception
        '        End Try
        '    End Sub)

        'Handlebars.RegisterHelper("fst",
        '    Sub(writer, context, parameters)
        '        Try
        '            If CInt(parameters(0)) > parameters(1).ToString.Length Then
        '                writer.WriteSafeString(parameters(1).ToString)
        '            Else
        '                writer.WriteSafeString(Mid(parameters(1).ToString, 1, parameters(0)))
        '            End If
        '        Catch ex As Exception
        '        End Try
        '    End Sub)

        'Handlebars.RegisterHelper("trans",
        '    Sub(writer, context, parameters)
        '        Try
        '            writer.WriteSafeString(Translit_RU_EN(parameters(0)))
        '        Catch ex As Exception
        '        End Try
        '    End Sub)

        'Handlebars.RegisterHelper("split",
        '    Sub(writer, context, parameters)
        '        Try
        '            writer.WriteSafeString(If(Split(parameters(1), " ").Count >= parameters(0), Split(parameters(1), " ")(parameters(0)), ""))
        '        Catch ex As Exception
        '        End Try
        '    End Sub)

        preferences = IRegistrySerializer.Deserialize(GetType(clsPreferences), regPreferences)
    End Sub

    Public Sub initializeDomains()
        domains = IRegistrySerializer.Deserialize(GetType(ObservableCollection(Of clsDomain)), regDomains)
    End Sub

    Public Sub deinitializePreferences()
        IRegistrySerializer.Serialize(preferences, regPreferences)
    End Sub

    Public Function ShowWindow(w As Window, Optional singleinstance As Boolean = False, Optional owner As Window = Nothing, Optional modal As Boolean = False) As Window
        If singleinstance Then
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If w.GetType Is wnd.GetType Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        w.Owner = owner
                        Return w
                    End If
                Next
            Else
                For Each wnd As Window In Application.Current.Windows
                    If w.GetType Is wnd.GetType Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        w.Owner = Nothing
                        Return w
                    End If
                Next
            End If
        End If

        w.Owner = owner
        If modal Then
            w.ShowDialog()
        Else
            w.Show()
        End If
        Return w
    End Function

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
        'ADToolsApplication.tsocErrorLog.Add(New clsErrorLog(Procedure,, ex))
    End Sub

    Public Sub ThrowCustomException(Message As String)
        'ADToolsApplication.tsocErrorLog.Add(New clsErrorLog(Message))
    End Sub

    Public Sub ThrowInformation(Message As String)
        With ADToolsApplication.nicon
            .BalloonTipIcon = ToolTipIcon.Info
            .BalloonTipTitle = My.Application.Info.AssemblyName
            .BalloonTipText = Message
            .Tag = Nothing
            .Visible = False
            .Visible = True
            .ShowBalloonTip(5000)
        End With
    End Sub

    Public Sub ShowWrongMemberMessage()
        IMsgBox("Глобальная группа может быть членом другой глобальной группы, универсальной группы или локальной группы домена." & vbCrLf &
               "Универсальная группа может быть членом другой универсальной группы или локальной группы домена, но не может быть членом глобальной группы." & vbCrLf &
               "Локальная группа домена может быть членом только другой локальной группы домена." & vbCrLf & vbCrLf &
               "Локальную группу домена можно преобразовать в универсальную группу лишь в том случае, если эта локальная группа домена не содержит других членов локальной группы домена. Локальная группа домена не может быть членом универсальной группы." & vbCrLf &
               "Глобальную группу можно преобразовать в универсальную лишь в том случае, если эта глобальная группа не входит в состав другой глобальной группы." & vbCrLf &
               "Универсальная группа не может быть членом глобальной группы.", vbOKOnly + vbExclamation, "Неверный тип группы")
    End Sub

    Public Function ShowDirectoryObjectProperties(obj As clsDirectoryObject, Optional owner As Window = Nothing) As Window
        If obj.SchemaClass = clsDirectoryObject.enmSchemaClass.User Then
            Dim w As wndUser
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndUser) Is wnd.GetType AndAlso CType(wnd, wndUser).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndUser
            If owner IsNot Nothing Then
                w.Owner = owner
            End If

            w.currentobject = obj
            w.Show()
            Return w
        ElseIf obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Then
            Dim w As wndContact
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndContact) Is wnd.GetType AndAlso CType(wnd, wndContact).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndContact
            If owner IsNot Nothing Then
                w.Owner = owner
            End If

            w.currentobject = obj
            w.Show()
            Return w
        ElseIf obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Then
            Dim w As wndComputer
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndComputer) Is wnd.GetType AndAlso CType(wnd, wndComputer).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndComputer
            If owner IsNot Nothing Then w.Owner = owner
            w.currentobject = obj
            w.Show()
            Return w
        ElseIf obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Group Then
            Dim w As wndGroup
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndGroup) Is wnd.GetType AndAlso CType(wnd, wndGroup).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndGroup
            If owner IsNot Nothing Then w.Owner = owner
            w.currentobject = obj
            w.Show()
            Return w
        ElseIf obj.objectClass.Contains("organizationalunit") Then
            Dim w As wndOrganizationalUnit
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndOrganizationalUnit) Is wnd.GetType AndAlso CType(wnd, wndOrganizationalUnit).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndOrganizationalUnit
            If owner IsNot Nothing Then w.Owner = owner
            w.currentobject = obj
            w.Show()
            Return w
        Else
            Dim w As wndUnknownObject
            If owner IsNot Nothing Then
                For Each wnd As Window In owner.OwnedWindows
                    If GetType(wndUnknownObject) Is wnd.GetType AndAlso CType(wnd, wndUnknownObject).currentobject Is obj Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                        Return Nothing
                    End If
                Next
            End If

            w = New wndUnknownObject
            If owner IsNot Nothing Then w.Owner = owner
            w.currentobject = obj
            w.Show()
            Return w
        End If
    End Function

    Public Sub Log(message As String)
        ADToolsApplication.tsocLog.Add(New clsLog(message))
    End Sub

    Public Function GetNextDomainTelephoneNumbers(domain As clsDomain) As ObservableCollection(Of clsTelephoneNumber)
        If domain Is Nothing Then Return Nothing
        Dim patterns As ObservableCollection(Of clsTelephoneNumberPattern) = domain.TelephoneNumberPattern
        If patterns.Count = 0 Then Return Nothing
        Dim de As DirectoryEntry = domain.DefaultNamingContext
        If de Is Nothing Then Return Nothing

        Dim result As New ObservableCollection(Of clsTelephoneNumber)

        Dim LDAPsearcher As New DirectorySearcher(de)
        Dim LDAPresults As SearchResultCollection = Nothing
        Dim LDAPresult As SearchResult

        LDAPsearcher.PropertiesToLoad.Add("objectCategory")
        LDAPsearcher.PropertiesToLoad.Add("objectClass")
        LDAPsearcher.PropertiesToLoad.Add("userAccountControl")
        LDAPsearcher.PropertiesToLoad.Add("telephoneNumber")

        For Each pattern As clsTelephoneNumberPattern In patterns
            If Not pattern.Range.Contains("-") Then Continue For
            Dim numstart As Long = 0
            Dim numend As Long = 0
            If Not Long.TryParse(pattern.Range.Split({"-"}, 2, StringSplitOptions.RemoveEmptyEntries)(0), numstart) Or
               Not Long.TryParse(pattern.Range.Split({"-"}, 2, StringSplitOptions.RemoveEmptyEntries)(1), numend) Then Continue For

            LDAPsearcher.Filter = "(&(objectCategory=person)(!(objectClass=inetOrgPerson))(!(UserAccountControl:1.2.840.113556.1.4.803:=2))(telephoneNumber=*))"
            LDAPsearcher.PageSize = 1000
            LDAPresults = LDAPsearcher.FindAll()

            Dim dummy As New List(Of String)
            For Each LDAPresult In LDAPresults
                dummy.Add(GetLDAPProperty(LDAPresult.Properties, "telephoneNumber"))
            Next LDAPresult

            For I As Long = numstart To numend
                Dim u As String = LCase(Format(I, pattern.Pattern))
                If Not dummy.Contains(u) Then
                    result.Add(New clsTelephoneNumber(pattern.Label, u))
                    Exit For
                End If
            Next
        Next

        Return result
    End Function

    Public Function GetApplicationIcon(fileName As String) As ImageSource
        Dim ai As System.Drawing.Icon = System.Drawing.Icon.ExtractAssociatedIcon(fileName)
        Return System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(ai.Handle, New Int32Rect(0, 0, ai.Width, ai.Height), BitmapSizeOptions.FromEmptyOptions())
    End Function

End Module
