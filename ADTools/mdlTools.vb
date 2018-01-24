Imports System.Collections.ObjectModel
Imports Microsoft.Win32
Imports IRegisty
Imports IPrompt.VisualBasic
Imports System.DirectoryServices.Protocols
Imports HandlebarsDotNet
Imports System.Windows.Markup
Imports System.Globalization
Imports System.Security
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation

Public Enum enmSearchMode
    [Default] = 0
    Advanced = 1
End Enum

Public Enum enmClipboardAction
    Copy = 0
    Cut = 1
End Enum

Module mdlTools

    Public regApplication As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\" & My.Application.Info.AssemblyName)
    Public regDomains As RegistryKey = regApplication.CreateSubKey("Domains")
    Public regPreferences As RegistryKey = regApplication.CreateSubKey("Preferences")

    Public preferences As clsPreferences
    Public domains As New ObservableCollection(Of clsDomain)

    Public applicationdeactivating As Boolean = False

    Public ClipboardBuffer As clsDirectoryObject()
    Public ClipboardAction As enmClipboardAction

    Public Const OBJECT_DUALPANEL_MINWIDTH As Integer = 610

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


    Public columnsDefault As New ObservableCollection(Of clsViewColumnInfo) From {
        New clsViewColumnInfo("Статус", New List(Of clsAttribute) From {New clsAttribute("StatusImage", "Статус ⬕")}, 0, 51),
        New clsViewColumnInfo("Имя", New List(Of clsAttribute) From {New clsAttribute("name", "Имя объекта"), New clsAttribute("description", "Описание")}, 1, 220),
        New clsViewColumnInfo("Имя входа", New List(Of clsAttribute) From {New clsAttribute("userPrincipalName", "Имя входа"), New clsAttribute("distinguishedNameFormated", "LDAP-путь (формат)")}, 2, 450),
        New clsViewColumnInfo("Телефон", New List(Of clsAttribute) From {New clsAttribute("telephoneNumber", "Телефон"), New clsAttribute("physicalDeliveryOfficeName", "Офис")}, 3, 100),
        New clsViewColumnInfo("Место работы", New List(Of clsAttribute) From {New clsAttribute("title", "Должность"), New clsAttribute("department", "Подразделение"), New clsAttribute("company", "Компания")}, 4, 300),
        New clsViewColumnInfo("Основной адрес", New List(Of clsAttribute) From {New clsAttribute("mail", "Основной адрес")}, 5, 170),
        New clsViewColumnInfo("Объект", New List(Of clsAttribute) From {New clsAttribute("whenCreatedFormated", "Создан (формат)"), New clsAttribute("lastLogonFormated", "Последний вход (формат)"), New clsAttribute("accountExpiresFormated", "Объект истекает (формат)")}, 6, 150),
        New clsViewColumnInfo("Пароль", New List(Of clsAttribute) From {New clsAttribute("pwdLastSetFormated", "Пароль изменен (формат)"), New clsAttribute("passwordExpiresFormated", "Пароль истекает (формат)")}, 7, 150)}

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
        {New clsAttribute("objectGUIDFormated", "Уникальный идентификатор (GUID) (формат)")},
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

    Public portlistDefault As New Dictionary(Of Integer, String) From
        {{135, "RPC"},
        {139, "NETBIOS-SSN"},
        {445, "SMB over TCP"},
        {3389, "RDP"},
        {4899, "Radmin"},
        {5900, "VNC"},
        {6129, "DameWare RC"}}

    Public Sub initializePreferences()
        preferences = IRegistrySerializer.Deserialize(GetType(clsPreferences), regPreferences)
    End Sub

    Public Sub initializeDomains()
        domains = IRegistrySerializer.Deserialize(GetType(ObservableCollection(Of clsDomain)), regDomains)
    End Sub

    Public Sub deinitializePreferences()
        Array.ForEach(Of String)(regPreferences.GetSubKeyNames, New Action(Of String)(Sub(p) regPreferences.DeleteSubKeyTree(p, False)))
        IRegistrySerializer.Serialize(preferences, regPreferences)
    End Sub

    Public Sub initializeGlobalParameters()
        FrameworkElement.LanguageProperty.OverrideMetadata(GetType(FrameworkElement), New FrameworkPropertyMetadata(XmlLanguage.GetLanguage(CultureInfo.CurrentCulture.IetfLanguageTag)))

        Handlebars.Configuration.TextEncoder = Nothing
        Handlebars.RegisterHelper("lz",
            Sub(writer, context, parameters)
                Try
                    If IsNumeric(parameters(1)) Then
                        Dim int As String = Format(parameters(1), New String("0", parameters(0)))
                        writer.WriteSafeString(int)
                    Else
                        writer.WriteSafeString(parameters(1))
                    End If
                Catch ex As Exception
                End Try
            End Sub)

        Handlebars.RegisterHelper("fst",
            Sub(writer, context, parameters)
                Try
                    If CInt(parameters(0)) > parameters(1).ToString.Length Then
                        writer.WriteSafeString(parameters(1).ToString)
                    Else
                        writer.WriteSafeString(Mid(parameters(1).ToString, 1, parameters(0)))
                    End If
                Catch ex As Exception
                End Try
            End Sub)

        Handlebars.RegisterHelper("trans",
            Sub(writer, context, parameters)
                Try
                    writer.WriteSafeString(Transliterate_RU_EN(parameters(0)))
                Catch ex As Exception
                End Try
            End Sub)

        Handlebars.RegisterHelper("split",
            Sub(writer, context, parameters)
                Try
                    writer.WriteSafeString(If(Split(parameters(1), " ").Count >= parameters(0), Split(parameters(1), " ")(parameters(0)), ""))
                Catch ex As Exception
                End Try
            End Sub)
    End Sub

    Public Function ShowPage(p As Page, Optional singleinstance As Boolean = False, Optional owner As Window = Nothing, Optional modal As Boolean = False) As NavigationWindow
        If applicationdeactivating Then Return Nothing
        If p Is Nothing Then Return Nothing

        Dim w As NavigationWindow = Nothing

        If preferences.PageInterface Then

            If owner Is Nothing Then
                w = New NavigationWindow
                w.WindowStartupLocation = WindowStartupLocation.CenterScreen
            Else
                w = owner
            End If

            p.WindowWidth = Double.NaN
            p.WindowHeight = Double.NaN
            w.Navigate(p)

            w.Show()

            Return w

        Else

            If owner Is Nothing Then

                w = New NavigationWindow
                w.WindowStartupLocation = WindowStartupLocation.CenterScreen
                w.ShowInTaskbar = False
                w.Navigate(p)
                w.UpdateLayout()

                If modal Then
                    w.ShowDialog()
                Else
                    w.Show()
                End If

                Return w

            Else

                For Each wnd As Window In owner.OwnedWindows
                    If TypeOf wnd Is NavigationWindow AndAlso p.GetType Is wnd.Content.GetType AndAlso TypeOf p Is pgObject AndAlso wnd.Content.CurrentObject Is CType(p, Object).CurrentObject Then
                        w = wnd
                        w.Show() : w.Activate()
                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                        w.Topmost = True : w.Topmost = False
                    End If
                Next

                If w Is Nothing Then w = New NavigationWindow
                w.Owner = owner
                w.WindowStartupLocation = WindowStartupLocation.CenterOwner
                w.ShowInTaskbar = False

                If TypeOf p Is pgObject Then w.Width = 900 : w.Height = 620

                w.Navigate(p)

                If modal Then
                    w.ShowDialog()
                Else
                    w.Show()
                End If

                Return w

            End If

        End If

    End Function

    Public Function ShowDirectoryObjectProperties(obj As clsDirectoryObject, Optional owner As Window = Nothing) As NavigationWindow
        Dim p As Page

        If obj.SchemaClass = clsDirectoryObject.enmSchemaClass.User Then
            p = New pgObject(obj)
        ElseIf obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Contact Then
            p = New pgObject(obj)
        ElseIf obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Computer Then
            p = New pgObject(obj)
        ElseIf obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Group Then
            p = New pgObject(obj)
        ElseIf obj.SchemaClass = clsDirectoryObject.enmSchemaClass.OrganizationalUnit Then
            p = New pgObject(obj)
        Else
            p = New pgObject(obj)
        End If

        Return ShowPage(p, False, owner, False)
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

    Public Function GetDefaultColumns()
        Dim results As New ObservableCollection(Of clsViewColumnInfo)
        For Each c In columnsDefault
            results.Add(c)
        Next
        Return results
    End Function

    Public Function GetViewDetailsStyle() As Style

        Dim newstyle As New Style
        newstyle.BasedOn = Windows.Application.Current.TryFindResource(GetType(ListView))
        newstyle.TargetType = GetType(ListView)

        newstyle.Setters.Add(New Setter(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto))
        newstyle.Setters.Add(New Setter(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Auto))
        newstyle.Setters.Add(New Setter(ScrollViewer.CanContentScrollProperty, True))
        newstyle.Setters.Add(New Setter(VirtualizingPanel.IsVirtualizingProperty, True))
        newstyle.Setters.Add(New Setter(VirtualizingPanel.IsVirtualizingWhenGroupingProperty, True))
        newstyle.Setters.Add(New Setter(VirtualizingStackPanel.VirtualizationModeProperty, VirtualizationMode.Recycling))
        newstyle.Setters.Add(New Setter(KeyboardNavigation.DirectionalNavigationProperty, KeyboardNavigationMode.None))

        Dim gridview As New GridView

        For Each columninfo As clsViewColumnInfo In preferences.Columns
            Dim column As New GridViewColumn()
            column.Header = columninfo.Header
            'column.SetValue(DataGridColumn.CanUserSortProperty, True)
            'If columninfo.DisplayIndex > 0 Then column.DisplayIndex = columninfo.DisplayIndex
            column.Width = If(columninfo.Width > 0, columninfo.Width, Double.NaN)
            'column.MinWidth = 58
            Dim panel As New FrameworkElementFactory(GetType(VirtualizingStackPanel))
            panel.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center)
            panel.SetValue(FrameworkElement.MarginProperty, New Thickness(5, 0, 5, 0))

            Dim first As Boolean = True
            For Each attr As clsAttribute In columninfo.Attributes
                Dim bind As New Binding(attr.Name) With {.Mode = BindingMode.OneWay, .Converter = New ConverterDataToUIElement, .ConverterParameter = attr.Name}

                Dim container As New FrameworkElementFactory(GetType(ContentControl))
                If first Then
                    first = False
                    container.SetValue(TextBlock.FontWeightProperty, FontWeights.Medium)
                    'column.SetValue(DataGridColumn.SortMemberPathProperty, attr.Name)
                Else
                    container.SetValue(TextBlock.FontWeightProperty, FontWeights.Light)
                End If

                container.SetBinding(ContentControl.ContentProperty, bind)
                container.SetValue(FrameworkElement.ToolTipProperty, attr.Label)
                container.SetValue(FrameworkElement.MaxHeightProperty, 48.0)
                panel.AppendChild(container)
            Next

            Dim template As New DataTemplate()
            template.VisualTree = panel
            column.CellTemplate = template
            gridview.Columns.Add(column)
        Next

        newstyle.Setters.Add(New Setter(ListView.ViewProperty, gridview))

        Return newstyle
    End Function

    Public Function GetAttributesExtended() As clsAttribute()
        Dim attributes As New HashSet(Of String)
        Dim filters As New List(Of String) From
            {"(&(objectCategory=person)(objectClass=user)(!(objectClass=inetOrgPerson)))",
            "(&(objectCategory=person)(objectClass=contact))",
            "(objectClass=computer)",
            "(objectClass=group)",
            "(objectClass=organizationalunit)"}

        For Each domain In domains
            Try
                For Each f In filters
                    Dim searchRequest As SearchRequest = New SearchRequest(domain.DefaultNamingContext, f, SearchScope.Subtree, Nothing)
                    searchRequest.Controls.Add(New PageResultRequestControl(1))
                    searchRequest.Controls.Add(New SearchOptionsControl(SearchOption.DomainScope))
                    Dim response As SearchResponse = domain.Connection.SendRequest(searchRequest)
                    If response.Entries.Count = 1 Then
                        Dim sampleobject As New clsDirectoryObject(response.Entries(0), domain)
                        For Each a As String In sampleobject.AllowedAttributes
                            attributes.Add(a)
                        Next
                    End If
                Next
            Catch ex As Exception

            End Try
        Next

        Return attributes.ToArray.OrderBy(Function(x As String) x).Select(Function(x As String) New clsAttribute(x, x)).ToArray
    End Function

    Public Sub ThrowException(ByVal ex As Exception, ByVal Procedure As String)
        ADToolsApplication.tsocErrorLog.Add(New clsErrorLog(Procedure,, ex))
    End Sub

    Public Sub ThrowCustomException(Message As String)
        ADToolsApplication.tsocErrorLog.Add(New clsErrorLog(Message))
    End Sub

    Public Sub ThrowInformation(Message As String)
        With ADToolsApplication.nicon
            .BalloonTipIcon = Forms.ToolTipIcon.Info
            .BalloonTipTitle = My.Application.Info.AssemblyName
            .BalloonTipText = Message
            .Tag = Nothing
            .Visible = False
            .Visible = True
            .ShowBalloonTip(5000)
        End With
    End Sub

    Public Sub ShowWrongMemberMessage()
        IMsgBox(My.Resources.str_WrongGroupMember, vbOKOnly + vbExclamation, My.Resources.str_WrongGroupMemberTitle)
    End Sub

    Public Sub Log(message As String)
        ADToolsApplication.tsocLog.Add(New clsLog(message))
    End Sub

    Public Function GetNextDomainUsers(domain As clsDomain, Optional displayname As String = "") As List(Of String)
        If domain Is Nothing Then Return Nothing

        Dim result As New List(Of String)
        Dim searcher As New clsSearcher

        For Each template In domain.UsernamePatternTemplates
            Dim starredData = New With {.displayname = displayname, .n = "*"}

            Dim users As ObservableCollection(Of clsDirectoryObject)
            users = searcher.SearchSync(New clsDirectoryObject(domain.DefaultNamingContext, domain),
                        New clsFilter("(&(objectCategory=person)(objectClass=user)(!(objectClass=inetOrgPerson))((userPrincipalName=" & template(starredData) & "@*)))"),
                        SearchScope.Subtree,
                        {"objectCategory", "objectClass", "userPrincipalName"})

            Dim dummy As New List(Of String)
            For Each obj As clsDirectoryObject In users
                dummy.Add(LCase(obj.userPrincipalNameName))
            Next

            For I As Integer = 1 To dummy.Count + 1
                Dim integerData = New With {.displayname = displayname, .n = I}
                Dim u As String = template(integerData)
                If Not dummy.Contains(u) Then
                    result.Add(u)
                    Exit For
                End If
            Next
        Next

        Return result
    End Function

    Public Function GetNextDomainComputers(domain As clsDomain) As List(Of String)
        If domain Is Nothing Then Return Nothing

        Dim result As New List(Of String)
        Dim searcher As New clsSearcher

        For Each template In domain.ComputerPatternTemplates
            Dim starredData = New With {.n = "*"}

            Dim computers As ObservableCollection(Of clsDirectoryObject)
            computers = searcher.SearchSync(New clsDirectoryObject(domain.DefaultNamingContext, domain),
                        New clsFilter("(&(objectCategory=computer)(name=" & template(starredData) & "))"),
                        SearchScope.Subtree,
                        {"objectCategory", "objectClass", "name"})

            Dim dummy As New List(Of String)
            For Each obj As clsDirectoryObject In computers
                dummy.Add(LCase(obj.name))
            Next

            Dim count = 0
            For I As Integer = 1 To dummy.Count + 10
                Dim integerData = New With {.n = I}
                Dim c As String = template(integerData)
                If Not dummy.Contains(c) Then
                    result.Add(c)
                    count += 1
                    If count = 10 Then Exit For
                End If
            Next
        Next

        Return result
    End Function

    Public Function GetNextDomainTelephoneNumbers(domain As clsDomain) As ObservableCollection(Of clsTelephoneNumber)
        If domain Is Nothing Then Return Nothing

        Dim result As New ObservableCollection(Of clsTelephoneNumber)
        Dim searcher As New clsSearcher

        For Each pattern As clsTelephoneNumberPattern In domain.TelephoneNumberPattern
            If Not pattern.Range.Contains("-") Then Continue For
            Dim numstart As Long = 0
            Dim numend As Long = 0
            If Not Long.TryParse(pattern.Range.Split({"-"}, 2, StringSplitOptions.RemoveEmptyEntries)(0), numstart) Or
               Not Long.TryParse(pattern.Range.Split({"-"}, 2, StringSplitOptions.RemoveEmptyEntries)(1), numend) Then Continue For

            Dim starredData = New With {.n = "*"}

            Dim telephonenumbers As ObservableCollection(Of clsDirectoryObject)
            telephonenumbers = searcher.SearchSync(New clsDirectoryObject(domain.DefaultNamingContext, domain),
                        New clsFilter("(&(objectCategory=person)(!(objectClass=inetOrgPerson))(!(UserAccountControl:1.2.840.113556.1.4.803:=2))(telephoneNumber=" & pattern.Template(starredData) & "))"),
                        SearchScope.Subtree,
                        {"objectCategory", "objectClass", "userAccountControl", "telephoneNumber"})

            Dim dummy As New List(Of String)
            For Each obj As clsDirectoryObject In telephonenumbers
                dummy.Add(LCase(obj.telephoneNumber))
            Next

            For I As Long = numstart To numend
                Dim integerData = New With {.n = I}
                Dim t As String = pattern.Template(integerData)
                If Not dummy.Contains(t) Then
                    result.Add(New clsTelephoneNumber(pattern.Label, t))
                    Exit For
                End If
            Next
        Next

        Return result
    End Function

    Public Function GetNameFromDN(DN As String) As String
        Return DN.Split({","}, StringSplitOptions.RemoveEmptyEntries).First.Split({"="}, StringSplitOptions.RemoveEmptyEntries).Last
    End Function

    Public Function GetParentDNFromDN(DN As String) As String
        Dim eDN As List(Of String) = DN.Split({","}, StringSplitOptions.RemoveEmptyEntries).ToList
        If eDN.Count <= 1 Then Return Nothing
        eDN.RemoveAt(0)
        Return Join(eDN.ToArray, ",")
    End Function

    Public Function GetApplicationIcon(fileName As String) As ImageSource
        Dim ai As System.Drawing.Icon = System.Drawing.Icon.ExtractAssociatedIcon(fileName)
        Return System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(ai.Handle, New Int32Rect(0, 0, ai.Width, ai.Height), BitmapSizeOptions.FromEmptyOptions())
    End Function

    Public Function GetNextUserMailbox(obj As clsDirectoryObject) As String
        If obj Is Nothing Then Return ""
        Dim hbTemplate As Func(Of Object, String) = Handlebars.Compile(obj.Domain.MailboxPattern)
        Dim hbData = New With {.displayname = obj.displayName}
        Return hbTemplate(hbData)
    End Function

    Public Function Transliterate_RU_EN(ByVal text As String) As String
        Dim Russian() As String = {"а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я"}
        Dim English() As String = {"a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "sch", "", "y", "", "e", "yu", "ya"}

        For I As Integer = 0 To Russian.Count - 1
            text = text.Replace(Russian(I), English(I))
            text = text.Replace(UCase(Russian(I)), UCase(English(I)))
        Next

        Return LCase(text)
    End Function

    Public Function BooleanToVisibility(value As Boolean) As Visibility
        If value Then
            Return Visibility.Visible
        Else
            Return Visibility.Collapsed
        End If
    End Function

    Public Function StringToSecureString(current As String) As SecureString
        Dim s = New SecureString()
        For Each c As Char In current.ToCharArray()
            s.AppendChar(c)
        Next
        Return s
    End Function


    Public Function Ping(hostname As String) As PingReply
        Dim pingsender As New Ping
        Dim pingoptions As New PingOptions
        Dim pingtimeout As Integer = 1000
        Dim pingreplytask As Task(Of PingReply) = Nothing
        Dim pingreply As PingReply = Nothing
        Dim pingbuffer() As Byte = Text.Encoding.ASCII.GetBytes(Space(32))
        Dim addresses() As IPAddress

        Try
            addresses = Dns.GetHostAddresses(hostname)
        Catch ex As Exception
            Return Nothing
        End Try

        If addresses.Count = 0 Then Return Nothing

        pingoptions.Ttl = 128
        pingoptions.DontFragment = False
        pingreplytask = Task.Run(Function() pingsender.Send(addresses(0), pingtimeout, pingbuffer, pingoptions))
        pingreplytask.Wait()

        Return pingreplytask.Result
    End Function

    Public Function TraceRoute(hostname As String) As List(Of PingReply)
        Dim pingsender As New Ping
        Dim pingoptions As New PingOptions
        Dim pingtimeout As Integer = 1000
        Dim pingreplytask As Task(Of PingReply) = Nothing
        Dim pingreply As PingReply = Nothing
        Dim pingbuffer() As Byte = Text.Encoding.ASCII.GetBytes(Space(32))
        Dim addresses() As IPAddress

        Try
            addresses = Dns.GetHostAddresses(hostname)
        Catch ex As Exception
            Return New List(Of PingReply)
        End Try

        If addresses.Count = 0 Then Return New List(Of PingReply)

        Dim resultlist As New List(Of PingReply)

        For ttl As Integer = 1 To 128
            pingoptions.Ttl = ttl
            pingoptions.DontFragment = False

            pingreplytask = Task.Run(Function() pingsender.Send(addresses(0), pingtimeout, pingbuffer, pingoptions))
            pingreplytask.Wait()

            resultlist.Add(pingreplytask.Result)
            If pingreplytask.Result.Status = IPStatus.Success Then Exit For
        Next

        Return resultlist
    End Function

    Public Function PortScan(hostname As String, portlist As Dictionary(Of Integer, String)) As Dictionary(Of Integer, Boolean)
        Dim resultlist As New Dictionary(Of Integer, Boolean)
        If hostname Is Nothing Then Return resultlist

        For Each port As Integer In portlist.Keys
            Dim tcpClient = New TcpClient()
            Dim connectionTask = tcpClient.ConnectAsync(hostname, port).ContinueWith(Function(tsk) If(tsk.IsFaulted, Nothing, tcpClient))
            Dim timeoutTask = Task.Delay(1000).ContinueWith(Of TcpClient)(Function(tsk) Nothing)
            Dim resultTask = Task.WhenAny(connectionTask, timeoutTask).Unwrap()

            resultTask.Wait()
            Dim resultTcpClient = resultTask.Result
            If resultTcpClient IsNot Nothing Then
                resultlist.Add(port, resultTcpClient.Connected)
                resultTcpClient.Close()
            Else
                resultlist.Add(port, False)
            End If
        Next

        Return resultlist
    End Function

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

    Public Function CountWords(str As String) As Integer
        If String.IsNullOrEmpty(str) Then Return 0
        Return str.Split({" "}, StringSplitOptions.RemoveEmptyEntries).Count
    End Function

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

    Public Function FindVisualChild(Of T As DependencyObject)(ByVal parent As Object) As T
        Dim queue = New Queue(Of DependencyObject)()
        queue.Enqueue(parent)
        While queue.Count > 0
            Dim child As DependencyObject = queue.Dequeue()
            If TypeOf child Is T Then
                Return child
            End If

            For I As Integer = 0 To VisualTreeHelper.GetChildrenCount(child) - 1
                queue.Enqueue(VisualTreeHelper.GetChild(child, I))
            Next
        End While
        Return Nothing
    End Function

    Public Sub ApplicationDeactivate()
        Dim w As New wndAboutDonate
        w.Show()
        applicationdeactivating = True
        'Application.Current.Shutdown()
    End Sub

    Public Sub Donate()
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=6YFL9PWPKYHWN&lc=GB&item_name=ADTools&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_SM%2egif%3aNonHosted")
    End Sub

End Module
