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
Imports Squirrel

Module mdlTools

    Public regApplication As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\" & My.Application.Info.AssemblyName)
    Public regDomains As RegistryKey = regApplication.CreateSubKey("Domains")
    Public regPreferences As RegistryKey = regApplication.CreateSubKey("Preferences")

    Public updateUrl As String = ""

    Public preferences As clsPreferences
    Public domains As New ObservableCollection(Of clsDomain)

    Public applicationdeactivating As Boolean = False

    Public ClipboardBuffer As clsDirectoryObject()
    Public ClipboardAction As enmClipboardAction

    Public Const OBJECT_DUALPANEL_MINWIDTH As Integer = 610

    Public columnsDefault As New ObservableCollection(Of clsViewColumnInfo) From {
        New clsViewColumnInfo("Имя", New List(Of String) From {"name", "description"}, 1, 220),
        New clsViewColumnInfo("Имя входа", New List(Of String) From {"userPrincipalName", "distinguishedNameFormated"}, 2, 450),
        New clsViewColumnInfo("Телефон", New List(Of String) From {"telephoneNumber", "physicalDeliveryOfficeName"}, 3, 100),
        New clsViewColumnInfo("Место работы", New List(Of String) From {"title", "department", "company"}, 4, 300),
        New clsViewColumnInfo("Основной адрес", New List(Of String) From {"mail"}, 5, 170),
        New clsViewColumnInfo("Объект", New List(Of String) From {"whenCreatedFormated", "lastLogonFormated", "accountExpiresFormated"}, 6, 150),
        New clsViewColumnInfo("Пароль", New List(Of String) From {"pwdLastSetFormated", "passwordExpiresFormated"}, 7, 150)}

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

    Public Sub checkApplicationUpdates()
        Task.Run(
            Async Function()
                ServicePointManager.Expect100Continue = True
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

                Dim gitHubManager As UpdateManager = Nothing
                Try
                    gitHubManager = Await UpdateManager.GitHubUpdateManager("https://github.com/ramer/ADTools")
                Catch ex As Exception

                End Try

                If gitHubManager Is Nothing Then Return Nothing

                Try
                    Dim release = Await gitHubManager.UpdateApp()
                    If release IsNot Nothing Then
                        Debug.WriteLine(release.Version)
                    End If
                Catch ex As Exception

                Finally
                    gitHubManager.Dispose()
                End Try

                Return Nothing
            End Function
        )
    End Sub

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
        newstyle.Setters.Add(New Setter(VirtualizingStackPanel.ScrollUnitProperty, ScrollUnit.Pixel))
        newstyle.Setters.Add(New Setter(KeyboardNavigation.DirectionalNavigationProperty, KeyboardNavigationMode.None))

        Dim gridview As New GridView

        gridview.Columns.Add(CreateViewDetailsStyleColumn(New clsViewColumnInfo("⬕", New List(Of String) From {"StatusImage"}, 0, 60)))
        For Each columninfo As clsViewColumnInfo In preferences.Columns
            gridview.Columns.Add(CreateViewDetailsStyleColumn(columninfo))
        Next

        newstyle.Setters.Add(New Setter(ListView.ViewProperty, gridview))

        Return newstyle
    End Function

    Public Function CreateViewDetailsStyleColumn(columninfo As clsViewColumnInfo) As GridViewColumn
        Dim column As New GridViewColumn()
        column.Header = columninfo.Header
        If columninfo.Attributes.Count > 0 Then column.SetValue(clsSorter.PropertyNameProperty, columninfo.Attributes(0))
        column.Width = If(columninfo.Width > 0, columninfo.Width, Double.NaN)
        Dim panel As New FrameworkElementFactory(GetType(VirtualizingStackPanel))
        panel.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center)
        panel.SetValue(FrameworkElement.MarginProperty, New Thickness(0))

        Dim firstline As Boolean = True
        For Each attr As String In columninfo.Attributes
            Dim bind As New Binding(attr) With {.Mode = BindingMode.OneWay, .Converter = New ConverterDataToUIElement, .ConverterParameter = attr}

            Dim container As New FrameworkElementFactory(GetType(ContentControl))
            If firstline Then
                firstline = False
                container.SetValue(TextBlock.FontWeightProperty, FontWeights.Medium)
                'column.SetValue(DataGridColumn.SortMemberPathProperty, attr.Name)
            Else
                container.SetValue(TextBlock.FontWeightProperty, FontWeights.Light)
            End If

            container.SetBinding(ContentControl.ContentProperty, bind)
            container.SetValue(FrameworkElement.ToolTipProperty, attr)
            container.SetValue(FrameworkElement.MaxHeightProperty, 48.0)
            panel.AppendChild(container)
        Next

        Dim template As New DataTemplate()
        template.VisualTree = panel
        column.CellTemplate = template

        Return column
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
