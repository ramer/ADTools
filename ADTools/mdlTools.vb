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
    Public regPreferences As RegistryKey = regApplication.CreateSubKey("Preferences")
    Public regDomains As RegistryKey = regApplication.CreateSubKey("Domains")

    Public updateUrl As String = ""

    Public preferences As clsPreferences
    Public domains As ObservableCollection(Of clsDomain)

    Public applicationdeactivating As Boolean = False

    Public ClipboardBuffer As clsDirectoryObject()
    Public ClipboardAction As enmClipboardAction

    Public Const OBJECT_DUALPANEL_MINWIDTH As Integer = 610

    Public columnsDefault As New ObservableCollection(Of clsViewColumnInfo) From {
        New clsViewColumnInfo(My.Resources.str_ColumnsDefaultName, New List(Of String) From {"name", "description"}, 220),
        New clsViewColumnInfo(My.Resources.str_ColumnsDefaultLogin, New List(Of String) From {"userPrincipalName", "distinguishedNameFormated"}, 450),
        New clsViewColumnInfo(My.Resources.str_ColumnsDefaultPhone, New List(Of String) From {"telephoneNumber", "physicalDeliveryOfficeName"}, 100),
        New clsViewColumnInfo(My.Resources.str_ColumnsDefaultWorkplace, New List(Of String) From {"title", "department", "company"}, 300),
        New clsViewColumnInfo(My.Resources.str_ColumnsDefaultMail, New List(Of String) From {"mail"}, 170),
        New clsViewColumnInfo(My.Resources.str_ColumnsDefaultObject, New List(Of String) From {"whenCreatedFormated", "lastLogonFormated", "accountExpiresFormated"}, 150),
        New clsViewColumnInfo(My.Resources.str_ColumnsDefaultPassword, New List(Of String) From {"pwdLastSetFormated", "passwordExpiresFormated"}, 150)}

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

    ''' <summary>
    ''' Checks Application updates available on github repository
    ''' </summary>
    Public Sub CheckApplicationUpdates()
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
            End Function)
    End Sub

    ''' <summary>
    ''' Creates preferences instance by deserializing preferences registry key
    ''' </summary>
    Public Sub InitializePreferences()
        preferences = IRegistrySerializer.Deserialize(GetType(clsPreferences), regPreferences)
    End Sub

    ''' <summary>
    ''' Creates domains instance by deserializing domains registry key, initializes each domain
    ''' </summary>
    ''' <param name="waitInit">If True waits for all domains initialization</param>
    Public Sub InitializeDomains(Optional waitInit As Boolean = True)
        domains = IRegistrySerializer.Deserialize(GetType(ObservableCollection(Of clsDomain)), regDomains)

        Task.Run(
            Sub()
                Dim initTasks As New List(Of Task)
                For Each domain In domains
                    initTasks.Add(domain.Initialize)
                Next
                If waitInit Then Task.WaitAll(initTasks.ToArray)
            End Sub).Wait()
    End Sub

    ''' <summary>
    ''' Removes current preferences registry key and serializes current preferences instance to the registry
    ''' </summary>
    Public Sub DeinitializePreferences()
        Array.ForEach(regPreferences.GetSubKeyNames, New Action(Of String)(Sub(p) regPreferences.DeleteSubKeyTree(p, False)))
        IRegistrySerializer.Serialize(preferences, regPreferences)
    End Sub

    ''' <summary>
    ''' Overrides current culture by system, sets handlebars patterns, registers the application URL protocol
    ''' </summary>
    Public Sub InitializeGlobalParameters()
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

        Try
            Dim regApplicationClassRoot As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Classes\" & My.Application.Info.AssemblyName.ToLower)

            If regApplicationClassRoot Is Nothing Then
                regApplicationClassRoot = Registry.CurrentUser.CreateSubKey("SOFTWARE\Classes\" & My.Application.Info.AssemblyName.ToLower)
                regApplicationClassRoot.SetValue(String.Empty, "URL:" & My.Application.Info.AssemblyName.ToLower)
                regApplicationClassRoot.SetValue("URL Protocol", String.Empty)
                regApplicationClassRoot = regApplicationClassRoot.CreateSubKey("shell\open\command")
                regApplicationClassRoot.SetValue(String.Empty, System.Reflection.Assembly.GetExecutingAssembly().Location & " " & "%1")
            End If

            regApplicationClassRoot.Close()
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Creates a new one or uses an existing window to display the page.
    ''' </summary>
    ''' <param name="page">Current page to display</param>
    ''' <param name="singleInstance">True = use existing window, False = create new</param>
    ''' <param name="owner">Owner window for new child window</param>
    ''' <param name="modal">True = modal window, False = can be unfocused</param>
    ''' <returns>NavigationWindow</returns>
    Public Function ShowPage(page As Page, Optional singleInstance As Boolean = False, Optional owner As Window = Nothing, Optional modal As Boolean = False) As NavigationWindow
        If applicationdeactivating Then Return Nothing
        If page Is Nothing Then Return Nothing

        Dim w As NavigationWindow = Nothing

        If preferences.PageInterface Then

            If owner Is Nothing Then
                w = New NavigationWindow
                w.WindowStartupLocation = WindowStartupLocation.CenterScreen
            Else
                w = owner
            End If

            page.WindowWidth = Double.NaN
            page.WindowHeight = Double.NaN
            w.Navigate(page)

            w.Show()

            Return w

        Else

            If owner Is Nothing Then

                w = New NavigationWindow
                w.WindowStartupLocation = WindowStartupLocation.CenterScreen
                w.Navigate(page)
                w.UpdateLayout()

                If modal Then
                    w.ShowDialog()
                Else
                    w.Show()
                End If

                Return w

            Else

                For Each wnd As Window In owner.OwnedWindows
                    If TypeOf wnd Is NavigationWindow AndAlso page.GetType Is wnd.Content.GetType AndAlso TypeOf page Is pgObject AndAlso wnd.Content.CurrentObject Is CType(page, Object).CurrentObject Then
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

                If TypeOf page Is pgObject Then w.Width = 900 : w.Height = 620

                w.Navigate(page)

                If modal Then
                    w.ShowDialog()
                Else
                    w.Show()
                End If

                Return w

            End If

        End If

    End Function

    ''' <summary>
    ''' Сreates the necessary page depending on the class of the object
    ''' </summary>
    ''' <param name="obj">Object to display</param>
    ''' <param name="owner">Owner window for new child window</param>
    ''' <returns></returns>
    Public Function ShowDirectoryObjectProperties(obj As clsDirectoryObject, Optional owner As Window = Nothing) As NavigationWindow
        Dim p As Page

        If obj.SchemaClass = enmDirectoryObjectSchemaClass.User Then
            p = New pgObject(obj)
        ElseIf obj.SchemaClass = enmDirectoryObjectSchemaClass.Contact Then
            p = New pgObject(obj)
        ElseIf obj.SchemaClass = enmDirectoryObjectSchemaClass.Computer Then
            p = New pgObject(obj)
        ElseIf obj.SchemaClass = enmDirectoryObjectSchemaClass.Group Then
            p = New pgObject(obj)
        ElseIf obj.SchemaClass = enmDirectoryObjectSchemaClass.OrganizationalUnit Then
            p = New pgObject(obj)
        Else
            p = New pgObject(obj)
        End If

        Return ShowPage(p, False, owner, False)
    End Function

    ''' <summary>
    ''' Adds an error to the thread-safe list and displays a window
    ''' </summary>
    ''' <param name="ex">Exception to display</param>
    ''' <param name="procedureName">Procedure that caused the error</param>
    Public Sub ThrowException(ByVal ex As Exception, ByVal procedureName As String)
        ADToolsApplication.tsocErrorLog.Add(New clsErrorLog(procedureName,, ex))
    End Sub

    ''' <summary>
    ''' Adds custom error message to the thread-safe list and displays a window
    ''' </summary>
    ''' <param name="Message"></param>
    Public Sub ThrowCustomException(Message As String)
        ADToolsApplication.tsocErrorLog.Add(New clsErrorLog(Message))
    End Sub

    ''' <summary>
    ''' Displays a message when adding groups incorrectly.
    ''' </summary>
    Public Sub ShowWrongMemberMessage()
        IMsgBox(My.Resources.str_WrongGroupMember, vbOKOnly + vbExclamation, My.Resources.str_WrongGroupMemberTitle)
    End Sub

    ''' <summary>
    ''' Receives the next domain username according to the pattern
    ''' </summary>
    ''' <param name="domain">Search domain</param>
    ''' <returns>Username</returns>
    Public Async Function GetNextDomainUserAsync(domain As clsDomain) As Task(Of List(Of String))
        If domain Is Nothing Then Return Nothing

        Dim result As New List(Of String)
        Dim searcher As New clsSearcher

        For Each template In domain.UsernamePatternTemplates
            Dim starredData = New With {.n = "*"}

            Dim users As ObservableCollection(Of clsDirectoryObject)
            users = Await Task.Run(Function() searcher.SearchSync(New clsDirectoryObject(domain.DefaultNamingContext, domain),
                        New clsFilter("(&(objectCategory=person)(objectClass=user)(!(objectClass=inetOrgPerson))((userPrincipalName=" & template(starredData) & "@*)))"),
                        SearchScope.Subtree,
                        {"objectCategory", "objectClass", "userPrincipalName"}))

            Dim dummy As New List(Of String)
            For Each obj As clsDirectoryObject In users
                dummy.Add(LCase(obj.userPrincipalNameName))
            Next

            For I As Integer = 1 To dummy.Count + 1
                Dim integerData = New With {.n = I}
                Dim u As String = template(integerData)
                If Not dummy.Contains(u) Then
                    result.Add(u)
                    Exit For
                End If
            Next
        Next

        Return result
    End Function

    ''' <summary>
    ''' Gets the next domain computer name according to the pattern
    ''' </summary>
    ''' <param name="domain">Search domain</param>
    ''' <returns>Computer name</returns>
    Public Async Function GetNextDomainComputerAsync(domain As clsDomain) As Task(Of List(Of String))
        If domain Is Nothing Then Return Nothing

        Dim result As New List(Of String)
        Dim searcher As New clsSearcher

        For Each template In domain.ComputerPatternTemplates
            Dim starredData = New With {.n = "*"}

            Dim computers As ObservableCollection(Of clsDirectoryObject)
            computers = Await Task.Run(Function() searcher.SearchSync(New clsDirectoryObject(domain.DefaultNamingContext, domain),
                        New clsFilter("(&(objectCategory=computer)(name=" & template(starredData) & "))"),
                        SearchScope.Subtree,
                        {"objectCategory", "objectClass", "name"}))

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

    ''' <summary>
    ''' Gets the next domain telephone number according to the pattern
    ''' </summary>
    ''' <param name="domain">Search domain</param>
    ''' <returns>Telephone number</returns>
    Public Async Function GetNextDomainTelephoneNumberAsync(domain As clsDomain) As Task(Of ObservableCollection(Of clsTelephoneNumber))
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
            telephonenumbers = Await Task.Run(Function() searcher.SearchSync(New clsDirectoryObject(domain.DefaultNamingContext, domain),
                        New clsFilter("(&(objectCategory=person)(!(objectClass=inetOrgPerson))(!(UserAccountControl:1.2.840.113556.1.4.803:=2))(telephoneNumber=" & pattern.Template(starredData) & "))"),
                        SearchScope.Subtree,
                        {"objectCategory", "objectClass", "userAccountControl", "telephoneNumber"}))

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


    ''' <summary>
    ''' Generates new email address according to the pattern
    ''' </summary>
    ''' <param name="obj">User object</param>
    ''' <returns>Email address</returns>
    Public Function GetUserMailbox(obj As clsDirectoryObject) As String
        If obj Is Nothing Then Return ""
        Dim hbTemplate As Func(Of Object, String) = Handlebars.Compile(obj.Domain.MailboxPattern)
        Dim hbData = New With {obj.displayName}
        Return hbTemplate(hbData)
    End Function

    ''' <summary>
    ''' Extracts object name according distinguished name
    ''' </summary>
    ''' <param name="DN">Current distinguished name</param>
    ''' <returns>Object name</returns>
    Public Function GetNameFromDN(DN As String) As String
        Return DN.Split({","}, StringSplitOptions.RemoveEmptyEntries).First.Split({"="}, StringSplitOptions.RemoveEmptyEntries).Last
    End Function

    ''' <summary>
    ''' Excracts parent distinguished name according current distinguished name
    ''' </summary>
    ''' <param name="DN">Current distinguished name</param>
    ''' <returns>Parent distinguished name</returns>
    Public Function GetParentDNFromDN(DN As String) As String
        Dim eDN As List(Of String) = DN.Split({","}, StringSplitOptions.RemoveEmptyEntries).ToList
        If eDN.Count <= 1 Then Return Nothing
        eDN.RemoveAt(0)
        Return Join(eDN.ToArray, ",")
    End Function

    ''' <summary>
    ''' Gets icon from given application
    ''' </summary>
    ''' <param name="fileName">Application executable file</param>
    ''' <returns>Application icon</returns>
    Public Function GetApplicationIcon(fileName As String) As ImageSource
        Dim ai As System.Drawing.Icon = System.Drawing.Icon.ExtractAssociatedIcon(fileName)
        Return Interop.Imaging.CreateBitmapSourceFromHIcon(ai.Handle, New Int32Rect(0, 0, ai.Width, ai.Height), BitmapSizeOptions.FromEmptyOptions())
    End Function

    ''' <summary>
    ''' Translates a string from Russian to English
    ''' </summary>
    ''' <param name="str">Input string</param>
    ''' <returns>Transliterated string</returns>
    Public Function Transliterate_RU_EN(ByVal str As String) As String
        Dim Russian() As String = {"а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я"}
        Dim English() As String = {"a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "sch", "", "y", "", "e", "yu", "ya"}

        For I As Integer = 0 To Russian.Count - 1
            str = str.Replace(Russian(I), English(I))
            str = str.Replace(UCase(Russian(I)), UCase(English(I)))
        Next

        Return LCase(str)
    End Function

    ''' <summary>
    ''' Changes the layout of Russian and English in both directions
    ''' </summary>
    ''' <param name="str">Input string</param>
    ''' <returns>Inverted string layout (RU-EN)</returns>
    Public Function SwitchLayout_EN_RU(ByVal str As String) As String
        Dim English As New List(Of Char) From {"`", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-", "=", "q", "w", "e", "r", "t", "y", "u", "i", "o", "p", "[", "]", "a", "s", "d", "f", "g", "h", "j", "k", "l", ";", "'", "\", "z", "x", "c", "v", "b", "n", "m", ",", ".", "/", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "{", "}", "A", "S", "D", "F", "G", "H", "J", "K", "L", ":", """", "|", "Z", "X", "C", "V", "B", "N", "M", "<", ">", "?"}
        Dim Russian As New List(Of Char) From {"ё", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-", "=", "й", "ц", "у", "к", "е", "н", "г", "ш", "щ", "з", "х", "ъ", "ф", "ы", "в", "а", "п", "р", "о", "л", "д", "ж", "э", "\", "я", "ч", "с", "м", "и", "т", "ь", "б", "ю", ".", "Ё", "!", """", "№", ";", "%", ":", "?", "*", "(", ")", "_", "+", "Й", "Ц", "У", "К", "Е", "Н", "Г", "Ш", "Щ", "З", "Х", "Ъ", "Ф", "Ы", "В", "А", "П", "Р", "О", "Л", "Д", "Ж", "Э", "/", "Я", "Ч", "С", "М", "И", "Т", "Ь", "Б", "Ю", ","}

        Dim chararray As Char() = str

        For I As Integer = 0 To chararray.Length - 1
            Dim ensymbol = English.IndexOf(chararray(I))
            If ensymbol > 0 Then chararray(I) = Russian(ensymbol) : Continue For

            Dim rusymbol = Russian.IndexOf(chararray(I))
            If rusymbol > 0 Then chararray(I) = English(rusymbol)
        Next

        Return New String(chararray)
    End Function

    ''' <summary>
    ''' Coverts Boolean value to visibility
    ''' </summary>
    ''' <param name="value">True = Visible, False = Collapsed</param>
    ''' <returns></returns>
    Public Function BooleanToVisibility(value As Boolean) As Visibility
        Return If(value, Visibility.Visible, Visibility.Collapsed)
    End Function

    ''' <summary>
    ''' Converts string to secure string (password)
    ''' </summary>
    ''' <param name="str">Input string</param>
    ''' <returns>Secure string</returns>
    Public Function StringToSecureString(str As String) As SecureString
        Dim s = New SecureString()
        For Each c As Char In str.ToCharArray()
            s.AppendChar(c)
        Next
        Return s
    End Function

    ''' <summary>
    ''' Synchronously pings the specified host
    ''' </summary>
    ''' <param name="hostname">Host to ping</param>
    ''' <returns>Ping result</returns>
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

    ''' <summary>
    ''' Synchronously traces the route to the specified host.
    ''' </summary>
    ''' <param name="hostname">Host to trace route</param>
    ''' <returns>List of ping results</returns>
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

    ''' <summary>
    ''' Connects to specified TCP ports
    ''' </summary>
    ''' <param name="hostname">Host to connect</param>
    ''' <param name="portlist">Ports to connect</param>
    ''' <returns>Connection results</returns>
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


    ''' <summary>
    ''' Counts words in specified string
    ''' </summary>
    ''' <param name="str">Input string</param>
    ''' <returns>Words count</returns>
    Public Function CountWords(str As String) As Integer
        If String.IsNullOrEmpty(str) Then Return 0
        Return str.Split({" "}, StringSplitOptions.RemoveEmptyEntries).Count
    End Function


    ''' <summary>
    ''' Searches a parent of the specified object in the visual tree. 
    ''' </summary>
    ''' <typeparam name="T">Type of parent object</typeparam>
    ''' <param name="child">Current object</param>
    ''' <param name="until">Top-object, limits search</param>
    ''' <returns>Parent object of T, or Nothing</returns>
    Public Function FindVisualParent(Of T As DependencyObject)(ByVal child As Object, Optional until As DependencyObject = Nothing) As T
        Dim parent As DependencyObject = If(child.Parent IsNot Nothing, child.Parent, VisualTreeHelper.GetParent(child))

        If parent IsNot Nothing Then
            If TypeOf parent Is T Then
                Return parent
            ElseIf parent Is until Then
                Return Nothing
            Else
                Return FindVisualParent(Of T)(parent, until)
            End If
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Searches a child of the specified object in the visual tree
    ''' </summary>
    ''' <typeparam name="T">Type of child object</typeparam>
    ''' <param name="parent">Current object</param>
    ''' <returns>Child object of T, or Nothing</returns>
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

    ''' <summary>
    ''' Shut down the application after a few seconds ;)
    ''' </summary>
    Public Sub ApplicationDeactivate()
        Dim w As New wndAboutDonate
        w.Show()
        applicationdeactivating = True
        'Application.Current.Shutdown()
    End Sub

    ''' <summary>
    ''' Opens URL in default browser
    ''' </summary>
    Public Sub Donate()
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=6YFL9PWPKYHWN&lc=GB&item_name=ADTools&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_SM%2egif%3aNonHosted")
    End Sub

End Module
