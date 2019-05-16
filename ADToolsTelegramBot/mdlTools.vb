Imports System.Collections.ObjectModel
Imports System.Reflection
Imports Microsoft.Win32
Imports IRegisty
Imports CredentialManagement
Imports System.IO
Imports Telegram.Bot.Types.Enums

Module mdlTools

    Public regADToolsApplication As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADTools")
    Public regDomains As RegistryKey = regADToolsApplication.CreateSubKey("Domains")

    Public domains As New ObservableCollection(Of clsDomain)

    Public LogFile As String = "ADToolsTelegramBot.log"
    Public nl As String = Environment.NewLine
    Public dnl As String = Environment.NewLine & Environment.NewLine

    Public TelegramUsername As String
    Public TelegramAPIKey As String
    Public TelegramUseProxy As Boolean
    Public TelegramProxyAddress As String
    Public TelegramProxyPort As Integer

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

        Dim regADToolsTelegramBot As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADToolsTelegramBot")
        Try
            TelegramUseProxy = Convert.ToBoolean(regADToolsTelegramBot.GetValue("UseProxy", False))
        Catch
        End Try
        TelegramProxyAddress = regADToolsTelegramBot.GetValue("ProxyAddress", "")
        TelegramProxyPort = Convert.ToInt32(regADToolsTelegramBot.GetValue("ProxyPort", 0))

        If String.IsNullOrEmpty(TelegramUsername) Or String.IsNullOrEmpty(TelegramAPIKey) Then Application.Current.Shutdown()

        If TelegramUseProxy Then
            Dim BotWebProxy As New Net.WebProxy(TelegramProxyAddress, TelegramProxyPort)
            Bot = New Telegram.Bot.TelegramBotClient(TelegramAPIKey, BotWebProxy)
        Else
            Bot = New Telegram.Bot.TelegramBotClient(TelegramAPIKey)
        End If

    End Sub

    Public Sub StartTelegramUpdater()
        Bot.StartReceiving({UpdateType.Message, UpdateType.InlineQuery, UpdateType.CallbackQuery})
    End Sub

    Public Sub initializeDomains(Optional waitInit As Boolean = True)
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

    Public Sub ThrowException(ByVal ex As Exception, ByVal Procedure As String)
        Dim st As String = ex.StackTrace
        Do While ex IsNot Nothing
            Log(Procedure & " - " & ex.Message)
            ex = ex.InnerException
        Loop
        Log(st)
    End Sub

    Public Sub ThrowCustomException(msg As String)
        Log(msg)
    End Sub

    Public Sub Log(msg As String)
        msg = msg.Replace(Environment.NewLine, " /n ")
        Try
            Using sw As StreamWriter = File.AppendText(Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), LogFile))
                sw.WriteLine(Now() & " " & msg)
                Debug.Print(Now() & " " & msg)
            End Using
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub

    Public Sub ShowWrongMemberMessage()
        ThrowCustomException("Глобальная группа может быть членом другой глобальной группы, универсальной группы или локальной группы домена." &
               "Универсальная группа может быть членом другой универсальной группы или локальной группы домена, но не может быть членом глобальной группы." &
               "Локальная группа домена может быть членом только другой локальной группы домена." &
               "Локальную группу домена можно преобразовать в универсальную группу лишь в том случае, если эта локальная группа домена не содержит других членов локальной группы домена. Локальная группа домена не может быть членом универсальной группы." &
               "Глобальную группу можно преобразовать в универсальную лишь в том случае, если эта глобальная группа не входит в состав другой глобальной группы." &
               "Универсальная группа не может быть членом глобальной группы.")
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
