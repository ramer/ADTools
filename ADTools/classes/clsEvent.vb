Imports System.Management

'https://habrahabr.ru/company/netwrix/blog/148501/

Public Class clsEvent
    Private _category As UShort
    Private _categorystring As String
    Private _computername As String
    Private _data As UShort()
    Private _eventCode As UShort
    Private _eventidentifier As UInteger
    Private _eventtype As UShort
    Private _insertionstring As String()
    Private _logfile As String
    Private _message As String
    Private _recordnumber As UInteger
    Private _sourcename As String
    Private _timegenerated As Date
    Private _timewritten As Date
    Private _type As String
    Private _user As String
    Private _messageaccountname As String
    Private _messageaccountdomain As String
    Private _messagelogontype As String
    Private _messagesourceaddress As String

    Public Const EVT_SECTION_SUBJECT = 1
    Public Const EVT_SECTION_LOGONTYPE = 2
    Public Const EVT_SECTION_NEWLOGON = 3
    Public Const EVT_SECTION_PROCESSINFORMATION = 4
    Public Const EVT_SECTION_NETWORKINFORMATION = 5
    Public Const EVT_SECTION_DETAILEDAUTHENTICATIONINFORMATION = 6
    Public Const EVT_SECTION_ACCOUNTFORWHITCHLOGONFAILED = 7
    Public Const EVT_SECTION_ACCOUNTTHATWASLOCKEDOUT = 8
    Public Const EVT_SECTION_ACCOUNTINFORMATION = 9

    Sub New()

    End Sub

    Sub New(evt As ManagementObject)
        _category = evt("Category")
        _categorystring = evt("CategoryString")
        _computername = evt("ComputerName")
        _data = evt("Data")
        _eventCode = evt("EventCode")
        _eventidentifier = evt("EventIdentifier")
        _eventtype = evt("EventType")
        _logfile = evt("Logfile")
        _message = evt("Message")
        _recordnumber = evt("RecordNumber")
        _sourcename = evt("SourceName")
        _timegenerated = ManagementDateTimeConverter.ToDateTime(evt("TimeGenerated"))
        _timewritten = ManagementDateTimeConverter.ToDateTime(evt("TimeWritten"))
        _type = evt("Type")
        _user = evt("User")

        ParseMessage(evt("Message"))
    End Sub

    Public ReadOnly Property Image() As BitmapImage
        Get
            Dim _image As String = ""

            Select Case EventCode
                Case 4624
                    _image = "images/ok.png"
                Case 4634
                    _image = "images/ok.png"
                Case 4625
                    _image = "images/warning.png"
                Case 4740
                    _image = "images/warning.png"
                Case 4771
                    _image = "images/warning.png"
                Case 4776
                    _image = "images/warning.png"
                Case Else
                    _image = "images/puzzle.png"
            End Select

            Return New BitmapImage(New Uri("pack://application:,,,/" & _image))
        End Get
    End Property

    Public Property Category As UShort
        Get
            Return _category
        End Get
        Set(value As UShort)
            _category = value
        End Set
    End Property

    Public Property CategoryString As String
        Get
            Return _categorystring
        End Get
        Set(value As String)
            _categorystring = value
        End Set
    End Property

    Public Property ComputerName As String
        Get
            Return _computername
        End Get
        Set(value As String)
            _computername = value
        End Set
    End Property

    Public Property Data As UShort()
        Get
            Return _data
        End Get
        Set(value As UShort())
            _data = value
        End Set
    End Property

    Public Property EventCode As UShort
        Get
            Return _eventCode
        End Get
        Set(value As UShort)
            _eventCode = value
        End Set
    End Property

    Public Property EventIdentifier As UInteger
        Get
            Return _eventidentifier
        End Get
        Set(value As UInteger)
            _eventidentifier = value
        End Set
    End Property

    Public Property EventType As UShort
        Get
            Return _eventtype
        End Get
        Set(value As UShort)
            _eventtype = value
        End Set
    End Property

    Public Property InsertionString As String()
        Get
            Return _insertionstring
        End Get
        Set(value As String())
            _insertionstring = value
        End Set
    End Property

    Public Property Logfile As String
        Get
            Return _logfile
        End Get
        Set(value As String)
            _logfile = value
        End Set
    End Property

    Public Property Message As String
        Get
            Return _message
        End Get
        Set(value As String)
            _message = value
        End Set
    End Property

    Public Property RecordNumber As UInteger
        Get
            Return _recordnumber
        End Get
        Set(value As UInteger)
            _recordnumber = value
        End Set
    End Property

    Public Property SourceName As String
        Get
            Return _sourcename
        End Get
        Set(value As String)
            _sourcename = value
        End Set
    End Property

    Public Property TimeGenerated As Date
        Get
            Return _timegenerated
        End Get
        Set(value As Date)
            _timegenerated = value
        End Set
    End Property

    Public Property TimeWritten As Date
        Get
            Return _timewritten
        End Get
        Set(value As Date)
            _timewritten = value
        End Set
    End Property

    Public Property Type As String
        Get
            Return _type
        End Get
        Set(value As String)
            _type = value
        End Set
    End Property

    Public Property User As String
        Get
            Return _user
        End Get
        Set(value As String)
            _user = value
        End Set
    End Property

    Private Sub ParseMessage(Message As String)
        If Message = "" Or Message Is Nothing Then Exit Sub
        Message = Replace(Message, vbTab, "")

        Dim lines As String() = Message.Split({vbCrLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
        lines = lines.Select(Function(x As String) Trim(x)).ToArray

        Dim section As Integer
        For Each line As String In lines
            Dim parameter As String = Nothing
            Dim value As String = Nothing
            Dim linearr As String() = line.Split({":"}, StringSplitOptions.RemoveEmptyEntries)
            If linearr.Count >= 1 Then
                parameter = linearr(0)
            End If
            If linearr.Count >= 2 Then
                value = linearr(1)
            End If
            If parameter = "Субъект" Or parameter = "Subject" Then section = EVT_SECTION_SUBJECT
            If parameter = "Учетная запись, которой не удалось выполнить вход" Or parameter = "Account For Which Logon Failed" Then section = EVT_SECTION_ACCOUNTFORWHITCHLOGONFAILED
            If parameter = "Account That Was Locked Out" Then section = EVT_SECTION_ACCOUNTTHATWASLOCKEDOUT
            If parameter = "Account Information" Then section = EVT_SECTION_ACCOUNTINFORMATION
            If parameter = "Сведения о входе" Or parameter = "Тип входа" Or parameter = "Logon Type" Then section = EVT_SECTION_LOGONTYPE
            If parameter = "Новый вход" Or parameter = "New Logon" Then section = EVT_SECTION_NEWLOGON
            If parameter = "Сведения о процессе" Or parameter = "Process Information" Then section = EVT_SECTION_PROCESSINFORMATION
            If parameter = "Сведения о сети" Or parameter = "Network Information" Then section = EVT_SECTION_NETWORKINFORMATION
            If parameter = "Подробные сведения о проверке подлинности" Or parameter = "Detailed Authentication Information" Then section = EVT_SECTION_DETAILEDAUTHENTICATIONINFORMATION

            Select Case EventCode
                Case 4624
                    If section = EVT_SECTION_LOGONTYPE Then
                        If parameter = "Тип входа" Or parameter = "Logon Type" Then MessageLogonType = GetLogonTypeString(value)
                    End If
                    If section = EVT_SECTION_NEWLOGON Then
                        If parameter = "Имя учетной записи" Or parameter = "Account Name" Then _messageaccountname = value
                        If parameter = "Домен учетной записи" Or parameter = "Account Domain" Then _messageaccountdomain = value
                    End If
                    If section = EVT_SECTION_NETWORKINFORMATION Then
                        If parameter = "Сетевой адрес источника" Or parameter = "Source Network Address" Then _messagesourceaddress = value
                    End If
                Case 4634
                    If section = EVT_SECTION_LOGONTYPE Then
                        If parameter = "Тип входа" Or parameter = "Logon Type" Then MessageLogonType = GetLogonTypeString(value)
                    End If
                    If section = EVT_SECTION_SUBJECT Then
                        If parameter = "Имя учетной записи" Or parameter = "Account Name" Then _messageaccountname = value
                        If parameter = "Домен учетной записи" Or parameter = "Account Domain" Then _messageaccountdomain = value
                    End If
                    If section = EVT_SECTION_NETWORKINFORMATION Then
                        If parameter = "Сетевой адрес источника" Or parameter = "Source Network Address" Then _messagesourceaddress = value
                    End If
                Case 4625
                    If section = EVT_SECTION_LOGONTYPE Then
                        If parameter = "Тип входа" Or parameter = "Logon Type" Then MessageLogonType = GetLogonTypeString(value)
                    End If
                    If section = EVT_SECTION_ACCOUNTFORWHITCHLOGONFAILED Then
                        If parameter = "Имя учетной записи" Or parameter = "Account Name" Then _messageaccountname = value
                        If parameter = "Домен учетной записи" Or parameter = "Account Domain" Then _messageaccountdomain = value
                    End If
                    If section = EVT_SECTION_NETWORKINFORMATION Then
                        If parameter = "Сетевой адрес источника" Or parameter = "Source Network Address" Then _messagesourceaddress = value
                    End If
                Case 4740
                    If section = EVT_SECTION_SUBJECT Then
                        If parameter = "Домен учетной записи" Or parameter = "Account Domain" Then _messageaccountdomain = value
                    End If
                    If section = EVT_SECTION_ACCOUNTTHATWASLOCKEDOUT Then
                        If parameter = "Имя учетной записи" Or parameter = "Account Name" Then _messageaccountname = value
                    End If
                Case 4771
                    If section = EVT_SECTION_ACCOUNTINFORMATION Then
                        If parameter = "Имя учетной записи" Or parameter = "Account Name" Then _messageaccountname = value
                    End If
                    If section = EVT_SECTION_NETWORKINFORMATION Then
                        If parameter = "Адрес клиента" Or parameter = "Client Address" Then _messagesourceaddress = value
                    End If
                Case 4776
                    If parameter = "Logon Account" Then _messageaccountname = value
                    If parameter = "Source Workstation" Then _messagesourceaddress = value
            End Select
        Next

    End Sub

    Private Function GetLogonTypeString(lt As String) As String
        Select Case lt
            Case 2 : Return My.Resources.str_LogonTypeInteractive
            Case 3 : Return My.Resources.str_LogonTypeNetwork
            Case 4 : Return My.Resources.str_LogonTypeBatch
            Case 5 : Return My.Resources.str_LogonTypeService
            Case 7 : Return My.Resources.str_LogonTypeUnlock
            Case 8 : Return My.Resources.str_LogonTypeNetworkCleartextPassword
            Case 9 : Return My.Resources.str_LogonTypeNewCredentials
            Case 10 : Return My.Resources.str_LogonTypeRemoteInteractive
            Case 11 : Return My.Resources.str_LogonTypeCachedInteractive
            Case Else : Return ""
        End Select

        '2 — Интерактивный (вход с клавиатуры или экрана системы)
        '3 — Сетевой (например, подключение к общей папке на этом компьютере из любого места в сети или IIS вход — Никогда не заходил 528 на Windows Server 2000 и выше. См. событие 540)
        '4 — Пакет (batch) (например, запланированная задача)
        '5 — Служба (Запуск службы)
        '7 — Разблокировка (например, необслуживаемая рабочая станция с защищенным паролем скринсейвером) 
        '8 — NetworkCleartext (Вход с полномочиями (credentials), отправленными в виде простого текст. Часто обозначает вход в IIS с “базовой аутентификацией”) 
        '9 — NewCredentials
        '10 — RemoteInteractive (Терминальные службы, Удаленный рабочий стол или удаленный помощник) 
        '11 — CachedInteractive (вход с кешированными доменными полномочиями, например, вход на рабочую станцию, которая находится не в сети) 
    End Function

    Public Property MessageAccountName As String
        Get
            Return _messageaccountname
        End Get
        Set(value As String)
            _messageaccountname = value
        End Set
    End Property

    Public Property MessageAccountDomain As String
        Get
            Return _messageaccountdomain
        End Get
        Set(value As String)
            _messageaccountdomain = value
        End Set
    End Property

    Public Property MessageLogonType As String
        Get
            Return _messagelogontype
        End Get
        Set(value As String)
            _messagelogontype = value
        End Set
    End Property

    Public Property MessageSourceAddress As String
        Get
            Return _messagesourceaddress
        End Get
        Set(value As String)
            _messagesourceaddress = value
        End Set
    End Property

End Class
