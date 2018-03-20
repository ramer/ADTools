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

    Public SIP As clsSIP

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

    Public Sub ThrowSIPInformation(request As LumiSoft.Net.SIP.Stack.SIP_Request)
        Dim displayName As String = Encoding.UTF8.GetString(Encoding.GetEncoding(1251).GetBytes(request.From.Address.DisplayName))
        Dim telephoneNumber As String = request.From.Address.Uri.Value.Split({"@"}, StringSplitOptions.RemoveEmptyEntries).First
        Dim data As String = Encoding.UTF8.GetString(request.Data)

        Dim w As wndPopup = Nothing
        For Each wnd In Application.Current.Windows
            If TypeOf wnd Is wndPopup Then w = wnd : Exit For
        Next
        If w Is Nothing Then w = New wndPopup

        w.lbCalls.Items.Add(New clsCall(displayName, telephoneNumber, data))
        w.Show()
        w.Activate()
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
