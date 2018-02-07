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

        ShowPopup(displayName, telephoneNumber)
    End Sub

    Public Sub ShowPopup(displayName As String, telephoneNumber As String)
        Dim popHeaderText = New TextBlock With {.Style = Windows.Application.Current.FindResource("PopupHeaderTextStyle")}
        popHeaderText.Text = "ADToolsSIP"

        Dim popContentText As New TextBlock With {.Style = Windows.Application.Current.FindResource("PopupContentTextStyle")}
        popContentText.Text = String.Format("{0} ({1})", displayName, telephoneNumber)

        Dim pop = New Popup With {.Style = Windows.Application.Current.FindResource("PopupStyle"), .HorizontalOffset = Forms.Screen.PrimaryScreen.WorkingArea.Right - 5, .VerticalOffset = Forms.Screen.PrimaryScreen.WorkingArea.Bottom - 5}
        pop.Child = New HeaderedContentControl With {.Style = Windows.Application.Current.FindResource("PopupContentStyle"), .Content = popContentText, .Header = popHeaderText}

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

        AddHandler pop.MouseLeftButtonDown,
            Sub(sender As Object, e As MouseButtonEventArgs)
                Debug.Print("..\..\ADTools.exe", "-search """ & displayName & "* / *" & telephoneNumber & """")
                Process.Start("..\..\ADTools.exe", "-search """ & displayName & "* / *" & telephoneNumber & """")

                pop.IsOpen = False
                popTimer.Stop()
                popTimer = Nothing
                pop = Nothing
            End Sub

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
