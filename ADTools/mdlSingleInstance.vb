Imports System.IO
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Ipc
Imports System.Runtime.Serialization.Formatters
Imports System.Threading
Imports System.Windows.Threading
Imports System.Security
Imports System.Runtime.InteropServices
Imports System.ComponentModel

Friend Enum WM
    NULL = 0
    CREATE = 1
    DESTROY = 2
    MOVE = 3
    SIZE = 5
    ACTIVATE = 6
    SETFOCUS = 7
    KILLFOCUS = 8
    ENABLE = 10
    SETREDRAW = 11
    SETTEXT = 12
    GETTEXT = 13
    GETTEXTLENGTH = 14
    PAINT = 15
    CLOSE = 16
    QUERYENDSESSION = 17
    QUIT = 18
    QUERYOPEN = 19
    ERASEBKGND = 20
    SYSCOLORCHANGE = 21
    SHOWWINDOW = 24
    ACTIVATEAPP = 28
    SETCURSOR = 32
    MOUSEACTIVATE = 33
    CHILDACTIVATE = 34
    QUEUESYNC = 35
    GETMINMAXINFO = 36
    WINDOWPOSCHANGING = 70
    WINDOWPOSCHANGED = 71
    CONTEXTMENU = 123
    STYLECHANGING = 124
    STYLECHANGED = 125
    DISPLAYCHANGE = 126
    GETICON = 127
    SETICON = 128
    NCCREATE = 129
    NCDESTROY = 130
    NCCALCSIZE = 131
    NCHITTEST = 132
    NCPAINT = 133
    NCACTIVATE = 134
    GETDLGCODE = 135
    SYNCPAINT = 136
    NCMOUSEMOVE = 160
    NCLBUTTONDOWN = 161
    NCLBUTTONUP = 162
    NCLBUTTONDBLCLK = 163
    NCRBUTTONDOWN = 164
    NCRBUTTONUP = 165
    NCRBUTTONDBLCLK = 166
    NCMBUTTONDOWN = 167
    NCMBUTTONUP = 168
    NCMBUTTONDBLCLK = 169
    SYSKEYDOWN = 260
    SYSKEYUP = 261
    SYSCHAR = 262
    SYSDEADCHAR = 263
    COMMAND = 273
    SYSCOMMAND = 274
    MOUSEMOVE = 512
    LBUTTONDOWN = 513
    LBUTTONUP = 514
    LBUTTONDBLCLK = 515
    RBUTTONDOWN = 516
    RBUTTONUP = 517
    RBUTTONDBLCLK = 518
    MBUTTONDOWN = 519
    MBUTTONUP = 520
    MBUTTONDBLCLK = 521
    MOUSEWHEEL = 522
    XBUTTONDOWN = 523
    XBUTTONUP = 524
    XBUTTONDBLCLK = 525
    MOUSEHWHEEL = 526
    CAPTURECHANGED = 533
    ENTERSIZEMOVE = 561
    EXITSIZEMOVE = 562
    IME_SETCONTEXT = 641
    IME_NOTIFY = 642
    IME_CONTROL = 643
    IME_COMPOSITIONFULL = 644
    IME_SELECT = 645
    IME_CHAR = 646
    IME_REQUEST = 648
    IME_KEYDOWN = 656
    IME_KEYUP = 657
    NCMOUSELEAVE = 674
    DWMCOMPOSITIONCHANGED = 798
    DWMNCRENDERINGCHANGED = 799
    DWMCOLORIZATIONCOLORCHANGED = 800
    DWMWINDOWMAXIMIZEDCHANGE = 801
    DWMSENDICONICTHUMBNAIL = 803
    DWMSENDICONICLIVEPREVIEWBITMAP = 806
    USER = 1024
    TRAYMOUSEMESSAGE = 2048
    APP = 32768
End Enum

<SuppressUnmanagedCodeSecurity>
Friend Module NativeMethods

    Public Delegate Function MessageHandler(ByVal uMsg As WM, ByVal wParam As IntPtr, ByVal lParam As IntPtr, <Out> ByRef handled As Boolean) As IntPtr

    <DllImport("shell32.dll", EntryPoint:="CommandLineToArgvW", CharSet:=CharSet.Unicode)>
    Private Function _CommandLineToArgvW(<MarshalAs(UnmanagedType.LPWStr)> ByVal cmdLine As String, <Out> ByRef numArgs As Integer) As IntPtr
    End Function

    <DllImport("kernel32.dll", EntryPoint:="LocalFree", SetLastError:=True)>
    Private Function _LocalFree(ByVal hMem As IntPtr) As IntPtr
    End Function

    Function CommandLineToArgvW(ByVal cmdLine As String) As String()
        Dim argv As IntPtr = IntPtr.Zero
        Try
            Dim numArgs As Integer = 0
            argv = _CommandLineToArgvW(cmdLine, numArgs)
            If argv = IntPtr.Zero Then
                Throw New Win32Exception()
            End If

            Dim result = New String(numArgs - 1) {}
            For i As Integer = 0 To numArgs - 1
                Dim currArg As IntPtr = Marshal.ReadIntPtr(argv, i * Marshal.SizeOf(GetType(IntPtr)))
                result(i) = Marshal.PtrToStringUni(currArg)
            Next

            Return result
        Finally
            Dim p As IntPtr = _LocalFree(argv)
        End Try
    End Function
End Module

Public Interface ISingleInstanceApp

    Function SignalExternalCommandLineArgs(ByVal args As IList(Of String)) As Boolean

End Interface

Public Class SingleInstance(Of TApplication As {Application, ISingleInstanceApp})

    Private Const Delimiter As String = ":"

    Private Const ChannelNameSuffix As String = "SingeInstanceIPCChannel"

    Private Const RemoteServiceName As String = "SingleInstanceApplicationService"

    Private Const IpcProtocol As String = "ipc://"

    Private Shared singleInstanceMutex As Mutex

    Private Shared channel As IpcServerChannel

    Private Shared _commandLineArgs As IList(Of String)

    Public Shared ReadOnly Property CommandLineArgs As IList(Of String)
        Get
            Return _commandLineArgs
        End Get
    End Property

    Public Shared Function InitializeAsFirstInstance(ByVal uniqueName As String) As Boolean
        _commandLineArgs = GetCommandLineArgs(uniqueName)
        Dim applicationIdentifier As String = uniqueName & Environment.UserName
        Dim channelName As String = String.Concat(applicationIdentifier, Delimiter, ChannelNameSuffix)
        Dim firstInstance As Boolean
        singleInstanceMutex = New Mutex(True, applicationIdentifier, firstInstance)

        If firstInstance Then
            CreateRemoteService(channelName)
        Else
            SignalFirstInstance(channelName, CommandLineArgs)
        End If

        Return firstInstance
    End Function

    Public Shared Sub Cleanup()
        If singleInstanceMutex IsNot Nothing Then
            singleInstanceMutex.Close()
            singleInstanceMutex = Nothing
        End If

        If channel IsNot Nothing Then
            ChannelServices.UnregisterChannel(channel)
            channel = Nothing
        End If
    End Sub

    Private Shared Function GetCommandLineArgs(ByVal uniqueApplicationName As String) As IList(Of String)
        Dim args As String() = Nothing
        If AppDomain.CurrentDomain.ActivationContext Is Nothing Then
            args = Environment.GetCommandLineArgs()
        Else
            Dim appFolderPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), uniqueApplicationName)
            Dim cmdLinePath As String = Path.Combine(appFolderPath, "cmdline.txt")
            If File.Exists(cmdLinePath) Then
                Try
                    Using reader As TextReader = New StreamReader(cmdLinePath, System.Text.Encoding.Unicode)
                        args = NativeMethods.CommandLineToArgvW(reader.ReadToEnd())
                    End Using

                    File.Delete(cmdLinePath)
                Catch __unusedIOException1__ As IOException
                End Try
            End If
        End If

        If args Is Nothing Then
            args = New String() {}
        End If

        Return New List(Of String)(args)
    End Function

    Private Shared Sub CreateRemoteService(ByVal channelName As String)
        Dim serverProvider As BinaryServerFormatterSinkProvider = New BinaryServerFormatterSinkProvider()
        serverProvider.TypeFilterLevel = TypeFilterLevel.Full
        Dim props As IDictionary = New Dictionary(Of String, String)()
        props("name") = channelName
        props("portName") = channelName
        props("exclusiveAddressUse") = "false"
        channel = New IpcServerChannel(props, serverProvider)
        ChannelServices.RegisterChannel(channel, True)
        Dim remoteService As IPCRemoteService = New IPCRemoteService()
        RemotingServices.Marshal(remoteService, RemoteServiceName)
    End Sub

    Private Shared Sub SignalFirstInstance(ByVal channelName As String, ByVal args As IList(Of String))
        Dim secondInstanceChannel As IpcClientChannel = New IpcClientChannel()
        ChannelServices.RegisterChannel(secondInstanceChannel, True)
        Dim remotingServiceUrl As String = IpcProtocol & channelName & "/" & RemoteServiceName
        Dim firstInstanceRemoteServiceReference As IPCRemoteService = CType(RemotingServices.Connect(GetType(IPCRemoteService), remotingServiceUrl), IPCRemoteService)
        If firstInstanceRemoteServiceReference IsNot Nothing Then
            firstInstanceRemoteServiceReference.InvokeFirstInstance(args)
        End If
    End Sub

    Private Shared Function ActivateFirstInstanceCallback(ByVal arg As Object) As Object
        Dim args As IList(Of String) = TryCast(arg, IList(Of String))
        ActivateFirstInstance(args)
        Return Nothing
    End Function

    Private Shared Sub ActivateFirstInstance(ByVal args As IList(Of String))
        If Application.Current Is Nothing Then
            Return
        End If

        CType(Application.Current, TApplication).SignalExternalCommandLineArgs(args)
    End Sub

    Private Class IPCRemoteService
        Inherits MarshalByRefObject

        Public Sub InvokeFirstInstance(ByVal args As IList(Of String))
            If Application.Current IsNot Nothing Then
                Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, New DispatcherOperationCallback(AddressOf ActivateFirstInstanceCallback), args)
            End If
        End Sub

        Public Overrides Function InitializeLifetimeService() As Object
            Return Nothing
        End Function
    End Class

End Class
