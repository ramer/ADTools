Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Drawing
Imports System.Globalization
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Markup
Imports System.Windows.Threading
Imports Microsoft.VisualBasic.ApplicationServices
Imports IPrompt.VisualBasic

Module mdlMain

    <STAThread>
    Sub Main()
        Try
            Dim manager As New ADToolsApplicationInstanceManager()
            manager.Run({""})
        Catch ex As Exception
            IMsgBox(ex.Message & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "Sub Main()")
        End Try
    End Sub

End Module

Public Class ADToolsApplicationInstanceManager
    Inherits WindowsFormsApplicationBase
    Private app As ADToolsApplication

    Public Sub New()
        Me.IsSingleInstance = True
    End Sub

    Protected Overrides Function OnStartup(e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) As Boolean
        ' First time app is launched
        Try
            app = New ADToolsApplication()
            app.Run()
            Return False
        Catch ex As Exception
            IMsgBox(ex.Message & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "OnStartup")
            Return False
        End Try
    End Function

    Protected Overrides Sub OnStartupNextInstance(eventArgs As StartupNextInstanceEventArgs)
        Try
            MyBase.OnStartupNextInstance(eventArgs)
            app.Activate()
        Catch ex As Exception
            IMsgBox(ex.Message & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "OnStartupNextInstance")
        End Try
    End Sub
End Class

Public Class ADToolsApplication
    Inherits Application

    Public Shared WithEvents nicon As New NotifyIcon
    Public Shared ctxmenu As New ContextMenu({New MenuItem(My.Resources.wndMain_mnuFile_Exit, AddressOf ni_ctxmenuExit)})

    Public Shared WithEvents tsocLog As New clsThreadSafeObservableCollection(Of clsLog)
    Public Shared WithEvents tsocErrorLog As New clsThreadSafeObservableCollection(Of clsErrorLog)
    Public Shared ocGlobalSearchHistory As New ObservableCollection(Of String)

    Protected Overrides Sub OnStartup(e As Windows.StartupEventArgs)
        MyBase.OnStartup(e)
        'Try
        ' unhandled exception handler
        'AddHandler Dispatcher.UnhandledException, AddressOf Dispatcher_UnhandledException
        'AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf AppDomain_CurrentDomain_UnhandledException

        ' notify icon initialization
        nicon.Icon = New Icon(Application.GetResourceStream(New Uri("images/app.ico", UriKind.Relative)).Stream)
            nicon.Text = My.Application.Info.AssemblyName
            nicon.ContextMenu = ctxmenu
            nicon.Visible = True

            ' command line args
            Dim commandlineargs() As String = Environment.GetCommandLineArgs()
            Dim minimizedstart As Boolean = False
            For Each cl In commandlineargs
                If cl = "-minimized" Or cl = "/minimized" Then
                    minimizedstart = True
                End If
            Next

            If Not minimizedstart Then
                ' loading splash screen
                Dim splash As New SplashScreen("images/splash.png")
                splash.Show(True, True)
            End If

            ' setting localization
            FrameworkElement.LanguageProperty.OverrideMetadata(GetType(FrameworkElement), New FrameworkPropertyMetadata(XmlLanguage.GetLanguage(CultureInfo.CurrentCulture.IetfLanguageTag)))

            '    ' preferences setup
            initializePreferences()

            ' domains setup
            initializeDomains()

            '    ' register SIP
            '    initializeSIP()

            '    ' start Redmine monitoring
            '    initializeRedmine()

            If Not minimizedstart Then
                ' loading main form
                wndMainActivate()
            End If

        'Catch ex As Exception
        '    IMsgBox(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "Application.OnStartup")
        'End Try
    End Sub

    Public Sub Activate()
        wndMainActivate()
    End Sub

    Protected Overrides Sub OnExit(e As ExitEventArgs)
        MyBase.OnExit(e)
        Try
            ' notify icon deinitialization
            nicon.Visible = False

            '    ' save preferences
            'deinitializePreferences()

            '    ' unregister SIP
            '    deinitializeSIP()

            '    ' stop Redmine monitoring
            '    deinitializeRedmine()

            '    ' removing unhandled exception handler
            '    RemoveHandler Dispatcher.UnhandledException, AddressOf Dispatcher_UnhandledException
        Catch ex As Exception

        End Try
    End Sub

    Sub New()
        InitializeComponent()
    End Sub


    Private Shared Sub ni_ctxmenuExit(sender As Object, e As EventArgs)
        Current.Shutdown()
    End Sub

    Private Shared Sub ni_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles nicon.MouseClick
        If e.Button = Forms.MouseButtons.Left Then
            wndMainActivate()
        End If
    End Sub

    'Private Shared Sub ni_BalloonTipClicked(sender As Object, e As EventArgs) Handles nicon.BalloonTipClicked
    '    If nicon.Tag IsNot Nothing Then

    '        If TypeOf (nicon.Tag) Is LumiSoft.Net.SIP.Stack.SIP_Request Then ' IncomingCall
    '            Dim w As New wndIncomingCall

    '            Dim request = CType(nicon.Tag, LumiSoft.Net.SIP.Stack.SIP_Request)
    '            Dim displayname As String = Encoding.UTF8.GetString(Encoding.GetEncoding(1251).GetBytes(request.From.Address.DisplayName))
    '            Dim uri() As String = request.From.Address.Uri.Value.Split({"@"}, StringSplitOptions.RemoveEmptyEntries)

    '            w.rnDisplayName.Text = displayname

    '            If uri.Count > 0 Then
    '                w.rnURI.Text = uri(0)
    '                w.Search("""*" & uri(0) & """")
    '            Else
    '                w.rnURI.Text = ""
    '            End If

    '            w.Show()
    '            w.Activate()
    '            w.Topmost = True
    '            w.Topmost = False
    '        End If

    '        If TypeOf (nicon.Tag) Is Issue Then ' IncomingCall
    '            Dim w As New wndRedmineIssue

    '            Dim issue = CType(nicon.Tag, Issue)

    '            w.rnAuthor.Text = issue.Author.Name
    '            w.rnId.Text = issue.Id
    '            If issue.CreatedOn.HasValue Then w.tblckTimeStamp.Text = issue.CreatedOn.Value.ToString
    '            w.tblckSubject.Text = issue.Subject
    '            w.wbDescription.NavigateToString("<head><meta http-equiv='Content-Type' content='text/html;charset=UTF-8'></head>" & issue.Description)
    '            w.Show()
    '            w.Activate()
    '            w.Topmost = True
    '            w.Topmost = False
    '        End If

    '    End If
    'End Sub

    Public Shared Sub wndMainActivate()
        Try
            Dim w As wndMain

            If Application.Current.Windows.Count > 0 Then
                w = Nothing
                For Each wnd In Application.Current.Windows
                    If GetType(wndMain) Is wnd.GetType Then
                        w = wnd

                        w.Show()
                        w.Activate()

                        If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal

                        w.Topmost = True
                        w.Topmost = False
                    End If
                Next
                If w Is Nothing Then w = New wndMain
            Else
                w = New wndMain

                w.Show()
                w.Activate()

                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal

                w.Topmost = True
                w.Topmost = False
            End If
        Catch ex As Exception
            'ThrowException(ex, "wndMainActivate")
        End Try
    End Sub

    Public Shared Sub Dispatcher_UnhandledException(ByVal sender As Object, ByVal e As DispatcherUnhandledExceptionEventArgs)
        ThrowException(e.Exception, "Необработанное исключение")
        e.Handled = True
    End Sub

    Public Shared Sub AppDomain_CurrentDomain_UnhandledException(sender As Object, e As Object)
        Dim ex = TryCast(e, UnhandledExceptionEventArgs)
        If ex IsNot Nothing Then
            ThrowException(ex.Exception, "Необработанное исключение")
        Else
            ThrowCustomException("Необработанное исключение")
        End If
    End Sub

    Private Shared Sub tsocErrorLog_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles tsocErrorLog.CollectionChanged
        Dim w As wndErrorLog

        For Each wnd As Window In Application.Current.Windows
            If GetType(wndErrorLog) Is wnd.GetType Then
                w = wnd
                w.Show()
                w.Activate()
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Normal
                Exit Sub
            End If
        Next

        w = New wndErrorLog
        w.Show()
    End Sub

End Class
