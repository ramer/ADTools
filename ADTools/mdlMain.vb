Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Drawing
Imports System.Windows.Threading
Imports Microsoft.VisualBasic.ApplicationServices
Imports IPrompt.VisualBasic

Module mdlMain
    Private app As ADToolsApplication
    Private Const Unique As String = "ADToolsApplication_ZWvxuwLF7Xx3T8eZ"

    <STAThread>
    Sub Main()
        If SingleInstance(Of ADToolsApplication).InitializeAsFirstInstance(Unique) Then
            app = New ADToolsApplication()
            app.Run()
            SingleInstance(Of ADToolsApplication).Cleanup()
        End If
    End Sub

End Module

Public Class ADToolsApplication
    Inherits Application
    Implements ISingleInstanceApp

    Public Shared WithEvents nicon As New Forms.NotifyIcon
    Public Shared ctxmenu As New Forms.ContextMenu({New Forms.MenuItem(My.Resources.ctxmnu_Exit, AddressOf ni_ctxmenuExit)})

    'Public Shared WithEvents tsocLog As New clsThreadSafeObservableCollection(Of clsLog)
    Public Shared WithEvents tsocErrorLog As New clsThreadSafeObservableCollection(Of clsErrorLog)
    Public Shared commandlineargs As String()

    Public Function SignalExternalCommandLineArgs(args As IList(Of String)) As Boolean Implements ISingleInstanceApp.SignalExternalCommandLineArgs
        commandlineargs = args.ToArray
        Return True
    End Function

    Protected Overrides Sub OnStartup(e As Windows.StartupEventArgs)
        MyBase.OnStartup(e)
        Try
            ' unhandled exception handler
            AddHandler Dispatcher.UnhandledException, AddressOf Dispatcher_UnhandledException
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf AppDomain_CurrentDomain_UnhandledException

            Dim startminimized As Boolean = False
            Dim startwithsearch As Boolean = False
            Dim startsearchstring As String = Nothing
            commandlineargs = Environment.GetCommandLineArgs()
            For i = 0 To commandlineargs.Count - 1
                If commandlineargs(i) = "-minimized" Or commandlineargs(i) = "/minimized" Then startminimized = True
                If (commandlineargs(i) = "-search" Or commandlineargs(i) = "/search") AndAlso i + 1 <= commandlineargs.Count - 1 Then
                    startwithsearch = True
                    startsearchstring = commandlineargs(i + 1)
                End If
            Next

            ' splash screen
            If Not startminimized And Not startwithsearch Then
                Dim splash As New SplashScreen("images/splash.png")
                splash.Show(True, True)
            End If

            ' check for application updates
            checkApplicationUpdates()

            ' notify icon initialization
            nicon.Icon = New Icon(Application.GetResourceStream(New Uri("images/app.ico", UriKind.Relative)).Stream)
            nicon.Text = My.Application.Info.AssemblyName
            nicon.ContextMenu = ctxmenu
            nicon.Visible = True

            ' global parameters
            initializeGlobalParameters()

            ' domains setup
            initializeDomains()

            ' preferences setup
            initializePreferences()

            ' main window
            If Not startminimized Then
                Dim w As NavigationWindow
                w = ActivateMainWindow()

                If w IsNot Nothing Then
                    If startwithsearch Then
                        If TypeOf w Is NavigationWindow AndAlso TypeOf CType(w, NavigationWindow).Content Is pgMain Then
                            Dim pgm = CType(CType(w, NavigationWindow).Content, pgMain)
                            Dim pgo = pgm.CurrentObjectsPage
                            If pgo IsNot Nothing And Not String.IsNullOrEmpty(startsearchstring) Then pgo.StartSearch(Nothing, New clsFilter(startsearchstring, preferences.AttributesForSearch, preferences.SearchObjectClasses))
                        End If
                    End If
                Else
                    w = CreateMainWindow()
                    If startwithsearch Then
                        Dim neh As NavigatedEventHandler
                        neh = Sub()
                                  If TypeOf w Is NavigationWindow AndAlso TypeOf CType(w, NavigationWindow).Content Is pgMain Then
                                      Dim pgm = CType(CType(w, NavigationWindow).Content, pgMain)
                                      Dim pgo = pgm.CurrentObjectsPage
                                      If pgo IsNot Nothing And Not String.IsNullOrEmpty(startsearchstring) Then pgo.StartSearch(Nothing, New clsFilter(startsearchstring, preferences.AttributesForSearch, preferences.SearchObjectClasses))
                                  End If
                                  RemoveHandler w.Navigated, neh
                              End Sub
                        AddHandler w.Navigated, neh
                    End If
                End If
            End If

        Catch ex As Exception
            IMsgBox(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "Application.OnStartup")
        End Try
    End Sub

    Protected Overrides Sub OnExit(e As ExitEventArgs)
        MyBase.OnExit(e)
        Try
            ' notify icon deinitialization
            nicon.Visible = False

            ' save preferences
            deinitializePreferences()

            ' removing unhandled exception handler
            RemoveHandler Dispatcher.UnhandledException, AddressOf Dispatcher_UnhandledException
            RemoveHandler AppDomain.CurrentDomain.UnhandledException, AddressOf AppDomain_CurrentDomain_UnhandledException

        Catch ex As Exception

        End Try
    End Sub

    Sub New()
        InitializeComponent()
    End Sub

    Private Shared Sub ni_ctxmenuExit(sender As Object, e As EventArgs)
        ApplicationDeactivate()
    End Sub

    Private Shared Sub ni_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles nicon.MouseClick
        If e.Button = Forms.MouseButtons.Left Then
            If ActivateMainWindow() Is Nothing Then CreateMainWindow()
        End If
    End Sub

    Public Shared Function ActivateMainWindow() As NavigationWindow
        For Each w As Window In Current.Windows
            If TypeOf w Is NavigationWindow AndAlso TypeOf CType(w, NavigationWindow).Content Is pgMain Then
                If w.WindowState = WindowState.Minimized Then w.WindowState = WindowState.Maximized
                w.Show()
                w.Activate()
                w.Topmost = True : w.Topmost = False
                Return w
            End If
        Next
        ' no window found
        Return Nothing
    End Function

    Public Shared Function CreateMainWindow() As NavigationWindow
        Dim w As NavigationWindow = ShowPage(New pgMain)
        If w IsNot Nothing Then
            w.WindowState = WindowState.Maximized
            w.Icon = New BitmapImage(New Uri("pack://application:,,,/images/app.ico"))
            w.ShowInTaskbar = True
            w.Title = My.Application.Info.AssemblyName
            AddHandler w.Closing,
                Sub()
                    Dim count As Integer = 0

                    For Each wnd As Window In Current.Windows
                        If TypeOf wnd Is NavigationWindow AndAlso TypeOf CType(wnd, NavigationWindow).Content Is pgMain Then count += 1
                    Next

                    If preferences.CloseOnXButton AndAlso count <= 1 Then ApplicationDeactivate()
                End Sub
        End If
        Return w
    End Function

    Public Shared Sub Dispatcher_UnhandledException(ByVal sender As Object, ByVal e As DispatcherUnhandledExceptionEventArgs)
        ThrowException(e.Exception, My.Resources.str_UnhandledException)
        e.Handled = True
    End Sub

    Public Shared Sub AppDomain_CurrentDomain_UnhandledException(sender As Object, e As Object)
        Dim ex = TryCast(e, UnhandledExceptionEventArgs)
        If ex IsNot Nothing Then
            ThrowException(ex.Exception, My.Resources.str_UnhandledException)
        Else
            ThrowCustomException(My.Resources.str_UnhandledException)
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
