Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Windows.Markup
Imports System.Windows.Threading
Imports Microsoft.VisualBasic.ApplicationServices


Module mdlMain

    <STAThread>
    Sub Main()
        Try
            Dim manager As New ADToolsTelegramBotApplicationInstanceManager()
            manager.Run({""})
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "Sub Main()")
        End Try
    End Sub

End Module

Public Class ADToolsTelegramBotApplicationInstanceManager
    Inherits WindowsFormsApplicationBase
    Private app As ADToolsTelegramBotApplication

    Public Sub New()
        Me.IsSingleInstance = True
    End Sub

    Protected Overrides Function OnStartup(e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) As Boolean
        ' First time app is launched
        Try
            app = New ADToolsTelegramBotApplication()
            app.Run()
            Return False
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "OnStartup")
            Return False
        End Try
    End Function

    Protected Overrides Sub OnStartupNextInstance(eventArgs As StartupNextInstanceEventArgs)
        Try
            MyBase.OnStartupNextInstance(eventArgs)
            'app.Activate()
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "OnStartupNextInstance")
        End Try
    End Sub
End Class

Public Class ADToolsTelegramBotApplication
    Inherits Application

    Protected Overrides Sub OnStartup(e As Windows.StartupEventArgs)
        MyBase.OnStartup(e)
        Try
            ' unhandled exception handler
            AddHandler Dispatcher.UnhandledException, AddressOf Dispatcher_UnhandledException
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf AppDomain_CurrentDomain_UnhandledException

            ' command line args
            Dim commandlineargs() As String = Environment.GetCommandLineArgs()
            'Dim minimizedstart As Boolean = False
            'For Each cl In commandlineargs
            '    If cl = "-minimized" Or cl = "/minimized" Then
            '        minimizedstart = True
            '    End If
            'Next

            ' setting localization
            FrameworkElement.LanguageProperty.OverrideMetadata(GetType(FrameworkElement), New FrameworkPropertyMetadata(XmlLanguage.GetLanguage(CultureInfo.CurrentCulture.IetfLanguageTag)))

            ' get credentials from windows storage
            initializeCredentials

            ' domains setup
            initializeDomains()

            ' telegram update timer
            initializeTimer()


        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "Application.OnStartup")
        End Try
    End Sub

    Protected Overrides Sub OnExit(e As ExitEventArgs)
        MyBase.OnExit(e)
        Try

        Catch ex As Exception

        End Try
    End Sub

    Sub New()
        InitializeComponent()
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

End Class
