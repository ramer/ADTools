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
            Dim manager As New ADToolsSIPApplicationInstanceManager()
            manager.Run({""})
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbOKOnly + vbExclamation, "Sub Main()")
        End Try
    End Sub

End Module

Public Class ADToolsSIPApplicationInstanceManager
    Inherits WindowsFormsApplicationBase
    Private app As ADToolsSIPApplication

    Public Sub New()
        Me.IsSingleInstance = True
    End Sub

    Protected Overrides Function OnStartup(e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) As Boolean
        ' First time app is launched
        Try
            app = New ADToolsSIPApplication()
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

Public Class ADToolsSIPApplication
    Inherits Application

    Protected Overrides Sub OnStartup(e As Windows.StartupEventArgs)
        MyBase.OnStartup(e)
        Try
            ' unhandled exception handler
            AddHandler Dispatcher.UnhandledException, AddressOf Dispatcher_UnhandledException
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf AppDomain_CurrentDomain_UnhandledException

            ' setting localization
            FrameworkElement.LanguageProperty.OverrideMetadata(GetType(FrameworkElement), New FrameworkPropertyMetadata(XmlLanguage.GetLanguage(CultureInfo.CurrentCulture.IetfLanguageTag)))

            ' SIP 
            initializeSIP()
            'ShowPopup("Сороковиков", "3200")

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
