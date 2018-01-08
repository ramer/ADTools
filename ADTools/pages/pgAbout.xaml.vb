Imports System.Reflection

Public Class pgAbout

    Public ReadOnly Property AppName As String
        Get
            Return My.Application.Info.AssemblyName
        End Get
    End Property

    Public ReadOnly Property Version As String
        Get
            Return Assembly.GetExecutingAssembly().GetName().Version.ToString()
        End Get
    End Property

    Public ReadOnly Property Copyright As String
        Get
            Return FileVersionInfo.GetVersionInfo(Me.GetType.Assembly.Location).LegalCopyright
        End Get
    End Property

    Public ReadOnly Property Company As String
        Get
            Return FileVersionInfo.GetVersionInfo(Me.GetType.Assembly.Location).CompanyName
        End Get
    End Property

    Private Sub imgDonate_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgDonate.MouseDown
        Donate()
    End Sub

End Class
