﻿Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces

'https://www.experts-exchange.com/questions/26329738/Programmatically-create-mailbox-in-Exchange-2010-Without-powershell.html

'if there is no connection to Exchange, the user's machine must be in PowerShell on behalf of the admin and type:
'winrm quickconfig
'Set-Item WSMan:\localhost\Client\TrustedHosts -Value SERVERNAME -Force

Public Class clsPowerShell
    Private _credential As PSCredential
    Private _connectionInfo As WSManConnectionInfo
    Private _rspace As Runspace
    Private _ps As PowerShell

    Sub New(login As String, password As String, server As String)
        Dim securepassword As New System.Security.SecureString

        For Each c As Char In password
            securepassword.AppendChar(c)
        Next

        Try
            _credential = New PSCredential(login, securepassword)
            _connectionInfo = New WSManConnectionInfo(New Uri("http://" & server & "/powershell?serializationLevel=Full"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", _credential)
            _connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos

            _ps = PowerShell.Create

            _rspace = RunspaceFactory.CreateRunspace(_connectionInfo)
            _rspace.Open()

            _ps.Runspace = _rspace
        Catch ex As Exception
        End Try
    End Sub

    Public ReadOnly Property State() As RunspaceStateInfo
        Get
            If _rspace IsNot Nothing Then
                Return _rspace.RunspaceStateInfo
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function Command(cmd As PSCommand) As Collection(Of PSObject)
        _ps.Commands = cmd

        Dim psHadException As Boolean
        Dim psException As Exception = Nothing
        Dim psResult As Collection(Of PSObject) = Nothing

        Dim res As String

        res = Join(cmd.Commands.Select(Function(x As Command) x.CommandText & vbCrLf & Join(x.Parameters.Select(Function(y As CommandParameter) "-" & y.Name & " " & y.Value.ToString).ToArray, vbCrLf)).ToArray, vbCrLf) & ". "

        For I As Integer = 0 To 4
            Try

                psHadException = False
                _ps.Streams.ClearStreams()
                psResult = _ps.Invoke()

            Catch ex As Exception
                psHadException = True
                psException = ex
            End Try

            If psHadException = False And _ps.HadErrors = False Then Exit For
        Next

        If psResult IsNot Nothing Then res &= "Objects count: " & psResult.Count & ". "

        If psHadException = True Then
            res &= "Error: " & If(psException IsNot Nothing, psException.Message & vbCrLf, "")
        End If

        If _ps.HadErrors = True Then
            For Each ex As ErrorRecord In _ps.Streams.Error
                res &= "Error: " & ex.Exception.Message & vbCrLf & "- CategoryInfo.Activity: " & ex.CategoryInfo.Activity & vbCrLf & "- CategoryInfo.Reason: " & ex.CategoryInfo.Reason
            Next
        End If

        Return psResult
    End Function

    Public Sub Close()
        Try

            _rspace.Close()

        Catch ex As Exception
            ThrowException(ex, "Exchange Dispose")
        Finally
            _rspace.Dispose()
            _rspace = Nothing
            _ps.Dispose()
            _ps = Nothing
        End Try
    End Sub

End Class
