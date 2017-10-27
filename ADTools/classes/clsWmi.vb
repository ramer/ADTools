Imports System.Management
Imports System.Threading
Imports System.Threading.Tasks

Public Class clsWmi
    Private basicsearchtasks As New List(Of Task)
    Private basicsearchtaskscts As New CancellationTokenSource

    Sub New()

    End Sub

    Public Async Function BasicSearchWmiAsync(returncollection As clsThreadSafeObservableCollection(Of clsEvent),
                                         currentobject As clsDirectoryObject,
                                         dtfrom As Date, dtto As Date,
                                         Optional evttype As Integer = 0) As Task

        basicsearchtaskscts.Cancel()
        If basicsearchtasks.Count > 0 Then Exit Function
        returncollection.Clear()
        basicsearchtaskscts = New CancellationTokenSource

        Dim mt = Task.Factory.StartNew(
                Function()

                    Dim scope As ManagementScope
                    Dim query As String
                    Dim searcher As ManagementObjectSearcher = Nothing
                    Dim connection As New ConnectionOptions
                    connection.Username = currentobject.Domain.Username
                    connection.Password = currentobject.Domain.Password

                    Try
                        scope = New ManagementScope("\\" & currentobject.name & "." & currentobject.Domain.Name & "\root\CIMV2", connection)
                        scope.Connect()

                        query = "Select * from Win32_NTLogEvent Where Logfile = 'Security' AND "

                        Select Case evttype
                            Case 0 : query &= "(EventCode = '4624' OR EventCode = '4625' OR EventCode = '4634' OR EventCode = '4740' OR EventCode = '4771' OR EventCode = '4776')"
                            Case 1 : query &= "(EventCode = '4624' OR EventCode = '4625' OR EventCode = '4634')"
                            Case 2 : query &= "(EventCode = '4740' OR EventCode = '4771' OR EventCode = '4776')"
                            Case Else : query &= "(EventCode = '4624' OR EventCode = '4625' OR EventCode = '4634' OR EventCode = '4740' OR EventCode = '4771' OR EventCode = '4776')"
                        End Select

                        query &= " AND TimeGenerated >= '" & ManagementDateTimeConverter.ToDmtfDateTime(dtfrom) & "' and TimeGenerated < '" & ManagementDateTimeConverter.ToDmtfDateTime(dtto) & "'"

                        searcher = New ManagementObjectSearcher(scope, New ObjectQuery(query))

                        For Each obj As ManagementObject In searcher.Get()
                            If basicsearchtaskscts.Token.IsCancellationRequested Then Return False

                            If obj Is Nothing Then Continue For
                            Dim evt As New clsEvent(obj)

                            If evt.MessageAccountName IsNot Nothing AndAlso
                                (LCase(evt.MessageAccountName).Contains("system") Or
                                LCase(evt.MessageAccountName).Contains("система") Or
                                LCase(evt.MessageAccountName).Contains(LCase(currentobject.name))) Then Continue For

                            returncollection.Add(evt)
                            Threading.Thread.Sleep(30)
                        Next

                    Catch ex As Exception
                        ThrowException(ex, "Retrieving events error (WMI)")
                    End Try

                    Return True
                End Function, basicsearchtaskscts.Token)
        basicsearchtasks.Add(mt)

        Dim ft = mt.ContinueWith(
                Function(parenttask As Task(Of Boolean))
                    Try
                        basicsearchtasks.Remove(mt)
                    Catch
                        basicsearchtasks.Clear()
                    End Try

                    Return True
                End Function)

        Await Task.WhenAll(basicsearchtasks.ToArray)
    End Function

End Class
