Imports System.Management
Imports System.Threading

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


    Public Async Function TotalSearchWmiAsync(returncollection As clsThreadSafeObservableCollection(Of clsEventTotal),
                                     currentobject As clsDirectoryObject,
                                     dtfrom As Date, dtto As Date,
                                     username As String) As Task

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
                    query &= "(EventCode = '4624' OR EventCode = '4625' OR EventCode = '4634' OR EventCode = '4740' OR EventCode = '4771' OR EventCode = '4776')"
                    query &= " AND TimeGenerated >= '" & ManagementDateTimeConverter.ToDmtfDateTime(dtfrom) & "' and TimeGenerated < '" & ManagementDateTimeConverter.ToDmtfDateTime(dtto) & "'"

                    Dim currentevts As New clsThreadSafeObservableCollection(Of clsEvent)
                    Dim currentdate As Date? = Nothing

                    searcher = New ManagementObjectSearcher(scope, New ObjectQuery(query))

                    For Each obj As ManagementObject In searcher.Get()
                        If basicsearchtaskscts.Token.IsCancellationRequested Then Return False

                        If obj Is Nothing Then Continue For
                        Dim evt As New clsEvent(obj)

                        If evt.MessageAccountName Is Nothing OrElse Not (LCase(evt.MessageAccountName) = LCase(username)) Then Continue For
                        If Not currentdate.HasValue Then currentdate = evt.TimeGenerated.Date

                        If Not (currentdate.Value.Date = evt.TimeGenerated.Date) Then
                            Dim tempevents As New clsThreadSafeObservableCollection(Of clsEvent)
                            For Each tmpevt In currentevts
                                tempevents.Add(tmpevt)
                            Next
                            returncollection.Add(New clsEventTotal(tempevents, currentdate))
                            currentdate = evt.TimeGenerated.Date
                            currentevts.Clear()
                            Threading.Thread.Sleep(30)
                        End If

                        currentevts.Add(evt)
                    Next

                    If currentevts.Count > 0 And currentdate.HasValue Then returncollection.Add(New clsEventTotal(currentevts, currentdate))

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
