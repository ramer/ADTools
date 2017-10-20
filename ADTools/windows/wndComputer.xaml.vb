Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Management
Imports System.Net.NetworkInformation
Imports System.Threading.Tasks
Imports System.Windows.Threading

Public Class wndComputer

    Public Property currentobject As clsDirectoryObject

    'Public Property events As New clsThreadSafeObservableCollection(Of clsEvent)
    'Public WithEvents wmisearcher As New clsWmi
    Dim pingtimer As New DispatcherTimer()

    Private Sub wndComputer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.DataContext = currentobject
        ctlMemberOf.CurrentObject = currentobject
        ctlAttributes.CurrentObject = currentobject

        'dgEvents.ItemsSource = events

        'dtpPeriodTo.Value = Now
        'dtpPeriodFrom.Value = Now.AddDays(-1)
        'AddHandler pingtimer.Tick, AddressOf pingtimer_Tick
        pingtimer.Interval = New TimeSpan(0, 0, 3)
    End Sub

    Private Sub wndComputer_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Me.Owner IsNot Nothing Then Me.Owner.Activate() 'magic - if we don't do that and wndUser(this) had children, wndMain becomes not focused and even under VisualStudio window, so we bring it back
    End Sub

    Private Sub tabctlComputer_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tabctlComputer.SelectionChanged
        If tabctlComputer.SelectedIndex = 1 Then
            ctlMemberOf.InitializeAsync()
        ElseIf tabctlComputer.SelectedIndex = 2 Then
            'pingtimer.Start()
            'GetPingStatus()
            'GetTrace()
            'GetPorts()
        ElseIf tabctlComputer.SelectedIndex = 4 Then
            ctlAttributes.InitializeAsync()
        End If

        If tabctlComputer.SelectedIndex <> 2 Then
            pingtimer.Stop()
        End If
    End Sub

    'Private Sub dtpPeriodFrom_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles dtpPeriodFrom.ValueChanged
    '    If dtpPeriodTo.Value < dtpPeriodFrom.Value Then dtpPeriodTo.Value = dtpPeriodFrom.Value
    'End Sub

    'Private Sub dtpPeriodTo_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles dtpPeriodTo.ValueChanged
    '    If dtpPeriodTo.Value < dtpPeriodFrom.Value Then dtpPeriodFrom.Value = dtpPeriodTo.Value
    'End Sub

    'Private Async Sub btnEventsSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnEventsSearch.Click
    '    pbSearch.Visibility = Visibility.Visible

    '    Await wmisearcher.BasicSearchWMIAsync(events, currentobject, dtpPeriodFrom.Value, dtpPeriodTo.Value, If(rbEventAll.IsChecked, 0, If(rbEventSuccess.IsChecked, 1, 2)))

    '    pbSearch.Visibility = Visibility.Hidden
    'End Sub

    'Private Sub dgEvents_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles dgEvents.MouseDoubleClick
    '    If dgEvents.SelectedItem Is Nothing Then Exit Sub
    '    IMsgBox(CType(dgEvents.SelectedItem, clsEvent).Message, CType(dgEvents.SelectedItem, clsEvent).CategoryString, vbOKOnly, vbInformation)
    'End Sub

    'Private Sub pingtimer_Tick(sender As Object, e As EventArgs)
    '    GetPingStatus()
    'End Sub

    'Private Async Sub GetPingStatus()
    '    Dim pingresult As PingReply = Nothing

    '    Await Task.Run(Sub() pingresult = Ping(currentobject.dNSHostName))

    '    Dim status As String
    '    Dim address As String
    '    Dim triptime As String

    '    If pingresult IsNot Nothing Then
    '        status = If(pingresult.Status = IPStatus.Success, "Доступен", "Не доступен")
    '        address = If(pingresult.Address IsNot Nothing, "Ответ от " & pingresult.Address.ToString & ": ", "")
    '        triptime = If(pingresult.Status = IPStatus.Success, pingresult.RoundtripTime & " мс", pingresult.Status.ToString)
    '    Else
    '        status = "Не доступен"
    '        address = "Адрес неизвестен"
    '        triptime = ""
    '    End If

    '    Dim sp As New StackPanel
    '    sp.Orientation = Orientation.Horizontal
    '    Dim img As New Image
    '    img.Source = New BitmapImage(If(pingresult IsNot Nothing, If(pingresult.Status = IPStatus.Success, New Uri("pack://application:,,,/img/ready.ico"), New Uri("pack://application:,,,/img/warning.ico")), New Uri("pack://application:,,,/img/warning.ico")))
    '    img.Width = 16.0
    '    img.Height = 16.0
    '    sp.Children.Add(img)
    '    Dim tb As New TextBlock(New Run(String.Format("{0}. {1}{2}", status, address, triptime)))
    '    tb.Margin = New Thickness(5, 0, 10, 5)
    '    tb.VerticalAlignment = VerticalAlignment.Center
    '    sp.Children.Add(tb)
    '    wpPing.Children.Clear()
    '    wpPing.Children.Add(sp)
    'End Sub

    'Private Async Sub GetTrace()
    '    wpTrace.Children.Clear()

    '    Dim tracelist As New List(Of PingReply)

    '    Await Task.Run(Sub() tracelist = TraceRoute(currentobject.dNSHostName))

    '    For Each trc In tracelist
    '        Dim sp As New StackPanel
    '        sp.Orientation = Orientation.Horizontal
    '        Dim img As New Image
    '        img.Source = New BitmapImage(New Uri("pack://application:,,,/img/ready.ico"))
    '        img.Width = 16.0
    '        img.Height = 16.0
    '        sp.Children.Add(img)
    '        Dim tb As New TextBlock(New Run(String.Format("{0} ({1} мс)", If(trc.Address IsNot Nothing, trc.Address.ToString, "Неизвестно"), trc.RoundtripTime)))
    '        tb.Margin = New Thickness(5, 0, 10, 5)
    '        tb.VerticalAlignment = VerticalAlignment.Center
    '        sp.Children.Add(tb)
    '        wpTrace.Children.Add(sp)
    '    Next

    'End Sub

    'Private Async Sub GetPorts()
    '    wpPorts.Children.Clear()

    '    Dim portlist As New Dictionary(Of Integer, Boolean)

    '    Await Task.Run(Sub() portlist = PortScan(currentobject.dNSHostName, portlistDefault))

    '    For Each prt As Integer In portlistDefault.Keys
    '        If Not portlist.ContainsKey(prt) Then Continue For
    '        Dim available As Boolean = portlist(prt)

    '        Dim sp As New StackPanel
    '        sp.Orientation = Orientation.Horizontal
    '        Dim img As New Image
    '        img.Source = New BitmapImage(If(available, New Uri("pack://application:,,,/img/ready.ico"), New Uri("pack://application:,,,/img/warning.ico")))
    '        img.Width = 16.0
    '        img.Height = 16.0
    '        sp.Children.Add(img)
    '        Dim tb As New TextBlock(New Run(String.Format("{0} ({1})", prt, portlistDefault(prt))))
    '        tb.Margin = New Thickness(5, 0, 10, 5)
    '        tb.VerticalAlignment = VerticalAlignment.Center
    '        sp.Children.Add(tb)
    '        wpPorts.Children.Add(sp)
    '    Next
    'End Sub

End Class
