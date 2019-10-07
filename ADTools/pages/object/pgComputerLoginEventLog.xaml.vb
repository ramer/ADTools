Class pgComputerLoginEventLog

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                    GetType(clsDirectoryObject),
                                                    GetType(pgComputerLoginEventLog),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Public Property events As New clsThreadSafeObservableCollection(Of clsEvent)
    Public Property eventsTotal As New clsThreadSafeObservableCollection(Of clsEventTotal)
    Public WithEvents wmisearcher As New clsWmi

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgComputerLoginEventLog = CType(d, pgComputerLoginEventLog)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Sub New(obj As clsDirectoryObject)
        InitializeComponent()
        CurrentObject = obj
    End Sub

    Private Sub pgComputerLoginEventLog_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dgEvents.ItemsSource = events
        dgTotalEvents.ItemsSource = eventsTotal

        dtpPeriodTo.Value = Now
        dtpPeriodFrom.Value = Now.AddMonths(-1)
        dtpTotalPeriodTo.SelectedDate = Today
        dtpTotalPeriodFrom.SelectedDate = Today 'New Date(Today.Hour, Today.Month, 1)
    End Sub

    Private Sub dtpPeriodFrom_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles dtpPeriodFrom.ValueChanged
        If dtpPeriodTo.Value < dtpPeriodFrom.Value Then dtpPeriodTo.Value = dtpPeriodFrom.Value
    End Sub

    Private Sub dtpPeriodTo_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles dtpPeriodTo.ValueChanged
        If dtpPeriodTo.Value < dtpPeriodFrom.Value Then dtpPeriodFrom.Value = dtpPeriodTo.Value
    End Sub

    Private Sub dtpTotalPeriodFrom_ValueChanged(sender As Object, e As SelectionChangedEventArgs) Handles dtpTotalPeriodFrom.SelectedDateChanged
        If dtpTotalPeriodTo.SelectedDate < dtpTotalPeriodFrom.SelectedDate Then dtpTotalPeriodTo.SelectedDate = dtpTotalPeriodFrom.SelectedDate
    End Sub

    Private Sub dtpTotalPeriodTo_ValueChanged(sender As Object, e As SelectionChangedEventArgs) Handles dtpTotalPeriodTo.SelectedDateChanged
        If dtpTotalPeriodTo.SelectedDate < dtpTotalPeriodFrom.SelectedDate Then dtpTotalPeriodFrom.SelectedDate = dtpTotalPeriodTo.SelectedDate
    End Sub

    Private Async Sub btnEventsSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnEventsSearch.Click
        pbSearch.Visibility = Visibility.Visible

        Await wmisearcher.BasicSearchWmiAsync(events, CurrentObject, dtpPeriodFrom.Value, dtpPeriodTo.Value, If(rbEventAll.IsChecked, 0, If(rbEventSuccess.IsChecked, 1, 2)))

        pbSearch.Visibility = Visibility.Hidden
    End Sub

    Private Async Sub btnTotalEventsSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnTotalEventsSearch.Click
        If String.IsNullOrEmpty(tbTotalUsername.Text) Then Exit Sub

        pbTotalSearch.Visibility = Visibility.Visible

        Dim df = If(dtpTotalPeriodFrom.SelectedDate.HasValue, dtpTotalPeriodFrom.SelectedDate.Value, Today)
        Dim dt = If(dtpTotalPeriodTo.SelectedDate.HasValue, dtpTotalPeriodTo.SelectedDate.Value.AddDays(1), Today)

        Await wmisearcher.TotalSearchWmiAsync(eventsTotal, CurrentObject, df, dt, tbTotalUsername.Text)

        pbTotalSearch.Visibility = Visibility.Hidden
    End Sub

    Private Sub ctxmnuCopyEvents(sender As Object, e As RoutedEventArgs)
        If dgEvents.SelectedItems Is Nothing Then Exit Sub

        Dim evts As New List(Of clsEvent)
        For Each evt In dgEvents.SelectedItems
            evts.Add(evt)
        Next

        Try
            Clipboard.SetDataObject(Join(evts.Select(Function(evt As clsEvent) Join({evt.CategoryString, evt.EventCode, evt.MessageAccountName, evt.MessageLogonType, evt.MessageSourceAddress, evt.TimeGenerated}, vbTab)).ToArray, vbCrLf), True)
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation)
        End Try
    End Sub

    Private Sub ctxmnuCopyEventsTotal(sender As Object, e As RoutedEventArgs)
        If dgTotalEvents.SelectedItems Is Nothing Then Exit Sub

        Dim evts As New List(Of clsEventTotal)
        For Each evt In dgTotalEvents.SelectedItems
            evts.Add(evt)
        Next

        Try
            Clipboard.SetDataObject(Join(evts.Select(Function(evt As clsEventTotal) Join({evt.Day.ToShortDateString, evt.First.ToString, evt.Last.ToString, evt.Diff.ToString}, vbTab)).ToArray, vbCrLf), True)
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation)
        End Try
    End Sub

End Class
