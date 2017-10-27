Imports System.Net.NetworkInformation
Imports System.Windows.Threading

Public Class ctlNetwork


    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(ctlNetwork),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Sub New()
        InitializeComponent()
    End Sub

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ctlNetwork = CType(d, ctlNetwork)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Sub ctlAttributes_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

    End Sub

    Public Async Sub InitializeAsync()
        If _currentobject Is Nothing Then Exit Sub
        Dim hostname As String = CurrentObject.dNSHostName

        cap.Visibility = Visibility.Visible

        Dim pingresult As PingReply = Nothing
        Await Task.Run(Sub() pingresult = Ping(hostname))
        Dim pingstatus As String
        Dim pingaddress As String
        Dim pingtriptime As String
        If pingresult IsNot Nothing Then
            pingstatus = If(pingresult.Status = IPStatus.Success, My.Resources.ctlNetwork_msg_Available, My.Resources.ctlNetwork_msg_NotAvailable)
            pingaddress = If(pingresult.Address IsNot Nothing, My.Resources.ctlNetwork_msg_ReplyFrom & pingresult.Address.ToString & ": ", "")
            pingtriptime = If(pingresult.Status = IPStatus.Success, pingresult.RoundtripTime & " ms", pingresult.Status.ToString)
        Else
            pingstatus = My.Resources.ctlNetwork_msg_NotAvailable
            pingaddress = My.Resources.ctlNetwork_msg_AddressUnknown
            pingtriptime = ""
        End If
        Dim pingsp As New StackPanel
        pingsp.Orientation = Orientation.Horizontal
        Dim pingimg As New Image
        pingimg.Source = New BitmapImage(If(pingresult IsNot Nothing, If(pingresult.Status = IPStatus.Success, New Uri("pack://application:,,,/images/ready.ico"), New Uri("pack://application:,,,/images/warning.ico")), New Uri("pack://application:,,,/images/warning.ico")))
        pingimg.Width = 16.0
        pingimg.Height = 16.0
        pingsp.Children.Add(pingimg)
        Dim pingtb As New TextBlock(New Run(String.Format("{0}. {1}{2}", pingstatus, pingaddress, pingtriptime)))
        pingtb.Margin = New Thickness(5, 0, 10, 5)
        pingtb.VerticalAlignment = VerticalAlignment.Center
        pingsp.Children.Add(pingtb)
        wpPing.Children.Clear()
        wpPing.Children.Add(pingsp)


        wpTrace.Children.Clear()
        Dim traceresult As New List(Of PingReply)
        Await Task.Run(Sub() traceresult = TraceRoute(hostname))
        For Each trc In traceresult
            Dim tracesp As New StackPanel
            tracesp.Orientation = Orientation.Horizontal
            Dim traceimg As New Image
            traceimg.Source = New BitmapImage(New Uri("pack://application:,,,/images/ready.ico"))
            traceimg.Width = 16.0
            traceimg.Height = 16.0
            tracesp.Children.Add(traceimg)
            Dim tracetb As New TextBlock(New Run(String.Format("{0} ({1} ms)", If(trc.Address IsNot Nothing, trc.Address.ToString, My.Resources.ctlNetwork_msg_Unknown), trc.RoundtripTime)))
            tracetb.Margin = New Thickness(5, 0, 10, 5)
            tracetb.VerticalAlignment = VerticalAlignment.Center
            tracesp.Children.Add(tracetb)
            wpTrace.Children.Add(tracesp)
        Next


        wpPorts.Children.Clear()
        Dim portscanresult As New Dictionary(Of Integer, Boolean)
        Await Task.Run(Sub() portscanresult = PortScan(hostname, portlistDefault))

        For Each prt As Integer In portlistDefault.Keys
            If Not portscanresult.ContainsKey(prt) Then Continue For
            Dim available As Boolean = portscanresult(prt)

            Dim portscansp As New StackPanel
            portscansp.Orientation = Orientation.Horizontal
            Dim portscanimg As New Image
            portscanimg.Source = New BitmapImage(If(available, New Uri("pack://application:,,,/images/ready.ico"), New Uri("pack://application:,,,/images/warning.ico")))
            portscanimg.Width = 16.0
            portscanimg.Height = 16.0
            portscansp.Children.Add(portscanimg)
            Dim portscantb As New TextBlock(New Run(String.Format("{0} ({1})", prt, portlistDefault(prt))))
            portscantb.Margin = New Thickness(5, 0, 10, 5)
            portscantb.VerticalAlignment = VerticalAlignment.Center
            portscansp.Children.Add(portscantb)
            wpPorts.Children.Add(portscansp)
        Next

        cap.Visibility = Visibility.Hidden
    End Sub

End Class
