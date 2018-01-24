
Public Class pgDomainBrowser

    Public Shared ReadOnly RootObjectProperty As DependencyProperty = DependencyProperty.Register("RootObject",
                                                    GetType(clsDirectoryObject),
                                                    GetType(pgDomainBrowser),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf RootObjectPropertyChanged))

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                    GetType(clsDirectoryObject),
                                                    GetType(pgDomainBrowser),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _rootobject As clsDirectoryObject
    Private Property _currentobject As clsDirectoryObject
    Private valuereturned As Boolean

    Public Property RootObject() As clsDirectoryObject
        Get
            Return GetValue(RootObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(RootObjectProperty, value)
        End Set
    End Property

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub RootObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgDomainBrowser = CType(d, pgDomainBrowser)
        With instance
            ._rootobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgDomainBrowser = CType(d, pgDomainBrowser)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Sub New(ByRef root As clsDirectoryObject)
        InitializeComponent()
        RootObject = root
    End Sub

    Private Sub wndDomainBrowser_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If RootObject IsNot Nothing Then tvObjects.ItemsSource = {RootObject}
        AddHandler NavigationService.Navigating, AddressOf Navigating
    End Sub

    Private Sub Navigating(sender As Object, e As NavigatingCancelEventArgs)
        If e.NavigationMode = NavigationMode.Back And Not valuereturned Then
            e.Cancel = True
            valuereturned = True
            OnReturn(New ReturnEventArgs(Of clsDirectoryObject)(Nothing))
        End If
    End Sub

    Private Sub tviDomains_TreeViewItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim sp As StackPanel = CType(sender, StackPanel)

        If TypeOf sp.Tag Is clsDirectoryObject Then
            CurrentObject = (CType(sp.Tag, clsDirectoryObject))
            tbCurrentObject.Text = CurrentObject.distinguishedName
        End If
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
        If CurrentObject Is Nothing Then Exit Sub
        valuereturned = True
        OnReturn(New ReturnEventArgs(Of clsDirectoryObject)(CurrentObject))
    End Sub

End Class
