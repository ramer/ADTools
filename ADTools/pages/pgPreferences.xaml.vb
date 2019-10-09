Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Data
Imports System.Reflection
Imports CredentialManagement
Imports IPrompt.VisualBasic
Imports Microsoft.Win32

Public Class pgPreferences

    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(pgPreferences),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)

            ShowObjectAttributes(value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgPreferences = CType(d, pgPreferences)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Public Property Attributes As ObservableCollection(Of clsAttribute)
    Private Property PluginProcessTelegramBot As Process
    Private Property PluginProcessSIP As Process

    Public Property MainWindowLayoutTable As DataTable

    Private Sub wndPreferences_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        'grid gridlines style (magic)
        Dim GLR = Activator.CreateInstance(Type.[GetType]("System.Windows.Controls.Grid+GridLinesRenderer," + " PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"))
        GLR.[GetType]().GetField("s_oddDashPen", BindingFlags.[Static] Or BindingFlags.NonPublic).SetValue(GLR, New Pen(New SolidColorBrush(preferences.ColorButtonBackground), 0.5))
        GLR.[GetType]().GetField("s_evenDashPen", BindingFlags.[Static] Or BindingFlags.NonPublic).SetValue(GLR, New Pen(Brushes.Transparent, 0.5))

        'main datacontext
        DataContext = preferences

        'mainwindow attributes
        Dim attrsdict As New Dictionary(Of String, clsAttribute)
        domains.Where(Function(domain) domain.Validated).SelectMany(Function(domain) domain.AttributesSchema.Values).ToList.ForEach(Sub(a As clsAttribute) If Not attrsdict.ContainsKey(a.lDAPDisplayName) Then attrsdict.Add(a.lDAPDisplayName, a))
        GetType(clsDirectoryObject).GetProperties.ToList.ForEach(Sub(prop As PropertyInfo) If prop.GetCustomAttribute(Of ExtendedPropertyAttribute) IsNot Nothing AndAlso Not attrsdict.ContainsKey(prop.Name) Then attrsdict.Add(prop.Name, New clsAttribute(prop.Name, True, 0, "2.5.5.12", prop.Name, 0, 1024, True)))
        Attributes = New ObservableCollection(Of clsAttribute)(attrsdict.Values)
        lvAttributes.ItemsSource = Attributes
        clsSorter.ApplySort(lvAttributes.Items, "IsDefault", ListSortDirection.Descending)
        lvAttributesForSearch.ItemsSource = preferences.AttributesForSearch 'domains.Where(Function(domain) domain.Validated).SelectMany(Function(domain) domain.AttributesSchema).Where(Function(attr) preferences.AttributesForSearch.Contains(attr.lDAPDisplayName))

        'mainwindow layout
        GenerateMainWindowLayoutTable()

        ' get TelegramBot settings
        chbPluginTelegramBotStartOnLogon.IsChecked = (My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).GetValue("ADToolsTelegramBot", "") = IO.Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Plugins\ADToolsTelegramBot\ADToolsTelegramBot.exe"))

        Dim cred As Credential
        cred = New Credential("", "", "ADToolsTelegramBot", CredentialType.Generic)
        cred.PersistanceType = PersistanceType.Enterprise
        cred.Load()
        tbPluginTelegramBotUsername.Text = cred.Username
        tbPluginTelegramBotAPIKey.Text = cred.Password
        Dim regADToolsTelegramBot As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADToolsTelegramBot")
        Try
            chbPluginTelegramBotUseProxy.IsChecked = Convert.ToBoolean(regADToolsTelegramBot.GetValue("UseProxy", False))
        Catch
        End Try
        tbPluginTelegramBotProxyAddress.Text = regADToolsTelegramBot.GetValue("ProxyAddress", "")
        tbPluginTelegramBotProxyPort.Text = regADToolsTelegramBot.GetValue("ProxyPort", "")

        ' get SIP settings
        chbPluginSIPStartOnLogon.IsChecked = (My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).GetValue("ADToolsSIP", "") = IO.Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Plugins\ADToolsSIP\ADToolsSIP.exe"))

        cred = New Credential("", "", "ADToolsSIP", CredentialType.Generic)
        cred.PersistanceType = PersistanceType.Enterprise
        cred.Load()
        tbPluginSIPUsername.Text = cred.Username
        tbPluginSIPPassword.Text = cred.Password
        Dim regADToolsSIP As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADToolsSIP")
        tbPluginSIPServer.Text = regADToolsSIP.GetValue("Server", "")
        cmboPluginSIPProtocol.Text = regADToolsSIP.GetValue("Protocol", "UDP")
        tbPluginSIPRegistrationName.Text = regADToolsSIP.GetValue("RegistrationName", "")
        tbPluginSIPDomain.Text = regADToolsSIP.GetValue("Domain", "")

        UpdatePluginsStatus()
    End Sub

#Region "MainWindow"

    Private Sub btnAttributesForSearchDefault_Click(sender As Object, e As RoutedEventArgs) Handles btnAttributesForSearchDefault.Click
        preferences.AttributesForSearch.Clear()
        attributesForSearchDefault.ToList.ForEach(Sub(attr) preferences.AttributesForSearch.Add(attr))
    End Sub

    Private Sub btnLayoutDefault_Click(sender As Object, e As RoutedEventArgs) Handles btnLayoutDefault.Click
        preferences.Columns = columnsDefault
        GenerateMainWindowLayoutTable()
    End Sub

    Private Async Sub ShowObjectAttributes(obj As clsDirectoryObject)
        capAttributes.Visibility = Visibility.Visible
        Await Task.Run(
            Sub()
                For Each attr In Attributes
                    attr = obj.GetValue(attr.lDAPDisplayName)
                Next
            End Sub)
        capAttributes.Visibility = Visibility.Hidden
    End Sub

    Private Sub lvAttributes_DragEnterDragOver(sender As Object, e As DragEventArgs) Handles lvAttributes.DragEnter, lvAttributes.DragOver
        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            e.Effects = DragDropEffects.Move
        Else
            e.Effects = DragDropEffects.None
        End If

        e.Handled = True
    End Sub

    Private Sub lvAttributes_Drop(sender As Object, e As DragEventArgs) Handles lvAttributes.Drop
        If e.Data.GetDataPresent(GetType(clsDirectoryObject())) Then
            Dim dropped As clsDirectoryObject() = e.Data.GetData(GetType(clsDirectoryObject()))
            If dropped.Count <> 1 Then Exit Sub
            CurrentObject = dropped(0)
        End If
    End Sub

    Private Sub lv_DragEnterDragOver(sender As Object, e As DragEventArgs) Handles lvAttributesForSearch.DragEnter,
                                                                            trashAttributesForSearch.DragEnter,
                                                                            lvAttributesForSearch.DragOver,
                                                                            trashAttributesForSearch.DragOver
        If Not e.Data.GetDataPresent(GetType(clsAttribute())) Then Exit Sub

        Dim dragsource = e.Data.GetData("DragSource")

        If dragsource Is lvAttributes AndAlso sender Is lvAttributesForSearch Then
            e.Effects = DragDropEffects.Copy
        ElseIf dragsource Is lvAttributesForSearch Then
            If sender Is lvAttributesForSearch Then
                e.Effects = DragDropEffects.None
                trashAttributesForSearch.Visibility = Visibility.Visible
            ElseIf sender Is trashAttributesForSearch Then
                e.Effects = DragDropEffects.Move
                trashAttributesForSearch.Visibility = Visibility.Visible : trashAttributesForSearch.Background = Application.Current.Resources("ColorButtonBackground")
            End If
        Else
            e.Effects = DragDropEffects.None
        End If

        e.Handled = True
    End Sub

    Private Sub lv_DragLeave(sender As Object, e As DragEventArgs) Handles lvAttributesForSearch.DragLeave,
                                                                           trashAttributesForSearch.DragLeave

        If sender Is lvAttributesForSearch Then trashAttributesForSearch.Visibility = Visibility.Collapsed
        If sender Is trashAttributesForSearch Then trashAttributesForSearch.Visibility = Visibility.Collapsed : trashAttributesForSearch.Background = Brushes.Transparent
    End Sub

    Dim _dglayoutmousedownpoint As Point
    Dim _dglayoutdragoperation As Boolean

    Private Sub dgLayout_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles dgLayout.PreviewMouseDown
        _dglayoutmousedownpoint = e.GetPosition(Me)
        _dglayoutdragoperation = False

        Dim r As HitTestResult = VisualTreeHelper.HitTest(Me, e.GetPosition(Me))
        If r IsNot Nothing Then

            Dim cell = FindVisualParent(Of DataGridCell)(r.VisualHit, Me)
            If cell Is Nothing Then Exit Sub
            If CType(cell.DataContext, DataRowView).Row(cell.Column.SortMemberPath) IsNot DBNull.Value Then
                _dglayoutdragoperation = True
            End If

        End If
    End Sub

    Private Sub dgLayout_MouseMove(sender As Object, e As MouseEventArgs) Handles dgLayout.MouseMove
        If _dglayoutdragoperation AndAlso Math.Abs(Point.Subtract(_dglayoutmousedownpoint, e.GetPosition(Me)).Length) > 5 Then
            _dglayoutdragoperation = False

            Dim r As HitTestResult = VisualTreeHelper.HitTest(Me, e.GetPosition(Me))
            If r IsNot Nothing Then

                Dim cell = FindVisualParent(Of DataGridCell)(r.VisualHit, Me)
                If cell Is Nothing Then Exit Sub
                If CType(cell.DataContext, DataRowView).Row(cell.Column.SortMemberPath) IsNot DBNull.Value Then
                    Dim attributename As String = CType(cell.DataContext, DataRowView).Row(cell.Column.SortMemberPath)
                    CType(cell.DataContext, DataRowView).Row(cell.Column.SortMemberPath) = DBNull.Value
                    DragDrop.DoDragDrop(sender, New DataObject({New clsAttribute("", False, 0, "", attributename, 0, 1024)}), DragDropEffects.All)
                    SaveMainWindowLayoutTable()
                    e.Handled = True
                End If

            End If
        End If
    End Sub

    Private Sub lv_dgLayout_Drop(sender As Object, e As DragEventArgs) Handles lvAttributesForSearch.Drop,
                                                                              trashAttributesForSearch.Drop,
                                                                              dgLayout.Drop

        trashAttributesForSearch.Visibility = Visibility.Collapsed : trashAttributesForSearch.Background = Brushes.Transparent

        If Not e.Data.GetDataPresent(GetType(clsAttribute())) Then Exit Sub

        Dim dragsource = e.Data.GetData("DragSource")

        Dim dropped As clsAttribute() = e.Data.GetData(GetType(clsAttribute()))
        If dropped.Count <> 1 Then Exit Sub

        Dim obj = dropped(0)

        If dragsource Is lvAttributes AndAlso sender Is lvAttributesForSearch Then ' adding attribute
            If obj.IsComplex Then IMsgBox(My.Resources.str_AttributesComplexCannotBeSearched, vbOKOnly + vbExclamation, My.Resources.str_AttributesForSearch) : Exit Sub
            If Not preferences.AttributesForSearch.Contains(obj) Then preferences.AttributesForSearch.Add(obj)
        ElseIf dragsource Is lvAttributesForSearch AndAlso sender Is trashAttributesForSearch Then ' removing attribute
            If preferences.AttributesForSearch.Contains(obj) Then
                If preferences.AttributesForSearch.Count <= 1 Then IMsgBox(My.Resources.str_AttributesAtLeastOne, vbOKOnly + vbExclamation, My.Resources.str_AttributesForSearch) : Exit Sub
                preferences.AttributesForSearch.Remove(obj)
            End If
        ElseIf sender Is dgLayout Then ' setting layout
            Dim cell = FindVisualParent(Of DataGridCell)(e.OriginalSource)
            If cell Is Nothing Then Exit Sub
            CType(cell.DataContext, DataRowView).Row(cell.Column.SortMemberPath) = obj.lDAPDisplayName
            SaveMainWindowLayoutTable()
        End If

    End Sub


    Private Sub dgLayout_ColumnReordered(sender As Object, e As DataGridColumnEventArgs) Handles dgLayout.ColumnReordered
        MainWindowLayoutTable.Columns(e.Column.Header).SetOrdinal(e.Column.DisplayIndex)
        SaveMainWindowLayoutTable()
    End Sub

    Private Sub dgLayout_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles dgLayout.PreviewKeyDown
        If e.Key = Key.Delete Then
            Dim cell = CType(e.OriginalSource, DataGridCell)
            If cell Is Nothing Then Exit Sub
            CType(cell.DataContext, DataRowView).Row(cell.Column.SortMemberPath) = DBNull.Value
            SaveMainWindowLayoutTable()
        End If
    End Sub

    Private Sub brdDataGridColumnHeader_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        If e.ClickCount = 2 Then
            CType(sender, Border).Visibility = Visibility.Collapsed
            Dim panel = FindVisualParent(Of Grid)(sender)
            If panel IsNot Nothing Then
                Dim tb = FindVisualChild(Of TextBox)(panel)
                If tb IsNot Nothing Then
                    tb.Focus()
                    tb.SelectionStart = tb.Text.Length
                    tb.SelectAll()
                End If
            End If
        End If
    End Sub

    Private Sub tbDataGridColumnHeader_LostFocus(sender As Object, e As RoutedEventArgs)
        Dim panel = FindVisualParent(Of Grid)(sender)
        If panel IsNot Nothing Then
            Dim cap = FindVisualChild(Of Border)(panel)
            If cap IsNot Nothing Then cap.Visibility = Visibility.Visible
        End If

        If CType(sender, TextBox).DataContext Is Nothing Then Exit Sub
        Dim column As DataColumn = MainWindowLayoutTable.Columns(CType(sender, TextBox).DataContext)
        If column IsNot Nothing Then
            column.ColumnName = CType(sender, TextBox).Text
            CType(sender, TextBox).DataContext = column.ColumnName
        End If

        SaveMainWindowLayoutTable()
        GenerateMainWindowLayoutTable()
    End Sub

    Private Sub mnuDataGridColumnHeaderAddColumn(sender As Object, e As RoutedEventArgs)
        Dim column As DataColumn = MainWindowLayoutTable.Columns(CType(sender, MenuItem).DataContext)
        If column IsNot Nothing Then
            Dim newcolumn = New DataColumn()
            MainWindowLayoutTable.Columns.Add(newcolumn)
            Dim columnnumber As Integer = 0
            For Each c In MainWindowLayoutTable.Columns
                If column Is c Then
                    Exit For
                End If
                columnnumber += 1
            Next
            newcolumn.SetOrdinal(columnnumber)
            SaveMainWindowLayoutTable()
            GenerateMainWindowLayoutTable()
        End If
    End Sub

    Private Sub mnuDataGridColumnHeaderRemoveColumn(sender As Object, e As RoutedEventArgs)
        If MainWindowLayoutTable.Columns.Count <= 1 Then Exit Sub
        Dim column As DataColumn = MainWindowLayoutTable.Columns(CType(sender, MenuItem).DataContext)
        If column IsNot Nothing Then
            MainWindowLayoutTable.Columns.Remove(column)
            SaveMainWindowLayoutTable()
            GenerateMainWindowLayoutTable()
        End If
    End Sub

    Private Sub GenerateMainWindowLayoutTable()
        MainWindowLayoutTable = New DataTable
        Dim columns = preferences.Columns.Count

        preferences.Columns.ToList.ForEach(
            Sub(col)
                Dim c As New DataColumn(col.Header)
                MainWindowLayoutTable.Columns.Add(c)
            End Sub)

        For r As Integer = 0 To 2
            Dim newRow = MainWindowLayoutTable.NewRow

            For c As Integer = 0 To columns - 1
                If r < preferences.Columns(c).Attributes.Count Then newRow(c) = preferences.Columns(c).Attributes(r)
            Next

            MainWindowLayoutTable.Rows.Add(newRow)
        Next

        dgLayout.ItemsSource = MainWindowLayoutTable.DefaultView
    End Sub

    Private Sub SaveMainWindowLayoutTable()
        Dim layout As New ObservableCollection(Of clsViewColumnInfo)

        For column As Integer = 0 To MainWindowLayoutTable.Columns.Count - 1
            Dim columninfo As New clsViewColumnInfo
            'columninfo.DisplayIndex = column
            columninfo.Header = MainWindowLayoutTable.Columns(column).ColumnName
            For row As Integer = 0 To MainWindowLayoutTable.Rows.Count - 1
                If MainWindowLayoutTable.Rows(row).Item(column) IsNot DBNull.Value Then
                    columninfo.Attributes.Add(MainWindowLayoutTable.Rows(row).Item(column))
                End If
            Next
            layout.Add(columninfo)
        Next
        preferences.Columns = layout
    End Sub

#End Region


    'Private Sub RebuildLayout()
    '    grdLayout.ColumnDefinitions.Clear()
    '    grdLayout.Children.Clear()

    '    For Each columnindo As clsViewColumnInfo In preferences.Columns
    '        grdLayout.ColumnDefinitions.Add(New ColumnDefinition())
    '    Next

    '    For x As Integer = 0 To preferences.Columns.Count - 1
    '        Dim tbHeader As New TextBox()
    '        tbHeader.SetValue(TextBox.TextProperty, preferences.Columns(x).Header)
    '        AddHandler tbHeader.TextChanged, AddressOf tbHeader_TextChanged

    '        Grid.SetRow(tbHeader, 0)
    '        Grid.SetColumn(tbHeader, x)
    '        grdLayout.Children.Add(tbHeader)

    '        For y As Integer = 0 To preferences.Columns(x).Attributes.Count - 1
    '            Dim attr As clsAttribute = preferences.Columns(x).Attributes(y)
    '            Dim tblck As New TextBlock(New Run(attr.Label))
    '            tblck.Tag = attr
    '            tblck.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap)
    '            tblck.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Center)
    '            tblck.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center)
    '            tblck.SetBinding(TextBlock.BackgroundProperty, New System.Windows.Data.Binding("{DynamicResource ColorButtonBackground}"))
    '            AddHandler tblck.MouseLeftButtonDown, AddressOf tblckAttribute_MouseLeftButtonDown

    '            Grid.SetRow(tblck, y + 1)
    '            Grid.SetColumn(tblck, x)
    '            grdLayout.Children.Add(tblck)
    '        Next
    '    Next

    'End Sub

    'Private Sub SaveLayout()
    '    Dim layout As New ObservableCollection(Of clsViewColumnInfo)

    '    For x As Integer = 0 To grdLayout.ColumnDefinitions.Count - 1
    '        Dim dgci As New clsViewColumnInfo
    '        dgci.DisplayIndex = x
    '        For y As Integer = 0 To grdLayout.RowDefinitions.Count - 1
    '            For Each element As UIElement In grdLayout.Children
    '                If Grid.GetColumn(element) = x And Grid.GetRow(element) = y Then
    '                    If y = 0 Then
    '                        dgci.Header = (CType(element, TextBox).GetValue(TextBox.TextProperty))
    '                    Else
    '                        Dim dummy As List(Of clsAttribute) = dgci.Attributes
    '                        dummy.Add(CType(element, TextBlock).Tag)
    '                        dgci.Attributes = dummy
    '                    End If
    '                End If
    '            Next
    '        Next
    '        layout.Add(dgci)
    '    Next
    '    preferences.Columns = layout
    'End Sub

    'Private Sub lv_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles lvAttributes.PreviewMouseLeftButtonDown, lvAttributesExtended.PreviewMouseLeftButtonDown, lvAttributesForSearch.PreviewMouseLeftButtonDown
    '    Dim listView As ListView = TryCast(sender, ListView)
    '    allowdrag = e.GetPosition(sender).X < listView.ActualWidth - SystemParameters.VerticalScrollBarWidth And e.GetPosition(sender).Y < listView.ActualHeight - SystemParameters.HorizontalScrollBarHeight
    'End Sub


    'Private Sub grdLayout_Drop(sender As Object, e As DragEventArgs) Handles grdLayout.Drop
    '    If Not e.Data.GetDataPresent(GetType(clsAttribute)) Then Exit Sub

    '    Dim point As Point = e.GetPosition(grdLayout)

    '    Dim row As Integer = 0
    '    Dim col As Integer = 0
    '    Dim accumulatedHeight As Double = 0.0
    '    Dim accumulatedWidth As Double = 0.0

    '    ' calc row mouse was over
    '    For Each rowDefinition As RowDefinition In grdLayout.RowDefinitions
    '        accumulatedHeight += rowDefinition.ActualHeight
    '        If accumulatedHeight >= point.Y Then
    '            Exit For
    '        End If
    '        row += 1
    '    Next

    '    ' calc col mouse was over
    '    For Each columnDefinition As ColumnDefinition In grdLayout.ColumnDefinitions
    '        accumulatedWidth += columnDefinition.ActualWidth
    '        If accumulatedWidth >= point.X Then
    '            Exit For
    '        End If
    '        col += 1
    '    Next

    '    If row = 0 Then Exit Sub

    '    ' create new
    '    Dim attr As clsAttribute = e.Data.GetData(GetType(clsAttribute))
    '    Dim tblck As New TextBlock(New Run(attr.Label))
    '    tblck.Tag = attr
    '    tblck.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap)
    '    tblck.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Center)
    '    tblck.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center)
    '    tblck.SetBinding(TextBlock.BackgroundProperty, New System.Windows.Data.Binding("{DynamicResource ColorButtonBackground}"))
    '    AddHandler tblck.MouseLeftButtonDown, AddressOf tblckAttribute_MouseLeftButtonDown

    '    ' delete old
    '    For Each element As UIElement In grdLayout.Children
    '        If Grid.GetRow(element) = row And Grid.GetColumn(element) = col Then
    '            grdLayout.Children.Remove(element)
    '            RemoveHandler element.MouseLeftButtonDown, AddressOf tblckAttribute_MouseLeftButtonDown
    '            Exit For
    '        End If
    '    Next

    '    ' add new
    '    Grid.SetRow(tblck, row)
    '    Grid.SetColumn(tblck, col)

    '    grdLayout.Children.Add(tblck)

    '    SaveLayout()
    'End Sub

    'Private Sub tblckAttribute_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
    '    Dim obj As TextBlock = CType(sender, TextBlock)
    '    grdLayout.Children.Remove(obj)
    '    RemoveHandler obj.MouseLeftButtonDown, AddressOf tblckAttribute_MouseLeftButtonDown

    '    Dim attr As clsAttribute = obj.Tag
    '    Dim dragData As New DataObject(attr)

    '    SaveLayout()

    '    DragDrop.DoDragDrop(grdLayout, dragData, DragDropEffects.Copy)
    'End Sub

    'Private Sub tbHeader_TextChanged(sender As Object, e As TextChangedEventArgs)
    '    SaveLayout()
    'End Sub

    'Private Sub btnLayoutAddColumn_Click(sender As Object, e As RoutedEventArgs) Handles btnLayoutAddColumn.Click
    '    grdLayout.ColumnDefinitions.Add(New ColumnDefinition)

    '    Dim tbheader As New TextBox()
    '    AddHandler tbheader.TextChanged, AddressOf tbHeader_TextChanged
    '    Grid.SetRow(tbheader, 0)
    '    Grid.SetColumn(tbheader, grdLayout.ColumnDefinitions.Count - 1)
    '    grdLayout.Children.Add(tbheader)

    '    SaveLayout()
    'End Sub

    'Private Sub btnLayoutRemoveLastColumn_Click(sender As Object, e As RoutedEventArgs) Handles btnLayoutRemoveLastColumn.Click
    '    Dim elementstoremove As New List(Of UIElement)
    '    For Each element As UIElement In grdLayout.Children
    '        If Grid.GetColumn(element) = grdLayout.ColumnDefinitions.Count - 1 Then
    '            elementstoremove.Add(element)
    '        End If
    '    Next
    '    For Each element As UIElement In elementstoremove
    '        grdLayout.Children.Remove(element)
    '        If element.GetType() Is GetType(TextBlock) Then RemoveHandler CType(element, TextBlock).MouseLeftButtonDown, AddressOf tblckAttribute_MouseLeftButtonDown
    '        If element.GetType() Is GetType(TextBox) Then RemoveHandler CType(element, TextBox).TextChanged, AddressOf tbHeader_TextChanged
    '    Next

    '    grdLayout.ColumnDefinitions.RemoveAt(grdLayout.ColumnDefinitions.Count - 1)

    '    SaveLayout()
    'End Sub



    Private Sub cmboTheme_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmboTheme.SelectionChanged
        Select Case cmboTheme.SelectedIndex
            Case 0 ' light gray
                With preferences
                    .ColorText = Colors.Black
                    .ColorWindowBackground = Colors.WhiteSmoke
                    .ColorElementBackground = Colors.White
                    .ColorMenuBackground = Colors.WhiteSmoke
                    .ColorButtonBackground = Colors.LightGray
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FFE8E6E6")
                    .ColorListviewRow = Colors.White
                    .ColorListviewAlternationRow = Colors.WhiteSmoke
                End With
            Case 1 ' light blue
                With preferences
                    .ColorText = Colors.Black
                    .ColorWindowBackground = Colors.WhiteSmoke
                    .ColorElementBackground = Colors.White
                    .ColorMenuBackground = Colors.WhiteSmoke
                    .ColorButtonBackground = Colors.LightSkyBlue
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FFD2EBFB")
                    .ColorListviewRow = Colors.White
                    .ColorListviewAlternationRow = Colors.AliceBlue
                End With
            Case 2 ' light green
                With preferences
                    .ColorText = Colors.Black
                    .ColorWindowBackground = Colors.WhiteSmoke
                    .ColorElementBackground = Colors.White
                    .ColorMenuBackground = Colors.WhiteSmoke
                    .ColorButtonBackground = Colors.LightGreen
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FFC8EEC8")
                    .ColorListviewRow = Colors.White
                    .ColorListviewAlternationRow = Colors.Honeydew
                End With
            Case 3 ' dark gray
                With preferences
                    .ColorText = Colors.WhiteSmoke
                    .ColorWindowBackground = Colors.Black
                    .ColorElementBackground = ColorConverter.ConvertFromString("#FF2E2E2E")
                    .ColorMenuBackground = Colors.Black
                    .ColorButtonBackground = Colors.DarkGray
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FF515151")
                    .ColorListviewRow = Colors.Black
                    .ColorListviewAlternationRow = ColorConverter.ConvertFromString("#FF2E2E2E")
                End With
            Case 4 ' dark blue
                With preferences
                    .ColorText = Colors.WhiteSmoke
                    .ColorWindowBackground = Colors.Black
                    .ColorElementBackground = ColorConverter.ConvertFromString("#FF2E2E2E")
                    .ColorMenuBackground = Colors.Black
                    .ColorButtonBackground = Colors.RoyalBlue
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FF1D2E63")
                    .ColorListviewRow = Colors.Black
                    .ColorListviewAlternationRow = ColorConverter.ConvertFromString("#FF2E2E2E")
                End With
            Case 5 ' dark green
                With preferences
                    .ColorText = Colors.WhiteSmoke
                    .ColorWindowBackground = Colors.Black
                    .ColorElementBackground = ColorConverter.ConvertFromString("#FF2E2E2E")
                    .ColorMenuBackground = Colors.Black
                    .ColorButtonBackground = Colors.Green
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FF003000")
                    .ColorListviewRow = Colors.Black
                    .ColorListviewAlternationRow = ColorConverter.ConvertFromString("#FF2E2E2E")
                End With
            Case 6 ' dark orange
                With preferences
                    .ColorText = Colors.WhiteSmoke
                    .ColorWindowBackground = Colors.Black
                    .ColorElementBackground = ColorConverter.ConvertFromString("#FF2E2E2E")
                    .ColorMenuBackground = Colors.Black
                    .ColorButtonBackground = Colors.Chocolate
                    .ColorButtonInactiveBackground = ColorConverter.ConvertFromString("#FF61310F")
                    .ColorListviewRow = Colors.Black
                    .ColorListviewAlternationRow = ColorConverter.ConvertFromString("#FF2E2E2E")
                End With
        End Select
    End Sub

    Private Sub chbStartOnLogon_Checked(sender As Object, e As RoutedEventArgs) Handles chbStartOnLogon.Checked, chbStartOnLogon.Unchecked, chbStartOnLogonMinimized.Checked, chbStartOnLogonMinimized.Unchecked
        If chbStartOnLogon.IsChecked Then
            If chbStartOnLogonMinimized.IsChecked Then
                My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).SetValue(My.Application.Info.AssemblyName, """" & System.Reflection.Assembly.GetExecutingAssembly().Location & """ -minimized") 'My.Application.Info.AssemblyName)
            Else
                My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).SetValue(My.Application.Info.AssemblyName, System.Reflection.Assembly.GetExecutingAssembly().Location) 'My.Application.Info.AssemblyName)
            End If
        Else
            My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(My.Application.Info.AssemblyName, False)
        End If
    End Sub

    Private Sub btnExternalSoftwareBrowse_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim es As clsExternalSoftware = CType(sender, FrameworkElement).DataContext
            Dim dlg As New Forms.OpenFileDialog
            dlg.Filter = My.Resources.str_ExternalSoftwareDialogFilter & "|*.exe"
            If dlg.ShowDialog = Forms.DialogResult.OK Then
                es.Path = dlg.FileName
            End If
        Catch
        End Try
    End Sub


#Region "Plugins"

    Private Sub UpdatePluginsStatus()
        PluginProcessTelegramBot = Nothing
        PluginProcessSIP = Nothing

        For Each p In Process.GetProcesses
            If LCase(p.ProcessName) = "adtoolstelegrambot" Then PluginProcessTelegramBot = p
            If LCase(p.ProcessName) = "adtoolssip" Then PluginProcessSIP = p
        Next

        If PluginProcessTelegramBot Is Nothing Then
            btnPluginTelegramBotStartStop.Content = My.Resources.str_PluginsTelegramBot_Start
        Else
            btnPluginTelegramBotStartStop.Content = My.Resources.str_PluginsTelegramBot_Stop
        End If

        If PluginProcessSIP Is Nothing Then
            btnPluginSIPStartStop.Content = My.Resources.str_PluginsSIP_Start
        Else
            btnPluginSIPStartStop.Content = My.Resources.str_PluginsSIP_Stop
        End If
    End Sub

    Private Sub tbPluginTelegramBot_LostFocus(sender As Object, e As RoutedEventArgs) Handles tbPluginTelegramBotUsername.LostFocus,
                                                                                              tbPluginTelegramBotAPIKey.LostFocus,
                                                                                              chbPluginTelegramBotUseProxy.Click,
                                                                                              tbPluginTelegramBotProxyPort.LostFocus,
                                                                                              tbPluginTelegramBotProxyAddress.LostFocus

        If String.IsNullOrEmpty(tbPluginTelegramBotUsername.Text) Or String.IsNullOrEmpty(tbPluginTelegramBotAPIKey.Text) Then Exit Sub

        Dim cred As New Credential("", "", "ADToolsTelegramBot", CredentialType.Generic)
        cred.PersistanceType = PersistanceType.Enterprise
        cred.Username = tbPluginTelegramBotUsername.Text
        cred.Password = tbPluginTelegramBotAPIKey.Text
        cred.Save()

        Dim regADToolsTelegramBot As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADToolsTelegramBot")
        regADToolsTelegramBot.SetValue("UseProxy", chbPluginTelegramBotUseProxy.IsChecked)
        regADToolsTelegramBot.SetValue("ProxyAddress", tbPluginTelegramBotProxyAddress.Text)
        regADToolsTelegramBot.SetValue("ProxyPort", tbPluginTelegramBotProxyPort.Text)

    End Sub

    Private Sub btnPluginTelegramBotStartStop_Click(sender As Object, e As RoutedEventArgs) Handles btnPluginTelegramBotStartStop.Click
        If PluginProcessTelegramBot Is Nothing Then
            Try
                Process.Start(IO.Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Plugins\ADToolsTelegramBot\ADToolsTelegramBot.exe"))
            Catch ex As Exception
                IMsgBox(ex.Message, vbOKOnly + vbExclamation,, Window.GetWindow(Me))
            End Try
        Else
            PluginProcessTelegramBot.Kill()
            PluginProcessTelegramBot.WaitForExit()
        End If

        UpdatePluginsStatus()
    End Sub

    Private Sub chbPluginTelegramBotStartOnLogon_Checked(sender As Object, e As RoutedEventArgs) Handles chbPluginTelegramBotStartOnLogon.Checked, chbPluginTelegramBotStartOnLogon.Unchecked
        If chbPluginTelegramBotStartOnLogon.IsChecked Then
            My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).SetValue("ADToolsTelegramBot", IO.Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Plugins\ADToolsTelegramBot\ADToolsTelegramBot.exe"))
        Else
            My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue("ADToolsTelegramBot", False)
        End If
    End Sub

    Private Sub tbPluginSIP_LostFocus(sender As Object, e As RoutedEventArgs) Handles tbPluginSIPUsername.LostFocus,
                                                                                      tbPluginSIPPassword.LostFocus,
                                                                                      tbPluginSIPServer.LostFocus,
                                                                                      cmboPluginSIPProtocol.LostFocus,
                                                                                      tbPluginSIPRegistrationName.LostFocus,
                                                                                      tbPluginSIPDomain.LostFocus

        If String.IsNullOrEmpty(tbPluginSIPUsername.Text) Or String.IsNullOrEmpty(tbPluginSIPPassword.Text) Then Exit Sub

        Dim cred As New Credential("", "", "ADToolsSIP", CredentialType.Generic)
        cred.PersistanceType = PersistanceType.Enterprise
        cred.Username = tbPluginSIPUsername.Text
        cred.Password = tbPluginSIPPassword.Text
        cred.Save()

        Dim regADToolsSIP As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\ADToolsSIP")
        regADToolsSIP.SetValue("Server", tbPluginSIPServer.Text)
        regADToolsSIP.SetValue("Protocol", cmboPluginSIPProtocol.Text)
        regADToolsSIP.SetValue("RegistrationName", tbPluginSIPRegistrationName.Text)
        regADToolsSIP.SetValue("Domain", tbPluginSIPDomain.Text)

    End Sub

    Private Sub btnPluginSIPStartStop_Click(sender As Object, e As RoutedEventArgs) Handles btnPluginSIPStartStop.Click
        If PluginProcessSIP Is Nothing Then
            Try
                Process.Start(IO.Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Plugins\ADToolsSIP\ADToolsSIP.exe"))
            Catch ex As Exception
                IMsgBox(ex.Message, vbOKOnly + vbExclamation,, Window.GetWindow(Me))
            End Try
        Else
            PluginProcessSIP.Kill()
            PluginProcessSIP.WaitForExit()
        End If

        UpdatePluginsStatus()
    End Sub

    Private Sub chbPluginSIPStartOnLogon_Checked(sender As Object, e As RoutedEventArgs) Handles chbPluginSIPStartOnLogon.Checked, chbPluginSIPStartOnLogon.Unchecked
        If chbPluginSIPStartOnLogon.IsChecked Then
            My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).SetValue("ADToolsSIP", IO.Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Plugins\ADToolsSIP\ADToolsSIP.exe"))
        Else
            My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue("ADToolsSIP", False)
        End If
    End Sub


#End Region

End Class
