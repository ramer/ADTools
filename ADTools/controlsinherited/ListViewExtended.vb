
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Controls.Primitives

Public Class ListViewExtended
    Inherits ListView

    Public Enum Views
        Details = 0
        Tiles = 1
        List = 2
        MediumIcons = 3
    End Enum

    Public Property CurrentView() As Views
        Get
            Return GetValue(CurrentViewProperty)
        End Get
        Set
            SetCurrentValue(CurrentViewProperty, Value)
        End Set
    End Property

    Public Property ViewStyleDetails() As Style
        Get
            Return GetValue(ViewStyleDetailsProperty)
        End Get
        Set
            SetCurrentValue(ViewStyleDetailsProperty, Value)
        End Set
    End Property

    Public Property ViewStyleTiles() As Style
        Get
            Return GetValue(ViewStyleTilesProperty)
        End Get
        Set
            SetCurrentValue(ViewStyleTilesProperty, Value)
        End Set
    End Property

    Public Property ViewStyleList() As Style
        Get
            Return GetValue(ViewStyleListProperty)
        End Get
        Set
            SetCurrentValue(ViewStyleListProperty, Value)
        End Set
    End Property

    Public Property GroupItemStyle() As Style
        Get
            Return GetValue(GroupItemStyleProperty)
        End Get
        Set
            SetCurrentValue(GroupItemStyleProperty, Value)
        End Set
    End Property

    Public Property ViewStyleMediumIcons() As Style
        Get
            Return GetValue(ViewStyleMediumIconsProperty)
        End Get
        Set
            SetCurrentValue(ViewStyleMediumIconsProperty, Value)
        End Set
    End Property

    Public Property EnableGrouping() As Boolean
        Get
            Return GetValue(EnableGroupingProperty)
        End Get
        Set
            SetCurrentValue(EnableGroupingProperty, Value)
        End Set
    End Property

    Public Shared ReadOnly CurrentViewProperty As DependencyProperty = DependencyProperty.Register("CurrentView",
                                                    GetType(Views),
                                                    GetType(ListViewExtended),
                                                    New FrameworkPropertyMetadata(Views.Details, AddressOf CurrentViewPropertyChanged))

    Public Shared ReadOnly ViewStyleDetailsProperty As DependencyProperty = DependencyProperty.Register("ViewStyleDetails",
                                                    GetType(Style),
                                                    GetType(ListViewExtended),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf ViewStylePropertyChanged))

    Public Shared ReadOnly ViewStyleTilesProperty As DependencyProperty = DependencyProperty.Register("ViewStyleTiles",
                                                    GetType(Style),
                                                    GetType(ListViewExtended),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf ViewStylePropertyChanged))

    Public Shared ReadOnly ViewStyleListProperty As DependencyProperty = DependencyProperty.Register("ViewStyleList",
                                                    GetType(Style),
                                                    GetType(ListViewExtended),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf ViewStylePropertyChanged))

    Public Shared ReadOnly ViewStyleMediumIconsProperty As DependencyProperty = DependencyProperty.Register("ViewStyleMediumIcons",
                                                    GetType(Style),
                                                    GetType(ListViewExtended),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf ViewStylePropertyChanged))

    Public Shared ReadOnly GroupItemStyleProperty As DependencyProperty = DependencyProperty.Register("GroupItemStyle",
                                                    GetType(Style),
                                                    GetType(ListViewExtended),
                                                    New FrameworkPropertyMetadata(Nothing, AddressOf GroupItemStylePropertyChanged))

    Public Shared ReadOnly EnableGroupingProperty As DependencyProperty = DependencyProperty.Register("EnableGrouping",
                                                    GetType(Boolean),
                                                    GetType(ListViewExtended),
                                                    New FrameworkPropertyMetadata(False, AddressOf EnableGroupingPropertyChanged))

    Private Shared Sub CurrentViewPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ListViewExtended = CType(d, ListViewExtended)
        With instance
            .UpdateGroupStyle()

            Select Case e.NewValue
                Case Views.Details
                    .Style = .GetValue(ViewStyleDetailsProperty)
                Case Views.Tiles
                    .Style = .GetValue(ViewStyleTilesProperty)
                Case Views.List
                    .Style = .GetValue(ViewStyleListProperty)
                Case Views.MediumIcons
                    .Style = .GetValue(ViewStyleMediumIconsProperty)
                Case Else
                    .Style = .GetValue(ViewStyleDetailsProperty)
            End Select
        End With
    End Sub

    Private Shared Sub ViewStylePropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ListViewExtended = CType(d, ListViewExtended)
        With instance
            If e.Property Is ViewStyleDetailsProperty And .CurrentView = Views.Details Then .Style = e.NewValue
            If e.Property Is ViewStyleTilesProperty And .CurrentView = Views.Tiles Then .Style = e.NewValue
            If e.Property Is ViewStyleListProperty And .CurrentView = Views.List Then .Style = e.NewValue
            If e.Property Is ViewStyleMediumIconsProperty And .CurrentView = Views.MediumIcons Then .Style = e.NewValue
        End With
    End Sub

    Private Shared Sub GroupItemStylePropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ListViewExtended = CType(d, ListViewExtended)
        With instance
            .UpdateGroupStyle()
        End With
    End Sub

    Private Shared Sub EnableGroupingPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As ListViewExtended = CType(d, ListViewExtended)
        With instance
            .UpdateGroupStyle()
        End With
    End Sub

    Public Sub UpdateGroupStyle()
        If CurrentView = Views.Details And EnableGrouping Then
            Dim gs = New GroupStyle()
            gs.ContainerStyle = GroupItemStyle
            If gs.ContainerStyle Is Nothing Then Exit Sub
            GroupStyle.Clear()
            GroupStyle.Add(gs)
        Else
            GroupStyle.Clear()
        End If
    End Sub

#Region "Drag'n'Drop"

    Private _startpoint As Point
    Private _isdragging As Boolean
    Private _allOpsCursor As Cursor

    Private Sub DragSource_PreviewMouseLeftButtonDown(ByVal sender As Object, ByVal e As MouseButtonEventArgs) Handles Me.PreviewMouseLeftButtonDown
        _startpoint = e.GetPosition(Nothing)
    End Sub

    Private Sub DragSource_PreviewMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.PreviewMouseMove
        If e.LeftButton = MouseButtonState.Pressed AndAlso Not _isdragging Then
            Dim position As Point = e.GetPosition(Nothing)
            If Math.Abs(position.X - _startpoint.X) > SystemParameters.MinimumHorizontalDragDistance OrElse Math.Abs(position.Y - _startpoint.Y) > SystemParameters.MinimumVerticalDragDistance Then
                StartDragWindow(e)
            End If
        End If
    End Sub

    Private Sub StartDrag(ByVal e As MouseEventArgs)
        _isdragging = True
        Dim dragData As New DataObject(SelectedItems.Cast(Of clsDirectoryObject).ToArray)
        Dim de As DragDropEffects = DragDrop.DoDragDrop(Me, dragData, DragDropEffects.All)
        _isdragging = False
    End Sub

    Private Sub StartDragCustomCursor(ByVal e As MouseEventArgs)
        Dim handler As GiveFeedbackEventHandler = New GiveFeedbackEventHandler(AddressOf DragSource_GiveFeedback)
        AddHandler Me.GiveFeedback, handler
        _isdragging = True
        Dim dragData As New DataObject(SelectedItems.Cast(Of clsDirectoryObject).ToArray)
        Dim de As DragDropEffects = DragDrop.DoDragDrop(Me, dragData, DragDropEffects.All)
        RemoveHandler Me.GiveFeedback, handler
        _isdragging = False
    End Sub

    Private Sub StartDragWindow(ByVal e As MouseEventArgs)
        Dim feedbackhandler As GiveFeedbackEventHandler = New GiveFeedbackEventHandler(AddressOf DragSource_GiveFeedback)
        AddHandler Me.GiveFeedback, feedbackhandler
        Dim queryhandler As QueryContinueDragEventHandler = New QueryContinueDragEventHandler(AddressOf DragSource_QueryContinueDrag)
        AddHandler Me.QueryContinueDrag, queryhandler
        _isdragging = True

        CreateDragDropWindow(FindVisualParent(Of VirtualizingStackPanel)(e.OriginalSource))
        Dim dragData As New DataObject(SelectedItems.Cast(Of clsDirectoryObject).ToArray)
        _dragdropWindow.Show()
        Dim de As DragDropEffects = DragDrop.DoDragDrop(Me, dragData, DragDropEffects.All)
        DestroyDragDropWindow()
        _isdragging = False
        RemoveHandler Me.GiveFeedback, feedbackhandler
        RemoveHandler Me.QueryContinueDrag, queryhandler
    End Sub

    Private Sub DragSource_GiveFeedback(ByVal sender As Object, ByVal e As GiveFeedbackEventArgs)
        Try
            'If _allOpsCursor Is Nothing Then
            '    Using cursorStream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("SimplestDragDrop.DDIcon.cur")
            '        _allOpsCursor = New Cursor(cursorStream)
            '    End Using
            'End If

            'Mouse.SetCursor(_allOpsCursor)
            e.UseDefaultCursors = False
            UpdateWindowLocation()
            e.Handled = True
        Finally
        End Try
    End Sub

    Private Sub DragSource_QueryContinueDrag(ByVal sender As Object, ByVal e As QueryContinueDragEventArgs)
        Try

            If e.EscapePressed Then e.Action = DragAction.Cancel
            e.Handled = True
        Finally
        End Try
    End Sub

    Private _dragdropWindow As Window = Nothing

    <DllImport("user32.dll", EntryPoint:="GetWindowLong")>
    Private Shared Function GetWindowLongPtr32(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowLongPtr")>
    Private Shared Function GetWindowLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function

    ' This static method is required because Win32 does not support GetWindowLongPtr dirctly
    Public Shared Function GetWindowLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
        If IntPtr.Size = 8 Then
            Return GetWindowLongPtr64(hWnd, nIndex)
        Else
            Return GetWindowLongPtr32(hWnd, nIndex)
        End If
    End Function

    <DllImport("user32.dll", EntryPoint:="SetWindowLong")>
    Private Shared Function SetWindowLong32(ByVal hWnd As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLongFlags, ByVal dwNewLong As Integer) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="SetWindowLongPtr")>
    Private Shared Function SetWindowLongPtr64(ByVal hWnd As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLongFlags, ByVal dwNewLong As IntPtr) As IntPtr
    End Function

    Public Shared Function SetWindowLongPtr(ByVal hWnd As IntPtr, nIndex As WindowLongFlags, ByVal dwNewLong As IntPtr) As IntPtr
        If IntPtr.Size = 8 Then
            Return SetWindowLongPtr64(hWnd, nIndex, dwNewLong)
        Else
            Return New IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32))
        End If
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
    Public Shared Function GetCursorPos(ByRef lpPoint As POINTFX) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Public Enum WindowLongFlags
        GWL_WNDPROC = -4
        GWL_HINSTANCE = -6
        GWL_HWNDPARENT = -8
        GWL_STYLE = -16
        GWL_EXSTYLE = -20
        GWL_USERDATA = -21
        GWL_ID = -12
    End Enum

    Public Enum WindowStylesEx
        WS_EX_ACCEPTFILES = &H10
        WS_EX_APPWINDOW = &H40000
        WS_EX_CLIENTEDGE = &H200
        WS_EX_COMPOSITED = &H2000000
        WS_EX_CONTEXTHELP = &H400
        WS_EX_CONTROLPARENT = &H10000
        WS_EX_DLGMODALFRAME = &H1
        WS_EX_LAYERED = &H80000
        WS_EX_LAYOUTRTL = &H400000
        WS_EX_LEFT = &H0
        WS_EX_LEFTSCROLLBAR = &H4000
        WS_EX_LTRREADING = &H0
        WS_EX_MDICHILD = &H40
        WS_EX_NOACTIVATE = &H8000000
        WS_EX_NOINHERITLAYOUT = &H100000
        WS_EX_NOPARENTNOTIFY = &H4
        WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
        WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
        WS_EX_RIGHT = &H1000
        WS_EX_RIGHTSCROLLBAR = &H0
        WS_EX_RTLREADING = &H2000
        WS_EX_STATICEDGE = &H20000
        WS_EX_TOOLWINDOW = &H80
        WS_EX_TOPMOST = &H8
        WS_EX_TRANSPARENT = &H20
        WS_EX_WINDOWEDGE = &H100
    End Enum

    Structure POINTFX
        Public X As Integer
        Public Y As Integer
    End Structure

    Private Sub CreateDragDropWindow(ByVal dragElement As Visual)
        System.Diagnostics.Debug.Assert(Me._dragdropWindow Is Nothing)
        System.Diagnostics.Debug.Assert(dragElement IsNot Nothing)
        System.Diagnostics.Debug.Assert(TypeOf dragElement Is FrameworkElement)
        _dragdropWindow = New Window()
        _dragdropWindow.WindowStyle = WindowStyle.None
        _dragdropWindow.AllowsTransparency = True
        _dragdropWindow.AllowDrop = False
        _dragdropWindow.Background = Nothing
        _dragdropWindow.IsHitTestVisible = False
        _dragdropWindow.SizeToContent = SizeToContent.WidthAndHeight
        _dragdropWindow.Topmost = True
        _dragdropWindow.ShowInTaskbar = False
        AddHandler _dragdropWindow.SourceInitialized, New EventHandler(
            Sub(ByVal sender As Object, ByVal args As EventArgs)
                Dim windowSource As PresentationSource = PresentationSource.FromVisual(Me._dragdropWindow)
                Dim handle As IntPtr = (CType(windowSource, System.Windows.Interop.HwndSource)).Handle
                Dim styles As Int32 = GetWindowLongPtr(handle, WindowLongFlags.GWL_EXSTYLE)
                SetWindowLongPtr(handle, WindowLongFlags.GWL_EXSTYLE, styles Or WindowStylesEx.WS_EX_LAYERED Or WindowStylesEx.WS_EX_TRANSPARENT)
            End Sub)
        Dim r As Rectangle = New Rectangle()
        r.Width = (CType(dragElement, FrameworkElement)).ActualWidth
        r.Height = (CType(dragElement, FrameworkElement)).ActualHeight
        r.Fill = New VisualBrush(dragElement)
        _dragdropWindow.Content = r
        UpdateWindowLocation()
    End Sub

    Private Sub DestroyDragDropWindow()
        _dragdropWindow.Close()
        _dragdropWindow = Nothing
    End Sub

    Private Sub UpdateWindowLocation()
        If _dragdropWindow IsNot Nothing Then
            Dim p As POINTFX
            If Not GetCursorPos(p) Then
                Return
            End If

            _dragdropWindow.Left = p.X
            _dragdropWindow.Top = p.Y
        End If
    End Sub

#End Region




End Class

'Private _scrollContent As ScrollContentPresenter
'Private _selectionRect As SelectionAdorner
'Private _autoScroller As AutoScroller
'Private _selector As ItemsControlSelector

'Private _mouseCaptured As Boolean
'Private _start As Point
'Private _end As Point
'Private _mousedownpoint As Point
'Private _operationstarted As Boolean

'Private _dragoperation As Boolean
'    Private Sub ListViewExtended_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
'        _scrollContent = FindVisualChild(Of ScrollContentPresenter)(Me)
'        If _scrollContent IsNot Nothing Then
'            _autoScroller = New AutoScroller(Me)
'            AddHandler _autoScroller.OffsetChanged, AddressOf OnOffsetChanged

'            _selectionRect = New SelectionAdorner(_scrollContent)
'            _scrollContent.AdornerLayer.Add(_selectionRect)

'            _selector = New ItemsControlSelector(Me)
'        End If
'    End Sub

'    Private Shared Function FindVisualChild(Of T As DependencyObject)(ByVal parent As Object) As T
'        Dim queue = New Queue(Of DependencyObject)()
'        queue.Enqueue(parent)
'        While queue.Count > 0
'            Dim child As DependencyObject = queue.Dequeue()
'            If TypeOf child Is T Then
'                Return child
'            End If

'            For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(child) - 1
'                queue.Enqueue(VisualTreeHelper.GetChild(child, i))
'            Next
'        End While
'        Return Nothing
'    End Function

'    Private Sub OnOffsetChanged(sender As Object, e As OffsetChangedEventArgs)
'        _selector.Scroll(e.HorizontalChange, e.VerticalChange)
'        UpdateSelection()
'    End Sub

'    Private Sub ListViewExtended_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseUp
'        If _mouseCaptured Then
'            _mouseCaptured = False
'            _scrollContent.ReleaseMouseCapture()
'            StopSelection()
'        End If
'    End Sub

'    Private Sub ListViewExtended_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles Me.PreviewMouseDown
'        If (TypeOf e.OriginalSource Is Hyperlink OrElse
'             (TypeOf e.OriginalSource Is Run AndAlso
'             TypeOf CType(e.OriginalSource, Run).Parent Is Hyperlink)) Or
'             (e.ClickCount > 1) Then

'            Exit Sub
'        End If

'        If _mouseCaptured Then
'            _mouseCaptured = False
'            _scrollContent.ReleaseMouseCapture()
'            StopSelection()
'        End If
'        _mousedownpoint = e.GetPosition(_scrollContent)
'        _dragoperation = False


'        Dim r As HitTestResult = VisualTreeHelper.HitTest(_scrollContent, e.GetPosition(_scrollContent))
'        If r IsNot Nothing Then
'            Dim dp As DependencyObject = r.VisualHit
'            While (dp IsNot Nothing) AndAlso Not (TypeOf dp Is ListBoxItem)
'                dp = VisualTreeHelper.GetParent(dp)
'            End While
'            If dp IsNot Nothing Then
'                If CType(dp, ListBoxItem).IsSelected Then
'                    _dragoperation = True
'                    e.Handled = True
'                End If
'            End If
'        End If
'    End Sub

'    Private Sub ListViewExtended_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
'        If (TypeOf e.OriginalSource Is Hyperlink OrElse
'             (TypeOf e.OriginalSource Is Run AndAlso
'             TypeOf CType(e.OriginalSource, Run).Parent Is Hyperlink)) Or
'             (e.LeftButton = MouseButtonState.Released And e.RightButton = MouseButtonState.Released) Then
'            If _mouseCaptured Then
'                _mouseCaptured = False
'                _scrollContent.ReleaseMouseCapture()
'                StopSelection()
'            End If
'            Exit Sub
'        End If

'        Dim mouse As Point = e.GetPosition(_scrollContent)
'        If Math.Abs(Point.Subtract(_mousedownpoint, mouse).Length) > 3 AndAlso
'                (mouse.X >= 0) AndAlso (mouse.X < _scrollContent.ActualWidth) AndAlso (mouse.Y >= 0) AndAlso (mouse.Y < _scrollContent.ActualHeight) Then

'            If _dragoperation Then
'                Dim dragData As New DataObject(Me.SelectedItems.Cast(Of clsDirectoryObject).ToArray)
'                DragDrop.DoDragDrop(Me, dragData, DragDropEffects.All)
'            Else
'                If Not _mouseCaptured Then
'                    _mouseCaptured = _scrollContent.CaptureMouse()
'                    If _mouseCaptured Then StartSelection(mouse)
'                End If
'            End If
'        End If

'        If _mouseCaptured Then
'            _end = e.GetPosition(_scrollContent)
'            _autoScroller.Update(_end)
'            UpdateSelection()
'        End If
'    End Sub

'    Private Sub StopSelection()
'        ' Hide the selection rectangle and stop the auto scrolling.
'        _selectionRect.IsEnabled = False
'        _autoScroller.IsEnabled = False
'    End Sub

'    Private Sub StartSelection(location As Point)
'        ' We've stolen the MouseLeftButtonDown event from the ListBox
'        ' so we need to manually give it focus.
'        Me.Focus()

'        _start = location
'        _end = location

'        ' Do we need to start a new selection?
'        If ((Keyboard.Modifiers And ModifierKeys.Control) = 0) AndAlso ((Keyboard.Modifiers And ModifierKeys.Shift) = 0) Then
'            ' Neither the shift key or control key is pressed, so
'            ' clear the selection.
'            Me.SelectedItems.Clear()
'        End If

'        _selector.Reset()
'        UpdateSelection()

'        _selectionRect.IsEnabled = True
'        _autoScroller.IsEnabled = True
'    End Sub

'    Private Sub UpdateSelection()
'        ' Offset the start point based on the scroll offset.
'        Dim start As Point = _autoScroller.TranslatePoint(_start)

'        ' Draw the selecion rectangle.
'        ' Rect can't have a negative width/height...
'        Dim x As Double = Math.Min(start.X, _end.X)
'        Dim y As Double = Math.Min(start.Y, _end.Y)
'        Dim width As Double = Math.Abs(_end.X - start.X)
'        Dim height As Double = Math.Abs(_end.Y - start.Y)
'        Dim area As New Rect(x, y, width, height)
'        _selectionRect.SelectionArea = area

'        ' Select the items.
'        ' Transform the points to be relative to the ListBox.
'        Dim topLeft As Point = _scrollContent.TranslatePoint(area.TopLeft, Me)
'        Dim bottomRight As Point = _scrollContent.TranslatePoint(area.BottomRight, Me)

'        ' And select the items.
'        _selector.UpdateSelection(New Rect(topLeft, bottomRight))
'    End Sub


'    ''' <summary>
'    ''' Automatically scrolls an ItemsControl when the mouse is dragged outside
'    ''' of the control.
'    ''' </summary>
'    Private NotInheritable Class AutoScroller
'        Private ReadOnly _autoScroll As New Windows.Threading.DispatcherTimer()
'        Private ReadOnly _itemsControl As ItemsControl
'        Private ReadOnly _scrollViewer As ScrollViewer
'        Private ReadOnly _scrollContent As ScrollContentPresenter
'        Private _isEnabled As Boolean
'        Private _offset As Point
'        Private _mouse As Point

'        ''' <summary>
'        ''' Initializes a new instance of the AutoScroller class.
'        ''' </summary>
'        ''' <param name="ic">The ItemsControl that is scrolled.</param>
'        ''' <exception cref="ArgumentNullException">itemsControl is null.</exception>
'        Public Sub New(ic As ItemsControl)
'            If ic Is Nothing Then
'                Throw New ArgumentNullException("itemsControl")
'            End If

'            _itemsControl = ic
'            _scrollViewer = FindVisualChild(Of ScrollViewer)(ic)
'            AddHandler _scrollViewer.ScrollChanged, AddressOf OnScrollChanged
'            _scrollContent = FindVisualChild(Of ScrollContentPresenter)(_scrollViewer)

'            AddHandler _autoScroll.Tick, AddressOf PreformScroll
'            _autoScroll.Interval = TimeSpan.FromMilliseconds(GetRepeatRate())
'        End Sub

'        ''' <summary>Occurs when the scroll offset has changed.</summary>
'        Public Event OffsetChanged As EventHandler(Of OffsetChangedEventArgs)

'        ''' <summary>
'        ''' Gets or sets a value indicating whether the auto-scroller is enabled
'        ''' or not.
'        ''' </summary>
'        Public Property IsEnabled() As Boolean
'            Get
'                Return _isEnabled
'            End Get
'            Set
'                If _isEnabled <> Value Then
'                    _isEnabled = Value

'                    ' Reset the auto-scroller and offset.
'                    _autoScroll.IsEnabled = False
'                    _offset = New Point()
'                End If
'            End Set
'        End Property

'        ''' <summary>
'        ''' Translates the specified point by the current scroll offset.
'        ''' </summary>
'        ''' <param name="point">The point to translate.</param>
'        ''' <returns>A new point offset by the current scroll amount.</returns>
'        Public Function TranslatePoint(point As Point) As Point
'            Return New Point(point.X - _offset.X, point.Y - _offset.Y)
'        End Function

'        ''' <summary>
'        ''' Removes all the event handlers registered on the control.
'        ''' </summary>
'        Public Sub UnRegister()
'            RemoveHandler _scrollViewer.ScrollChanged, AddressOf OnScrollChanged
'        End Sub

'        ''' <summary>
'        ''' Updates the location of the mouse and automatically scrolls if required.
'        ''' </summary>
'        ''' <param name="mouse">
'        ''' The location of the mouse, relative to the ScrollViewer's content.
'        ''' </param>
'        Public Sub Update(mouse As Point)
'            _mouse = mouse

'            ' If scrolling isn't enabled then see if it needs to be.
'            If Not _autoScroll.IsEnabled Then
'                PreformScroll()
'            End If
'        End Sub

'        ' Returns the default repeat rate in milliseconds.
'        Private Shared Function GetRepeatRate() As Integer
'            ' The RepeatButton uses the SystemParameters.KeyboardSpeed as the
'            ' default value for the Interval property. KeyboardSpeed returns
'            ' a value between 0 (400ms) and 31 (33ms).
'            Const Ratio As Double = (400.0 - 33.0) / 31.0
'            Return 400 - CInt(SystemParameters.KeyboardSpeed * Ratio)
'        End Function

'        Private Function CalculateOffset(startIndex As Integer, endIndex As Integer) As Double
'            Dim sum As Double = 0
'            Dim i As Integer = startIndex
'            While i <> endIndex
'                Dim container As FrameworkElement = TryCast(_itemsControl.ItemContainerGenerator.ContainerFromIndex(i), FrameworkElement)
'                If container IsNot Nothing Then
'                    ' Height = Actual height + margin
'                    sum += container.ActualHeight
'                    sum += container.Margin.Top + container.Margin.Bottom
'                End If
'                i += 1
'            End While
'            Return sum
'        End Function

'        Private Sub OnScrollChanged(sender As Object, e As ScrollChangedEventArgs)
'            ' Do we need to update the offset?
'            If IsEnabled Then
'                Dim horizontal As Double = e.HorizontalChange
'                Dim vertical As Double = e.VerticalChange

'                ' VerticalOffset means two seperate things based on the CanContentScroll
'                ' property. If this property is true then the offset is the number of
'                ' items to scroll; false then it's in Device Independant Pixels (DIPs).
'                If _scrollViewer.CanContentScroll Then
'                    ' We need to either increase the offset or decrease it.
'                    If e.VerticalChange < 0 Then
'                        Dim start As Integer = CInt(e.VerticalOffset)
'                        Dim [end] As Integer = CInt(e.VerticalOffset - e.VerticalChange)
'                        vertical = -CalculateOffset(start, [end])
'                    Else
'                        Dim start As Integer = CInt(e.VerticalOffset - e.VerticalChange)
'                        Dim [end] As Integer = CInt(e.VerticalOffset)
'                        vertical = CalculateOffset(start, [end])
'                    End If
'                End If

'                _offset.X += horizontal
'                _offset.Y += vertical

'                RaiseEvent OffsetChanged(Me, New OffsetChangedEventArgs(horizontal, vertical))
'            End If
'        End Sub

'        Private Sub PreformScroll()
'            Dim scrolled As Boolean = False

'            If _mouse.X > _scrollContent.ActualWidth Then
'                _scrollViewer.LineRight()
'                scrolled = True
'            ElseIf _mouse.X < 0 Then
'                _scrollViewer.LineLeft()
'                scrolled = True
'            End If

'            If _mouse.Y > _scrollContent.ActualHeight Then
'                _scrollViewer.LineDown()
'                scrolled = True
'            ElseIf _mouse.Y < 0 Then
'                _scrollViewer.LineUp()
'                scrolled = True
'            End If

'            ' It's important to disable scrolling if we're inside the bounds of
'            ' the control so that when the user does leave the bounds we can
'            ' re-enable scrolling and it will have the correct initial delay.
'            _autoScroll.IsEnabled = scrolled
'        End Sub
'    End Class

'    ''' <summary>Enables the selection of items by a specified rectangle.</summary>
'    Private NotInheritable Class ItemsControlSelector
'        Private ReadOnly _itemsControl As ItemsControl
'        Private _previousArea As Rect

'        ''' <summary>
'        ''' Initializes a new instance of the ItemsControlSelector class.
'        ''' </summary>
'        ''' <param name="ic">
'        ''' The control that contains the items to select.
'        ''' </param>
'        ''' <exception cref="ArgumentNullException">itemsControl is null.</exception>
'        Public Sub New(ic As ItemsControl)
'            If ic Is Nothing Then
'                Throw New ArgumentNullException("itemsControl")
'            End If
'            _itemsControl = ic
'        End Sub

'        ''' <summary>
'        ''' Resets the cached information, allowing a new selection to begin.
'        ''' </summary>
'        Public Sub Reset()
'            _previousArea = New Rect()
'        End Sub

'        ''' <summary>
'        ''' Scrolls the selection area by the specified amount.
'        ''' </summary>
'        ''' <param name="x">The horizontal scroll amount.</param>
'        ''' <param name="y">The vertical scroll amount.</param>
'        Public Sub Scroll(x As Double, y As Double)
'            _previousArea.Offset(-x, -y)
'        End Sub

'        ''' <summary>
'        ''' Updates the controls selection based on the specified area.
'        ''' </summary>
'        ''' <param name="area">
'        ''' The selection area, relative to the control passed in the contructor.
'        ''' </param>
'        Public Sub UpdateSelection(area As Rect)
'            ' Check eack item to see if it intersects with the area.
'            For i As Integer = 0 To _itemsControl.Items.Count - 1
'                Dim item As FrameworkElement = TryCast(_itemsControl.ItemContainerGenerator.ContainerFromIndex(i), FrameworkElement)
'                If item IsNot Nothing Then
'                    ' Get the bounds in the parent's co-ordinates.
'                    Dim topLeft As Point = item.TranslatePoint(New Point(0, 0), _itemsControl)
'                    Dim itemBounds As New Rect(topLeft.X, topLeft.Y, item.ActualWidth, item.ActualHeight)

'                    ' Only change the selection if it intersects with the area
'                    ' (or intersected i.e. we changed the value last time).
'                    If itemBounds.IntersectsWith(area) Then
'                        Selector.SetIsSelected(item, True)
'                    ElseIf itemBounds.IntersectsWith(_previousArea) Then
'                        ' We previously changed the selection to true but it no
'                        ' longer intersects with the area so clear the selection.
'                        Selector.SetIsSelected(item, False)
'                    End If
'                End If
'            Next
'            _previousArea = area
'        End Sub
'    End Class

'    ''' <summary>The event data for the AutoScroller.OffsetChanged event.</summary>
'    Private NotInheritable Class OffsetChangedEventArgs
'        Inherits EventArgs
'        Private ReadOnly _horizontal As Double
'        Private ReadOnly _vertical As Double

'        ''' <summary>
'        ''' Initializes a new instance of the OffsetChangedEventArgs class.
'        ''' </summary>
'        ''' <param name="horizontal">The change in horizontal scroll.</param>
'        ''' <param name="vertical">The change in vertical scroll.</param>
'        Friend Sub New(horizontal As Double, vertical As Double)
'            _horizontal = horizontal
'            _vertical = vertical
'        End Sub

'        ''' <summary>Gets the change in horizontal scroll position.</summary>
'        Public ReadOnly Property HorizontalChange() As Double
'            Get
'                Return _horizontal
'            End Get
'        End Property

'        ''' <summary>Gets the change in vertical scroll position.</summary>
'        Public ReadOnly Property VerticalChange() As Double
'            Get
'                Return _vertical
'            End Get
'        End Property
'    End Class

'    ''' <summary>Draws a selection rectangle on an AdornerLayer.</summary>
'    Private NotInheritable Class SelectionAdorner
'        Inherits Adorner
'        Private selectionRect As Rect

'        ''' <summary>
'        ''' Initializes a new instance of the SelectionAdorner class.
'        ''' </summary>
'        ''' <param name="parent">
'        ''' The UIElement which this instance will overlay.
'        ''' </param>
'        ''' <exception cref="ArgumentNullException">parent is null.</exception>
'        Public Sub New(parent As UIElement)
'            MyBase.New(parent)
'            ' Make sure the mouse doesn't see us.
'            IsHitTestVisible = False

'            ' We only draw a rectangle when we're enabled.
'            AddHandler IsEnabledChanged, AddressOf InvalidateVisual
'        End Sub

'        ''' <summary>Gets or sets the area of the selection rectangle.</summary>
'        Public Property SelectionArea() As Rect
'            Get
'                Return selectionRect
'            End Get
'            Set
'                selectionRect = Value
'                InvalidateVisual()
'            End Set
'        End Property

'        ''' <summary>
'        ''' Participates in rendering operations that are directed by the layout system.
'        ''' </summary>
'        ''' <param name="drawingContext">The drawing instructions.</param>
'        Protected Overrides Sub OnRender(drawingContext As DrawingContext)
'            MyBase.OnRender(drawingContext)

'            If IsEnabled Then
'                ' Make the lines snap to pixels (add half the pen width [0.5])
'                Dim x As Double() = {SelectionArea.Left + 0.5, SelectionArea.Right + 0.5}
'                Dim y As Double() = {SelectionArea.Top + 0.5, SelectionArea.Bottom + 0.5}
'                drawingContext.PushGuidelineSet(New GuidelineSet(x, y))

'                Dim fill As Brush = SystemColors.HighlightBrush.Clone()
'                fill.Opacity = 0.4
'                drawingContext.DrawRectangle(fill, New Pen(SystemColors.HighlightBrush, 1.0), SelectionArea)
'            End If
'        End Sub
'    End Class
'End Class
