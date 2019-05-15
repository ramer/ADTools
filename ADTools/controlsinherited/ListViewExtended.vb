
Public Class ListViewExtended
    Inherits ListView

    Public Event ItemDoubleClick(sender As Object, item As ListViewItemExtended)

    Private _scrollContent As ScrollContentPresenter
    Private _selectionAdorner As SelectionAdorner
    Private _autoScroller As AutoScroller
    Private _selector As ItemsControlSelector
    Private _dragadorner As DragAdorner

    Private _mousecaptured As Boolean
    Private _start As Point
    Private _end As Point
    Private _mousedownpoint As Point

    Private _dragoperation As Boolean

    Public Sub New()
        MyBase.New()

        If IsLoaded Then
            Initialize()
        Else
            AddHandler Loaded, AddressOf OnLoaded
        End If
    End Sub

    Private Sub OnLoaded(sender As Object, e As EventArgs)
        If Initialize() Then
            RemoveHandler Loaded, AddressOf OnLoaded
        End If
    End Sub

    Protected Overrides Function GetContainerForItemOverride() As DependencyObject
        Return New ListViewItemExtended()
    End Function

    Private Function Initialize() As Boolean
        _scrollContent = FindVisualChild(Of ScrollContentPresenter)(Me)
        If _scrollContent IsNot Nothing Then
            _autoScroller = New AutoScroller(Me)
            AddHandler _autoScroller.OffsetChanged, AddressOf OnAutoScrollerOffsetChanged

            _selectionAdorner = New SelectionAdorner(_scrollContent)
            _scrollContent.AdornerLayer.Add(_selectionAdorner)

            _selector = New ItemsControlSelector(Me)
        End If

        Return _scrollContent IsNot Nothing
    End Function

#Region "DependencyProperties"

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

    Public Property ViewStyleMediumIcons() As Style
        Get
            Return GetValue(ViewStyleMediumIconsProperty)
        End Get
        Set
            SetCurrentValue(ViewStyleMediumIconsProperty, Value)
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

#End Region

#Region "Events"

    Protected Overrides Sub OnMouseDown(e As MouseButtonEventArgs)
        MyBase.OnMouseDown(e)

        Dim r = VisualTreeHelper.HitTest(Me, e.GetPosition(Me))
        If TypeOf r.VisualHit Is ScrollViewer Then Me.UnselectAll()
    End Sub

    Protected Overrides Sub OnPreviewMouseDown(e As MouseButtonEventArgs)
        MyBase.OnPreviewMouseDown(e)

        ' Check that the mouse is inside the scroll content (could be on the
        ' scroll bars for example).
        _mousedownpoint = e.GetPosition(Me)

        _dragoperation = False

        Dim r As HitTestResult = VisualTreeHelper.HitTest(Me, e.GetPosition(Me))
        If r IsNot Nothing Then
            Dim cp = FindVisualParent(Of ContentPresenter)(r.VisualHit, Me)
            Dim lvi = FindVisualParent(Of ListViewItemExtended)(r.VisualHit, Me)
            If lvi IsNot Nothing Then
                If lvi.IsSelected Then
                    If e.ClickCount = 2 Then
                        RaiseEvent ItemDoubleClick(Me, lvi)
                    Else
                        _dragoperation = True
                        'e.Handled = True
                    End If
                Else
                    If cp IsNot Nothing Then
                        UnselectAll()
                        lvi.IsSelected = True
                        _dragoperation = True
                        'e.Handled = True
                    End If
                End If

            End If
        End If

        If Not _dragoperation And SelectionMode <> SelectionMode.Single Then
            Dim mousepoint As Point = e.GetPosition(_scrollContent)
            If (mousepoint.X >= 0) AndAlso (mousepoint.X < _scrollContent.ActualWidth) AndAlso (mousepoint.Y >= 0) AndAlso (mousepoint.Y < _scrollContent.ActualHeight) Then
                _mousecaptured = TryCaptureMouse(e)
                If _mousecaptured Then
                    StartSelection(mousepoint)
                End If
            End If
        End If

    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        If _dragoperation AndAlso Math.Abs(Point.Subtract(_mousedownpoint, e.GetPosition(Me)).Length) > 5 Then
            _mousecaptured = False
            _scrollContent.ReleaseMouseCapture()
            StopSelection()

            _dragoperation = False

            If SelectedItems.Count > 0 Then
                DoDragDrop(SelectedItems, e.GetPosition(Me))
            End If
        End If

        If _mousecaptured Then
            ' Get the position relative to the content of the ScrollViewer.
            _end = e.GetPosition(_scrollContent)
            _autoScroller.Update(_end)
            UpdateSelection()
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseButtonEventArgs)
        If _mousecaptured Then
            _mousecaptured = False
            _scrollContent.ReleaseMouseCapture()
            StopSelection()
        End If
    End Sub

    Private Sub OnAutoScrollerOffsetChanged(sender As Object, e As OffsetChangedEventArgs)
        _selector.Scroll(e.HorizontalChange, e.VerticalChange)
        UpdateSelection()
    End Sub

#End Region

#Region "Functions"

    Public Function FindVisualParent(Of T As DependencyObject)(ByVal child As Object, Optional until As DependencyObject = Nothing) As T
        If child Is Nothing Then Return Nothing
        Dim parent As DependencyObject = If(child.Parent IsNot Nothing, child.Parent, VisualTreeHelper.GetParent(child))

        If parent IsNot Nothing Then
            If TypeOf parent Is T Then
                Return parent
            ElseIf parent Is until Then
                Return Nothing
            Else
                Return FindVisualParent(Of T)(parent, until)
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function FindVisualChild(Of T As DependencyObject)(ByVal parent As Object) As T
        Dim queue = New Queue(Of DependencyObject)()
        queue.Enqueue(parent)
        While queue.Count > 0
            Dim child As DependencyObject = queue.Dequeue()
            If TypeOf child Is T Then
                Return child
            End If

            For I As Integer = 0 To VisualTreeHelper.GetChildrenCount(child) - 1
                queue.Enqueue(VisualTreeHelper.GetChild(child, I))
            Next
        End While
        Return Nothing
    End Function

    Private Function TryCaptureMouse(e As MouseButtonEventArgs) As Boolean
        Dim pos As Point = e.GetPosition(_scrollContent)

        ' Check if there is anything under the mouse.
        Dim element As UIElement = TryCast(_scrollContent.InputHitTest(pos), UIElement)
        If element IsNot Nothing Then
            ' Simulate a mouse click by sending it the MouseButtonDown
            ' event based on the data we received.
            Dim args = New MouseButtonEventArgs(e.MouseDevice, e.Timestamp, MouseButton.Left, e.StylusDevice)
            args.RoutedEvent = Mouse.MouseDownEvent
            args.Source = e.Source
            element.[RaiseEvent](args)

            ' The ListBox will try to capture the mouse unless something
            ' else captures it.
            If Mouse.Captured IsNot Me Then
                ' Something else wanted the mouse, let it keep it.
                Return False
            End If
        End If

        ' Either there's nothing under the mouse or the element doesn't want the mouse.
        Return _scrollContent.CaptureMouse()
    End Function

    Private Sub StopSelection()
        ' Hide the selection rectangle and stop the auto scrolling.
        _selectionAdorner.IsEnabled = False
        _autoScroller.IsEnabled = False
    End Sub

    Private Sub StartSelection(location As Point)
        ' We've stolen the MouseLeftButtonDown event from the ListBox
        ' so we need to manually give it focus.
        Focus()

        _start = location
        _end = location

        ' Do we need to start a new selection?
        If ((Keyboard.Modifiers And ModifierKeys.Control) = 0) AndAlso ((Keyboard.Modifiers And ModifierKeys.Shift) = 0) Then
            ' Neither the shift key or control key is pressed, so
            ' clear the selection.
            SelectedItems.Clear()
        End If

        _selector.Reset()
        UpdateSelection()

        _selectionAdorner.IsEnabled = True
        _autoScroller.IsEnabled = True
    End Sub

    Private Sub UpdateSelection()
        ' Offset the start point based on the scroll offset.
        Dim start As Point = _autoScroller.TranslatePoint(_start)

        ' Draw the selecion rectangle.
        ' Rect can't have a negative width/height...
        Dim x As Double = Math.Min(start.X, _end.X)
        Dim y As Double = Math.Min(start.Y, _end.Y)
        Dim width As Double = Math.Abs(_end.X - start.X)
        Dim height As Double = Math.Abs(_end.Y - start.Y)
        Dim area As New Rect(x, y, width, height)
        _selectionAdorner.SelectionArea = area

        ' Select the items.
        ' Transform the points to be relative to the ListBox.
        Dim topLeft As Point = _scrollContent.TranslatePoint(area.TopLeft, Me)
        Dim bottomRight As Point = _scrollContent.TranslatePoint(area.BottomRight, Me)

        ' And select the items.
        _selector.UpdateSelection(New Rect(topLeft, bottomRight))
    End Sub

    Private Sub DoDragDrop(Data As IList, Position As Point)
        Dim dragscope = CType(Window.GetWindow(Me).Content, FrameworkElement)
        Dim previousDrop As Boolean = dragscope.AllowDrop

        AddHandler dragscope.PreviewDragOver, AddressOf DragScope_PreviewDragOver

        _dragadorner = New DragAdorner(Me, CreateAdornerContent(Data), Position)
        AdornerLayer.GetAdornerLayer(dragscope).Add(_dragadorner)

        dragscope.AllowDrop = True

        Dim objectlist As Object = Activator.CreateInstance(GetType(List(Of)).MakeGenericType(SelectedItems(0).GetType))
        For Each element In SelectedItems
            objectlist.Add(element)
        Next

        Dim dataObj = New DataObject(objectlist.ToArray)
        dataObj.SetData("DragSource", Me)
        DragDrop.DoDragDrop(Me, dataObj, DragDropEffects.All)

        dragscope.AllowDrop = previousDrop

        AdornerLayer.GetAdornerLayer(dragscope).Remove(_dragadorner)
        _dragadorner = Nothing

        RemoveHandler dragscope.PreviewDragOver, AddressOf DragScope_PreviewDragOver
    End Sub

    Private Sub DragScope_PreviewDragOver(sender As Object, e As DragEventArgs)
        If _dragadorner IsNot Nothing Then
            _dragadorner.LeftOffset = e.GetPosition(Me).X
            _dragadorner.TopOffset = e.GetPosition(Me).Y
        End If
    End Sub

    Private Function CreateAdornerContent(objects As IList) As UIElement
        Dim brd As New Border With {
            .SnapsToDevicePixels = True,
            .CornerRadius = New CornerRadius(5, 5, 5, 5),
            .Background = Brushes.AliceBlue,
            .BorderThickness = New Thickness(1),
            .BorderBrush = Brushes.Black,
            .Opacity = 0.7,
            .Width = 100,
            .Height = 100}

        Dim grd As New Grid

        If objects.Count > 0 Then
            Dim imageprop As Reflection.PropertyInfo = Array.Find(objects(0).GetType.GetProperties, Function(prop) prop.Name = "Image")
            Dim imgwrapper As New Border With {.Margin = New Thickness(10)}
            Dim control = If(imageprop IsNot Nothing, TryCast(objects(0).Image, UIElement), Nothing)
            Dim bitmap = If(imageprop IsNot Nothing, TryCast(objects(0).Image, BitmapImage), Nothing)

            If control IsNot Nothing Then
                imgwrapper.Child = objects(0).Image
            ElseIf bitmap IsNot Nothing Then
                imgwrapper.Child = New Image With {.Source = objects(0).Image}
            Else
                Try
                    imgwrapper.Child = New Image With {.Source = New BitmapImage(New Uri("pack://application:,,,/images/property.png"))}
                Catch
                End Try
            End If

            grd.Children.Add(imgwrapper)
        End If

        If objects.Count > 1 Then
            grd.Children.Add(
             New Border With {
                 .SnapsToDevicePixels = True,
                 .VerticalAlignment = VerticalAlignment.Bottom,
                 .HorizontalAlignment = HorizontalAlignment.Center,
                 .Background = Brushes.AliceBlue,
                 .BorderThickness = New Thickness(1, 1, 1, 0),
                 .BorderBrush = Brushes.Black,
                 .Child = New TextBlock(New Run(objects.Count)) With {
                     .Margin = New Thickness(5, 0, 5, 0),
                     .FontSize = 16,
                     .HorizontalAlignment = HorizontalAlignment.Center,
                     .FontWeight = FontWeights.Medium}})
        End If

        brd.Child = grd
        brd.Measure(New Size(100, 100))
        brd.Arrange(New Rect(brd.DesiredSize))
        brd.UpdateLayout()
        Return brd
    End Function

#End Region

#Region "Subclasses"

    ''' <summary>
    ''' Automatically scrolls an ItemsControl when the mouse is dragged outside
    ''' of the control.
    ''' </summary>
    Private NotInheritable Class AutoScroller
        Private ReadOnly _autoScroll As New Windows.Threading.DispatcherTimer()
        Private ReadOnly _itemsControl As ItemsControl
        Private ReadOnly _scrollViewer As ScrollViewer
        Private ReadOnly _scrollContent As ScrollContentPresenter
        Private _isEnabled As Boolean
        Private _offset As Point
        Private _mouse As Point

        ''' <summary>
        ''' Initializes a new instance of the AutoScroller class.
        ''' </summary>
        ''' <param name="ic">The ItemsControl that is scrolled.</param>
        ''' <exception cref="ArgumentNullException">itemsControl is null.</exception>
        Public Sub New(ic As ItemsControl)
            If ic Is Nothing Then
                Throw New ArgumentNullException("itemsControl")
            End If

            _itemsControl = ic
            _scrollViewer = FindVisualChild(Of ScrollViewer)(ic)
            AddHandler _scrollViewer.ScrollChanged, AddressOf OnScrollChanged
            _scrollContent = FindVisualChild(Of ScrollContentPresenter)(_scrollViewer)

            AddHandler _autoScroll.Tick, AddressOf PreformScroll
            _autoScroll.Interval = TimeSpan.FromMilliseconds(GetRepeatRate())
        End Sub

        ''' <summary>Occurs when the scroll offset has changed.</summary>
        Public Event OffsetChanged As EventHandler(Of OffsetChangedEventArgs)

        ''' <summary>
        ''' Gets or sets a value indicating whether the auto-scroller is enabled
        ''' or not.
        ''' </summary>
        Public Property IsEnabled() As Boolean
            Get
                Return _isEnabled
            End Get
            Set
                If _isEnabled <> Value Then
                    _isEnabled = Value

                    ' Reset the auto-scroller and offset.
                    _autoScroll.IsEnabled = False
                    _offset = New Point()
                End If
            End Set
        End Property

        ''' <summary>
        ''' Translates the specified point by the current scroll offset.
        ''' </summary>
        ''' <param name="point">The point to translate.</param>
        ''' <returns>A new point offset by the current scroll amount.</returns>
        Public Function TranslatePoint(point As Point) As Point
            Return New Point(point.X - _offset.X, point.Y - _offset.Y)
        End Function

        ''' <summary>
        ''' Removes all the event handlers registered on the control.
        ''' </summary>
        Public Sub UnRegister()
            RemoveHandler _scrollViewer.ScrollChanged, AddressOf OnScrollChanged
        End Sub

        ''' <summary>
        ''' Updates the location of the mouse and automatically scrolls if required.
        ''' </summary>
        ''' <param name="mouse">
        ''' The location of the mouse, relative to the ScrollViewer's content.
        ''' </param>
        Public Sub Update(mouse As Point)
            _mouse = mouse

            ' If scrolling isn't enabled then see if it needs to be.
            If Not _autoScroll.IsEnabled Then
                PreformScroll()
            End If
        End Sub

        ' Returns the default repeat rate in milliseconds.
        Private Shared Function GetRepeatRate() As Integer
            ' The RepeatButton uses the SystemParameters.KeyboardSpeed as the
            ' default value for the Interval property. KeyboardSpeed returns
            ' a value between 0 (400ms) and 31 (33ms).
            Const Ratio As Double = (400.0 - 33.0) / 31.0
            Return 400 - CInt(SystemParameters.KeyboardSpeed * Ratio)
        End Function

        Private Function CalculateOffset(startIndex As Integer, endIndex As Integer) As Double
            Dim sum As Double = 0
            Dim i As Integer = startIndex
            While i <> endIndex
                Dim container As FrameworkElement = TryCast(_itemsControl.ItemContainerGenerator.ContainerFromIndex(i), FrameworkElement)
                If container IsNot Nothing Then
                    ' Height = Actual height + margin
                    sum += container.ActualHeight
                    sum += container.Margin.Top + container.Margin.Bottom
                End If
                i += 1
            End While
            Return sum
        End Function

        Private Sub OnScrollChanged(sender As Object, e As ScrollChangedEventArgs)
            ' Do we need to update the offset?
            If IsEnabled Then
                Dim horizontal As Double = e.HorizontalChange
                Dim vertical As Double = e.VerticalChange

                ' VerticalOffset means two seperate things based on the CanContentScroll
                ' property. If this property is true then the offset is the number of
                ' items to scroll; false then it's in Device Independant Pixels (DIPs).
                If _scrollViewer.CanContentScroll Then
                    ' We need to either increase the offset or decrease it.
                    If e.VerticalChange < 0 Then
                        Dim start As Integer = CInt(e.VerticalOffset)
                        Dim [end] As Integer = CInt(e.VerticalOffset - e.VerticalChange)
                        vertical = -CalculateOffset(start, [end])
                    Else
                        Dim start As Integer = CInt(e.VerticalOffset - e.VerticalChange)
                        Dim [end] As Integer = CInt(e.VerticalOffset)
                        vertical = CalculateOffset(start, [end])
                    End If
                End If

                _offset.X += horizontal
                _offset.Y += vertical

                RaiseEvent OffsetChanged(Me, New OffsetChangedEventArgs(horizontal, vertical))
            End If
        End Sub

        Private Sub PreformScroll()
            Dim scrolled As Boolean = False

            If _mouse.X > _scrollContent.ActualWidth Then
                _scrollViewer.LineRight()
                scrolled = True
            ElseIf _mouse.X < 0 Then
                _scrollViewer.LineLeft()
                scrolled = True
            End If

            If _mouse.Y > _scrollContent.ActualHeight Then
                _scrollViewer.LineDown()
                scrolled = True
            ElseIf _mouse.Y < 0 Then
                _scrollViewer.LineUp()
                scrolled = True
            End If

            ' It's important to disable scrolling if we're inside the bounds of
            ' the control so that when the user does leave the bounds we can
            ' re-enable scrolling and it will have the correct initial delay.
            _autoScroll.IsEnabled = scrolled
        End Sub
    End Class

    ''' <summary>Enables the selection of items by a specified rectangle.</summary>
    Private NotInheritable Class ItemsControlSelector
        Private ReadOnly _itemsControl As ItemsControl
        Private _previousArea As Rect

        ''' <summary>
        ''' Initializes a new instance of the ItemsControlSelector class.
        ''' </summary>
        ''' <param name="ic">
        ''' The control that contains the items to select.
        ''' </param>
        ''' <exception cref="ArgumentNullException">itemsControl is null.</exception>
        Public Sub New(ic As ItemsControl)
            If ic Is Nothing Then
                Throw New ArgumentNullException("itemsControl")
            End If
            _itemsControl = ic
        End Sub

        ''' <summary>
        ''' Resets the cached information, allowing a new selection to begin.
        ''' </summary>
        Public Sub Reset()
            _previousArea = New Rect()
        End Sub

        ''' <summary>
        ''' Scrolls the selection area by the specified amount.
        ''' </summary>
        ''' <param name="x">The horizontal scroll amount.</param>
        ''' <param name="y">The vertical scroll amount.</param>
        Public Sub Scroll(x As Double, y As Double)
            _previousArea.Offset(-x, -y)
        End Sub

        ''' <summary>
        ''' Updates the controls selection based on the specified area.
        ''' </summary>
        ''' <param name="area">
        ''' The selection area, relative to the control passed in the contructor.
        ''' </param>
        Public Sub UpdateSelection(area As Rect)
            ' Check eack item to see if it intersects with the area.
            For i As Integer = 0 To _itemsControl.Items.Count - 1
                Dim item As FrameworkElement = TryCast(_itemsControl.ItemContainerGenerator.ContainerFromIndex(i), FrameworkElement)
                If item IsNot Nothing Then
                    ' Get the bounds in the parent's co-ordinates.
                    Dim topLeft As Point = item.TranslatePoint(New Point(0, 0), _itemsControl)
                    Dim itemBounds As New Rect(topLeft.X, topLeft.Y, item.ActualWidth, item.ActualHeight)

                    ' Only change the selection if it intersects with the area
                    ' (or intersected i.e. we changed the value last time).
                    If itemBounds.IntersectsWith(area) Then
                        Primitives.Selector.SetIsSelected(item, True)
                    ElseIf itemBounds.IntersectsWith(_previousArea) Then
                        ' We previously changed the selection to true but it no
                        ' longer intersects with the area so clear the selection.
                        Primitives.Selector.SetIsSelected(item, False)
                    End If
                End If
            Next
            _previousArea = area
        End Sub
    End Class

    ''' <summary>The event data for the AutoScroller.OffsetChanged event.</summary>
    Private NotInheritable Class OffsetChangedEventArgs
        Inherits EventArgs
        Private ReadOnly _horizontal As Double
        Private ReadOnly _vertical As Double

        ''' <summary>
        ''' Initializes a new instance of the OffsetChangedEventArgs class.
        ''' </summary>
        ''' <param name="horizontal">The change in horizontal scroll.</param>
        ''' <param name="vertical">The change in vertical scroll.</param>
        Friend Sub New(horizontal As Double, vertical As Double)
            _horizontal = horizontal
            _vertical = vertical
        End Sub

        ''' <summary>Gets the change in horizontal scroll position.</summary>
        Public ReadOnly Property HorizontalChange() As Double
            Get
                Return _horizontal
            End Get
        End Property

        ''' <summary>Gets the change in vertical scroll position.</summary>
        Public ReadOnly Property VerticalChange() As Double
            Get
                Return _vertical
            End Get
        End Property
    End Class

    ''' <summary>Draws a selection rectangle on an AdornerLayer.</summary>
    Private NotInheritable Class SelectionAdorner
        Inherits Adorner
        Private selectionRect As Rect

        ''' <summary>
        ''' Initializes a new instance of the SelectionAdorner class.
        ''' </summary>
        ''' <param name="parent">
        ''' The UIElement which this instance will overlay.
        ''' </param>
        ''' <exception cref="ArgumentNullException">parent is null.</exception>
        Public Sub New(parent As UIElement)
            MyBase.New(parent)
            ' Make sure the mouse doesn't see us.
            IsHitTestVisible = False

            ' We only draw a rectangle when we're enabled.
            AddHandler IsEnabledChanged, AddressOf InvalidateVisual
        End Sub

        ''' <summary>Gets or sets the area of the selection rectangle.</summary>
        Public Property SelectionArea() As Rect
            Get
                Return selectionRect
            End Get
            Set
                selectionRect = Value
                InvalidateVisual()
            End Set
        End Property

        ''' <summary>
        ''' Participates in rendering operations that are directed by the layout system.
        ''' </summary>
        ''' <param name="drawingContext">The drawing instructions.</param>
        Protected Overrides Sub OnRender(drawingContext As DrawingContext)
            MyBase.OnRender(drawingContext)

            If IsEnabled Then
                ' Make the lines snap to pixels (add half the pen width [0.5])
                Dim x As Double() = {SelectionArea.Left + 0.5, SelectionArea.Right + 0.5}
                Dim y As Double() = {SelectionArea.Top + 0.5, SelectionArea.Bottom + 0.5}
                drawingContext.PushGuidelineSet(New GuidelineSet(x, y))

                Dim fill As Brush = SystemColors.HighlightBrush.Clone()
                fill.Opacity = 0.4
                drawingContext.DrawRectangle(fill, New Pen(SystemColors.HighlightBrush, 1.0), SelectionArea)
            End If
        End Sub
    End Class

    ''' <summary>Draws a rectangle above a mouse on an AdornerLayer.</summary>
    Private NotInheritable Class DragAdorner
        Inherits Adorner

        Private _child As Rectangle
        Private _leftoffset As Double
        Private _topoffset As Double
        Private _startoffset As Point

        Public Sub New(ByVal adornedElement As UIElement, ByVal content As UIElement, Optional startOffset As Point = Nothing)
            MyBase.New(adornedElement)

            _startoffset = startOffset
            _child = New Rectangle()
            _child.Width = content.RenderSize.Width
            _child.Height = content.RenderSize.Height
            _child.Fill = New VisualBrush(content)
        End Sub

        Protected Overrides Function MeasureOverride(ByVal constraint As System.Windows.Size) As System.Windows.Size
            _child.Measure(constraint)
            Return _child.DesiredSize
        End Function

        Protected Overrides Function ArrangeOverride(ByVal finalSize As System.Windows.Size) As System.Windows.Size
            _child.Arrange(New Rect(finalSize))
            Return finalSize
        End Function

        Protected Overrides Function GetVisualChild(ByVal index As Integer) As System.Windows.Media.Visual
            Return _child
        End Function

        Protected Overrides ReadOnly Property VisualChildrenCount() As Integer
            Get
                Return 1
            End Get
        End Property

        'With this bit Of code, we can already show the rectangle we wanted.
        'We'll want to allow the drag/drop code we wrote in the window to update the adorner to follow the mouse, so we'll add a couple of properties for this.

        Public Property LeftOffset() As Double
            Get
                Return _leftoffset
            End Get
            Set(ByVal value As Double)
                _leftoffset = value
                UpdatePosition()
            End Set
        End Property

        Public Property TopOffset() As Double
            Get
                Return _topoffset
            End Get
            Set(ByVal value As Double)
                _topoffset = value
                UpdatePosition()
            End Set
        End Property

        Private Sub UpdatePosition()
            Dim adornerLayer As AdornerLayer = Me.Parent
            If Not adornerLayer Is Nothing Then
                adornerLayer.Update(AdornedElement)
            End If
        End Sub

        'Finally, adorners are always placed relative To the element they adorn. You can Then place them relative To the corners, the middle, off To the side, within, etc. We'll just offset the adorner to where the user would like to drop the dragged element, by adding a translate transform to whatever was necessary to get to the adorned element.

        Public Overrides Function GetDesiredTransform(ByVal transform As System.Windows.Media.GeneralTransform) As System.Windows.Media.GeneralTransform
            Dim result As GeneralTransformGroup = New GeneralTransformGroup()
            result.Children.Add(MyBase.GetDesiredTransform(transform))
            result.Children.Add(New TranslateTransform(LeftOffset - _child.ActualWidth / 2, TopOffset - _child.ActualHeight))
            Return result
        End Function
    End Class



#End Region

End Class

''' <summary>ListViewItem with drag behaviour corrections.</summary>
Partial Public Class ListViewItemExtended
    Inherits ListViewItem

    Private _dragselected As Boolean = False

    Protected Overrides Sub OnDragEnter(e As DragEventArgs)
        MyBase.OnDragEnter(e)

        If Not IsSelected Then
            IsSelected = True
            _dragselected = True
        End If
    End Sub

    Protected Overrides Sub OnDragLeave(e As DragEventArgs)
        MyBase.OnDragLeave(e)

        If _dragselected Then
            IsSelected = False
            _dragselected = False
        End If
    End Sub

    Protected Overrides Sub OnDrop(e As DragEventArgs)
        MyBase.OnDrop(e)

        If _dragselected Then
            IsSelected = False
            _dragselected = False
        End If
    End Sub

    Private _selected As Boolean = False

    Protected Overrides Sub OnMouseLeftButtonDown(ByVal e As MouseButtonEventArgs)
        _selected = False

        If Not IsSelected Then
            MyBase.OnMouseLeftButtonDown(e)
            _selected = True
        End If
    End Sub

    Protected Overrides Sub OnMouseLeftButtonUp(ByVal e As MouseButtonEventArgs)
        If Not _selected Then
            MyBase.OnMouseLeftButtonDown(e)
        End If

        _selected = False
    End Sub

End Class