Imports System.Windows.Controls.Primitives
Imports System.Windows.Threading

''' <summary>
''' Enables the selection inside of a ListBox using a seleciton rectangle.
''' </summary>
Public NotInheritable Class clsSelector
    ''' <summary>Identifies the IsEnabled attached property.</summary>
    Public Shared ReadOnly EnabledProperty As DependencyProperty = DependencyProperty.RegisterAttached("Enabled", GetType(Boolean), GetType(clsSelector), New UIPropertyMetadata(False, AddressOf IsEnabledChangedCallback))

    ' This stores the clsSelector for each ListBox so we can unregister it.
    Private Shared ReadOnly attachedControls As New Dictionary(Of ListBox, clsSelector)()

    Private ReadOnly _listBox As ListBox
    Private _scrollContent As ScrollContentPresenter

    Private _selectionRect As SelectionAdorner
    Private _autoScroller As AutoScroller
    Private _selector As ItemsControlSelector

    Private _mouseCaptured As Boolean
    Private _start As Point
    Private _end As Point
    Private _mousedownpoint As Point

    Private _dragoperation As Boolean

    Private Sub New(lb As ListBox)
        _listBox = lb
        If _listBox.IsLoaded Then
            Register()
        Else
            ' We need to wait for it to be loaded so we can find the
            ' child controls.
            AddHandler _listBox.Loaded, AddressOf OnListBoxLoaded
        End If
    End Sub

    ''' <summary>
    ''' Gets the value of the IsEnabled attached property that indicates
    ''' whether a selection rectangle can be used to select items or not.
    ''' </summary>
    ''' <param name="obj">Object on which to get the property.</param>
    ''' <returns>
    ''' true if items can be selected by a selection rectangle; otherwise, false.
    ''' </returns>
    Public Shared Function GetEnabled(obj As DependencyObject) As Boolean
        Return CBool(obj.GetValue(EnabledProperty))
    End Function

    ''' <summary>
    ''' Sets the value of the IsEnabled attached property that indicates
    ''' whether a selection rectangle can be used to select items or not.
    ''' </summary>
    ''' <param name="obj">Object on which to set the property.</param>
    ''' <param name="value">Value to set.</param>
    Public Shared Sub SetEnabled(obj As DependencyObject, value As Boolean)
        obj.SetValue(EnabledProperty, value)
    End Sub

    Private Shared Sub IsEnabledChangedCallback(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim lb As ListBox = TryCast(d, ListBox)
        If lb IsNot Nothing Then
            If CBool(e.NewValue) Then
                ' If we're enabling selection by a rectangle we can assume
                ' this means we want to be able to select more than one item.
                If lb.SelectionMode = SelectionMode.[Single] Then
                    lb.SelectionMode = SelectionMode.Extended
                End If

                attachedControls.Add(lb, New clsSelector(lb))
            Else
                ' Unregister the selector
                Dim selector As clsSelector = Nothing
                If attachedControls.TryGetValue(lb, selector) Then
                    attachedControls.Remove(lb)
                    selector.UnRegister()
                End If
            End If
        End If
    End Sub

    ' Finds the nearest child of the specified type, or null if one wasn't found.
    Private Shared Function FindChild(Of T As Class)(reference As DependencyObject) As T
        ' Do a breadth first search.
        Dim queue = New Queue(Of DependencyObject)()
        queue.Enqueue(reference)
        While queue.Count > 0
            Dim child As DependencyObject = queue.Dequeue()
            Dim obj As T = TryCast(child, T)
            If obj IsNot Nothing Then
                Return obj
            End If

            ' Add the children to the queue to search through later.
            For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(child) - 1
                queue.Enqueue(VisualTreeHelper.GetChild(child, i))
            Next
        End While
        Return Nothing
        ' Not found.
    End Function

    Private Function Register() As Boolean
        _scrollContent = FindChild(Of ScrollContentPresenter)(_listBox)
        If _scrollContent IsNot Nothing Then
            _autoScroller = New AutoScroller(_listBox)
            AddHandler _autoScroller.OffsetChanged, AddressOf OnOffsetChanged

            _selectionRect = New SelectionAdorner(_scrollContent)
            _scrollContent.AdornerLayer.Add(_selectionRect)

            _selector = New ItemsControlSelector(_listBox)

            ' The ListBox intercepts the regular MouseLeftButtonDown event
            ' to do its selection processing, so we need to handle the
            ' PreviewMouseLeftButtonDown. The scroll content won't receive
            ' the message if we click on a blank area so use the ListBox.
            AddHandler _listBox.PreviewMouseDown, AddressOf OnPreviewMouseDown
            AddHandler _listBox.MouseUp, AddressOf OnMouseUp
            AddHandler _listBox.MouseMove, AddressOf OnMouseMove
        End If

        ' Return success if we found the ScrollContentPresenter
        Return (_scrollContent IsNot Nothing)
    End Function

    Private Sub UnRegister()
        StopSelection()

        ' Remove all the event handlers so this instance can be reclaimed by the GC.
        RemoveHandler _listBox.PreviewMouseDown, AddressOf OnPreviewMouseDown
        RemoveHandler _listBox.MouseUp, AddressOf OnMouseUp
        RemoveHandler _listBox.MouseMove, AddressOf OnMouseMove

        _autoScroller.UnRegister()
    End Sub

    Private Sub OnListBoxLoaded(sender As Object, e As EventArgs)
        If Register() Then
            RemoveHandler _listBox.Loaded, AddressOf OnListBoxLoaded
        End If
    End Sub

    Private Sub OnOffsetChanged(sender As Object, e As OffsetChangedEventArgs)
        _selector.Scroll(e.HorizontalChange, e.VerticalChange)
        UpdateSelection()
    End Sub

    Private Sub OnMouseUp(sender As Object, e As MouseButtonEventArgs)
        If _mouseCaptured Then
            _mouseCaptured = False
            _scrollContent.ReleaseMouseCapture()
            StopSelection()
        End If
    End Sub

    Private Sub OnMouseMove(sender As Object, e As MouseEventArgs)
        If _dragoperation AndAlso Math.Abs(Point.Subtract(_mousedownpoint, e.GetPosition(sender)).Length) > 3 Then
            _mouseCaptured = False
            _scrollContent.ReleaseMouseCapture()
            StopSelection()

            _dragoperation = False

            DragDrop.DoDragDrop(e.Source, New DataObject(CType(e.Source, ListBox).SelectedItems), DragDropEffects.All)
        End If

        If _mouseCaptured Then
            ' Get the position relative to the content of the ScrollViewer.
            _end = e.GetPosition(_scrollContent)
            _autoScroller.Update(_end)
            UpdateSelection()
        End If
    End Sub

    Private Sub OnPreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
        If TypeOf e.OriginalSource Is Hyperlink OrElse
         (TypeOf e.OriginalSource Is Run AndAlso
         TypeOf CType(e.OriginalSource, Run).Parent Is Hyperlink) Then Exit Sub

        _mousedownpoint = e.GetPosition(sender)

        _dragoperation = False

        Dim r As HitTestResult = VisualTreeHelper.HitTest(sender, e.GetPosition(sender))
        If r IsNot Nothing Then
            Dim dp As DependencyObject = r.VisualHit
            While (dp IsNot Nothing) AndAlso Not (TypeOf dp Is ListBoxItem)
                dp = VisualTreeHelper.GetParent(dp)
            End While
            If dp IsNot Nothing Then
                If CType(dp, ListBoxItem).IsSelected Then
                    _dragoperation = True
                    e.Handled = True
                End If
            End If
        End If

        If Not _dragoperation Then
            Dim mouse As Point = e.GetPosition(_scrollContent)
            If (mouse.X >= 0) AndAlso (mouse.X < _scrollContent.ActualWidth) AndAlso (mouse.Y >= 0) AndAlso (mouse.Y < _scrollContent.ActualHeight) Then
                _mouseCaptured = TryCaptureMouse(e)
                If _mouseCaptured Then
                    StartSelection(mouse)
                End If
            End If
        End If
    End Sub

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
            If Mouse.Captured IsNot _listBox Then
                ' Something else wanted the mouse, let it keep it.
                Return False
            End If
        End If

        ' Either there's nothing under the mouse or the element doesn't want the mouse.
        Return _scrollContent.CaptureMouse()
    End Function

    Private Sub StopSelection()
        ' Hide the selection rectangle and stop the auto scrolling.
        _selectionRect.IsEnabled = False
        _autoScroller.IsEnabled = False
    End Sub

    Private Sub StartSelection(location As Point)
        ' We've stolen the MouseLeftButtonDown event from the ListBox
        ' so we need to manually give it focus.
        _listBox.Focus()

        _start = location
        _end = location

        ' Do we need to start a new selection?
        If ((Keyboard.Modifiers And ModifierKeys.Control) = 0) AndAlso ((Keyboard.Modifiers And ModifierKeys.Shift) = 0) Then
            ' Neither the shift key or control key is pressed, so
            ' clear the selection.
            _listBox.SelectedItems.Clear()
        End If

        _selector.Reset()
        UpdateSelection()

        _selectionRect.IsEnabled = True
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
        _selectionRect.SelectionArea = area

        ' Select the items.
        ' Transform the points to be relative to the ListBox.
        Dim topLeft As Point = _scrollContent.TranslatePoint(area.TopLeft, _listBox)
        Dim bottomRight As Point = _scrollContent.TranslatePoint(area.BottomRight, _listBox)

        ' And select the items.
        _selector.UpdateSelection(New Rect(topLeft, bottomRight))
    End Sub

    ''' <summary>
    ''' Automatically scrolls an ItemsControl when the mouse is dragged outside
    ''' of the control.
    ''' </summary>
    Private NotInheritable Class AutoScroller
        Private ReadOnly _autoScroll As New DispatcherTimer()
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
            _scrollViewer = FindChild(Of ScrollViewer)(ic)
            AddHandler _scrollViewer.ScrollChanged, AddressOf OnScrollChanged
            _scrollContent = FindChild(Of ScrollContentPresenter)(_scrollViewer)

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
                        Selector.SetIsSelected(item, True)
                    ElseIf itemBounds.IntersectsWith(_previousArea) Then
                        ' We previously changed the selection to true but it no
                        ' longer intersects with the area so clear the selection.
                        Selector.SetIsSelected(item, False)
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
End Class
