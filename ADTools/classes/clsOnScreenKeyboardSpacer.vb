
Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Public Class clsOnScreenKeyboardSpacer

#Region "Attached properties"

    Public Shared Function GetIsEnabled(ByVal obj As DependencyObject) As Boolean
        Dim objRtn As Object = obj.GetValue(IsEnabledProperty)
        If TypeOf objRtn Is Boolean Then Return objRtn
        Return Nothing
    End Function

    Public Shared Sub SetIsEnabled(ByVal obj As DependencyObject, ByVal value As Boolean)
        obj.SetValue(IsEnabledProperty, value)
    End Sub

    Public Shared ReadOnly IsEnabledProperty As DependencyProperty = DependencyProperty.RegisterAttached(
        "IsEnabled",
        GetType(Boolean),
        GetType(clsOnScreenKeyboardSpacer),
        New UIPropertyMetadata(False,
            Sub(o, e)
                Dim key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
                If key Is Nothing OrElse key.GetValue("ProductName") Is Nothing OrElse Not key.GetValue("ProductName").ToString.Contains("10") Then Exit Sub
                Dim wnd = Window.GetWindow(o)
                If e.NewValue = False Then
                    If wnd IsNot Nothing Then
                        RemoveHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf display_SettingsChanged
                        RemoveHandler wnd.SizeChanged, AddressOf wnd_SizeChanged
                        RemoveHandler wnd.LocationChanged, AddressOf wnd_LocationChanged
                    End If
                    If Objects.Contains(o) Then Objects.Remove(o)
                    OnScreenKeyboardWindowTimer.IsEnabled = Objects.Count > 0
                End If
                If e.NewValue IsNot Nothing Then
                    If wnd IsNot Nothing Then
                        AddHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf display_SettingsChanged
                        AddHandler wnd.SizeChanged, AddressOf wnd_SizeChanged
                        AddHandler wnd.LocationChanged, AddressOf wnd_LocationChanged
                    End If
                    If Not Objects.Contains(o) Then Objects.Add(o)
                    OnScreenKeyboardWindowTimer.IsEnabled = Objects.Count > 0
                End If
            End Sub))

#End Region

#Region "Helper methods"

    Private Shared WithEvents OnScreenKeyboardWindowTimer As New Threading.DispatcherTimer With {.Interval = TimeSpan.FromMilliseconds(500)}

    Private Const OnScreenKeyboardWindowClass = "IPTip_Main_Window"
    Private Const OnScreenKeyboardWindowParentClass1709 = "ApplicationFrameWindow"
    Private Const OnScreenKeyboardWindowClass1709 = "Windows.UI.Core.CoreWindow"
    Private Const OnScreenKeyboardWindowCaption1709 = "Microsoft Text Input Application"

    Private Shared OnScreenKeyboardRect As RECT
    Private Shared WindowRect As RECT
    Private Shared Objects As New List(Of Object)
    Private Shared OnScreenKeyboardOpened As Boolean

    Private Enum WindowStyle
        Disabled = &H8000000
        Visible = &H10000000
    End Enum

    <DllImport("user32.dll")>
    Private Shared Function GetWindowRect(ByVal hWnd As HandleRef, ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr,
                      ByVal childAfter As IntPtr,
                      ByVal lclassName As String,
                      ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowLong")>
    Private Shared Function GetWindowLongPtr32(ByVal hWnd As HandleRef, ByVal nIndex As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowLongPtr")>
    Private Shared Function GetWindowLongPtr64(ByVal hWnd As HandleRef, ByVal nIndex As Integer) As IntPtr
    End Function

    ' This static method is required because Win32 does not support GetWindowLongPtr dirctly
    Public Shared Function GetWindowLongPtr(ByVal hWnd As HandleRef, ByVal nIndex As Integer) As IntPtr
        If IntPtr.Size = 8 Then
            Return GetWindowLongPtr64(hWnd, nIndex)
        Else
            Return GetWindowLongPtr32(hWnd, nIndex)
        End If
    End Function

    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    Public Enum enmAction
        Edit
        [New]
    End Enum

    Private Shared Function GetIsOpen1709() As Boolean?
        Dim parent = IntPtr.Zero
        'Do
        parent = FindWindowEx(IntPtr.Zero, parent, OnScreenKeyboardWindowParentClass1709, Nothing)
        If parent = IntPtr.Zero Then Return Nothing ' no more windows, keyboard state Is unknown

        'if it's a child of a WindowParentClass1709 window - the keyboard is open
        Dim wnd = FindWindowEx(parent, IntPtr.Zero, OnScreenKeyboardWindowClass1709, OnScreenKeyboardWindowCaption1709)

        GetWindowRect(New HandleRef(Nothing, wnd), OnScreenKeyboardRect)

        If wnd <> IntPtr.Zero Then Return True
        'Loop
    End Function

    Private Shared Function GetIsOpenLegacy() As Boolean
        Dim wnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, OnScreenKeyboardWindowClass, Nothing)
        If wnd = IntPtr.Zero Then Return False

        GetWindowRect(New HandleRef(Nothing, wnd), OnScreenKeyboardRect)

        Dim style = GetWindowStyle(wnd)
        Return style.HasFlag(WindowStyle.Visible) AndAlso Not style.HasFlag(WindowStyle.Disabled)
    End Function

    Private Shared Function GetWindowStyle(wnd As IntPtr) As WindowStyle
        Return GetWindowLongPtr(New HandleRef(Nothing, wnd), -16)
    End Function

    Public Shared Function GetIsOpen() As Boolean
        Return If(GetIsOpen1709(), GetIsOpenLegacy())
    End Function

    Private Shared Sub timer_Tick(sender As Object, e As EventArgs) Handles OnScreenKeyboardWindowTimer.Tick
        OnScreenKeyboardOpened = GetIsOpen()
        UpdateSpacerHeight()
    End Sub

    Private Shared Sub wnd_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        UpdateSpacerHeight()
    End Sub

    Private Shared Sub wnd_LocationChanged(sender As Object, e As EventArgs)
        UpdateSpacerHeight()
    End Sub

    Private Shared Sub display_SettingsChanged(sender As Object, e As EventArgs)
        UpdateSpacerHeight()
    End Sub

    Private Shared Sub UpdateSpacerHeight()
        For Each o In Objects
            If OnScreenKeyboardOpened Then
                Dim wnd = Window.GetWindow(o) : If wnd Is Nothing Then Exit Sub
                If Not GetWindowRect(New HandleRef(Nothing, New WindowInteropHelper(wnd).Handle), WindowRect) Then Exit Sub
                If WindowRect.Bottom = 0 Or OnScreenKeyboardRect.Top = 0 Then Exit Sub
                Dim spacerheight As Integer = WindowRect.Bottom - OnScreenKeyboardRect.Top
                Dim spacerheightscaled As Integer

                If (SystemParameters.WorkArea.Width > SystemParameters.WorkArea.Height) Then
                    spacerheightscaled = spacerheight * SystemParameters.WorkArea.Width / Forms.Screen.PrimaryScreen.WorkingArea.Width 'dpiy
                Else
                    spacerheightscaled = spacerheight * SystemParameters.WorkArea.Height / Forms.Screen.PrimaryScreen.WorkingArea.Height 'dpix
                End If
                o.Height = If(spacerheightscaled > 0, spacerheightscaled, 0)
                o.Visibility = Visibility.Visible
            Else
                o.Visibility = Visibility.Collapsed
            End If
        Next
    End Sub

#End Region

End Class
