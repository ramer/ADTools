Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports CredentialManagement
Imports IRegisty
Imports Microsoft.Win32

Public Class clsPreferences
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private _firstrun As Boolean = False
    Private _firstversionrun As Boolean = False

    ' basic
    Private _clipboardsource As Boolean = False
    Private _clipboardsourcelimit As Boolean = True
    Private _viewshownavigation As Boolean = True
    Private _viewshowpreview As Boolean = True
    Private _viewresultgrouping As Boolean = False
    Private _viewshowdeletedobjects As Boolean = False

    ' layout
    Private _columns As New ObservableCollection(Of clsViewColumnInfo)
    Private _defaultview As enmView

    ' search attributes
    Private _attributesforsearch As New ObservableCollection(Of clsAttribute)

    ' behavior
    Private _startwithwindows As Boolean?
    Private _startwithwindowsminimized As Boolean?
    Private _closeonxbutton As Boolean?

    ' appearance
    Private _colortext As Color
    Private _colorwindowbackground As Color
    Private _colorelementbackground As Color
    Private _colormenubackground As Color
    Private _colorbuttonbackground As Color
    Private _colorbuttoninactivebackground As Color
    Private _colorlistviewrow As Color
    Private _colorlistviewalternationrow As Color

    ' externalsoftware
    Private _externalsoftware As New ObservableCollection(Of clsExternalSoftware)

    ' favorites
    Private _favorites As New ObservableCollection(Of clsDirectoryObject)

    ' saved filters
    Private _filters As New ObservableCollection(Of clsFilter)

    Private Sub NotifyPropertyChanged(propertyName As String)
        Me.OnPropertyChanged(New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Public Enum enmView
        Details = 0
        Tiles = 1
        List = 2
        MediumIcons = 3
    End Enum

    Sub New()

    End Sub

    Public Property LastVersion As String
        Get
            Return My.Application.Info.Version.Major.ToString() & "." & My.Application.Info.Version.Minor.ToString()
        End Get
        Set(value As String)
            _firstrun = (value Is Nothing)
            _firstversionrun = (My.Application.Info.Version.Major.ToString() & "." & My.Application.Info.Version.Minor.ToString() <> value)
            NotifyPropertyChanged("FirstRun")
            NotifyPropertyChanged("ShowHints")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property FirstRun As String
        Get
            Return _firstrun
        End Get
    End Property

    <RegistrySerializerIgnorable(True)>
    Public ReadOnly Property FirstVersionRun As String
        Get
            Return _firstversionrun
        End Get
    End Property

    Public Property ClipboardSource As Boolean
        Get
            Return _clipboardsource
        End Get
        Set(value As Boolean)
            _clipboardsource = value
            NotifyPropertyChanged("ClipboardSource")
        End Set
    End Property

    Public Property ClipboardSourceLimit As Boolean
        Get
            Return _clipboardsourcelimit
        End Get
        Set(value As Boolean)
            _clipboardsourcelimit = value
            NotifyPropertyChanged("ClipboardSourceLimit")
        End Set
    End Property

    Public Property ViewShowNavigation As Boolean
        Get
            Return _viewshownavigation
        End Get
        Set(value As Boolean)
            _viewshownavigation = value
            NotifyPropertyChanged("ViewShowNavigation")
        End Set
    End Property

    Public Property ViewShowPreview As Boolean
        Get
            Return _viewshowpreview
        End Get
        Set(value As Boolean)
            _viewshowpreview = value
            NotifyPropertyChanged("ViewShowPreview")
        End Set
    End Property

    Public Property ViewResultGrouping As Boolean
        Get
            Return _viewresultgrouping
        End Get
        Set(value As Boolean)
            _viewresultgrouping = value
            NotifyPropertyChanged("ViewResultGrouping")
        End Set
    End Property

    Public Property ViewShowDeletedObjects As Boolean
        Get
            Return _viewshowdeletedobjects
        End Get
        Set(value As Boolean)
            _viewshowdeletedobjects = value
            NotifyPropertyChanged("ViewShowDeletedObjects")
        End Set
    End Property

    Public Property Columns() As ObservableCollection(Of clsViewColumnInfo)
        Get
            Return _columns
        End Get
        Set(value As ObservableCollection(Of clsViewColumnInfo))
            _columns = If(value, GetDefaultColumns())
            NotifyPropertyChanged("Columns")
        End Set
    End Property

    Public Property ViewValue As Integer
        Get
            Return _defaultview
        End Get
        Set(value As Integer)
            _defaultview = value
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property DefaultView As enmView
        Get
            Return _defaultview
        End Get
        Set(value As enmView)
            _defaultview = value
            NotifyPropertyChanged("DefaultView")
        End Set
    End Property

    Public Property AttributesForSearch() As ObservableCollection(Of clsAttribute)
        Get
            Return _attributesforsearch
        End Get
        Set(value As ObservableCollection(Of clsAttribute))
            _attributesforsearch = If(value, attributesForSearchDefault)
            NotifyPropertyChanged("AttributesForSearch")
        End Set
    End Property

    Public Property StartWithWindows As Boolean?
        Get
            Return _startwithwindows
        End Get
        Set(value As Boolean?)
            _startwithwindows = If(value, False)
            NotifyPropertyChanged("StartWithWindows")
        End Set
    End Property

    Public Property StartWithWindowsMinimized As Boolean?
        Get
            Return _startwithwindowsminimized
        End Get
        Set(value As Boolean?)
            _startwithwindowsminimized = If(value, True)
            NotifyPropertyChanged("StartWithWindowsMinimized")
        End Set
    End Property

    Public Property CloseOnXButton As Boolean?
        Get
            Return _closeonxbutton
        End Get
        Set(value As Boolean?)
            _closeonxbutton = If(value, True)
            NotifyPropertyChanged("CloseOnXButton")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ColorText As Color
        Get
            Return _colortext
        End Get
        Set(value As Color)
            _colortext = value
            Application.Current.Resources("ColorText") = New SolidColorBrush(_colortext)
            NotifyPropertyChanged("ColorText")
        End Set
    End Property

    <RegistrySerializerAlias("ColorText")>
    Public Property ColorTextValue As String
        Get
            Return ColorText.ToString
        End Get
        Set(value As String)
            ColorText = ColorConverter.ConvertFromString(If(String.IsNullOrEmpty(value), Colors.Black.ToString, value))
            NotifyPropertyChanged("ColorTextValue")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ColorWindowBackground As Color
        Get
            Return _colorwindowbackground
        End Get
        Set(value As Color)
            _colorwindowbackground = value
            Application.Current.Resources("ColorWindowBackground") = New SolidColorBrush(_colorwindowbackground)
            NotifyPropertyChanged("ColorWindowBackground")
        End Set
    End Property

    <RegistrySerializerAlias("ColorWindowBackground")>
    Public Property ColorWindowBackgroundValue As String
        Get
            Return ColorWindowBackground.ToString
        End Get
        Set(value As String)
            ColorWindowBackground = ColorConverter.ConvertFromString(If(String.IsNullOrEmpty(value), Colors.WhiteSmoke.ToString, value))
            NotifyPropertyChanged("ColorWindowBackgroundValue")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ColorElementBackground As Color
        Get
            Return _colorelementbackground
        End Get
        Set(value As Color)
            _colorelementbackground = value
            Application.Current.Resources("ColorElementBackground") = New SolidColorBrush(_colorelementbackground)
            NotifyPropertyChanged("ColorElementBackground")
        End Set
    End Property

    <RegistrySerializerAlias("ColorElementBackground")>
    Public Property ColorElementBackgroundValue As String
        Get
            Return ColorElementBackground.ToString
        End Get
        Set(value As String)
            ColorElementBackground = ColorConverter.ConvertFromString(If(String.IsNullOrEmpty(value), Colors.White.ToString, value))
            NotifyPropertyChanged("ColorElementBackgroundValue")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ColorMenuBackground As Color
        Get
            Return _colormenubackground
        End Get
        Set(value As Color)
            _colormenubackground = value
            Application.Current.Resources("ColorMenuBackground") = New SolidColorBrush(_colormenubackground)
            NotifyPropertyChanged("ColorMenuBackground")
        End Set
    End Property

    <RegistrySerializerAlias("ColorMenuBackground")>
    Public Property ColorMenuBackgroundValue As String
        Get
            Return ColorMenuBackground.ToString
        End Get
        Set(value As String)
            ColorMenuBackground = ColorConverter.ConvertFromString(If(String.IsNullOrEmpty(value), Colors.WhiteSmoke.ToString, value))
            NotifyPropertyChanged("ColorMenuBackgroundValue")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ColorButtonBackground As Color
        Get
            Return _colorbuttonbackground
        End Get
        Set(value As Color)
            _colorbuttonbackground = value
            Application.Current.Resources("ColorButtonBackground") = New SolidColorBrush(_colorbuttonbackground)
            NotifyPropertyChanged("ColorButtonBackground")
        End Set
    End Property

    <RegistrySerializerAlias("ColorButtonBackground")>
    Public Property ColorButtonBackgroundValue As String
        Get
            Return ColorButtonBackground.ToString
        End Get
        Set(value As String)
            ColorButtonBackground = ColorConverter.ConvertFromString(If(String.IsNullOrEmpty(value), Colors.LightSkyBlue.ToString, value))
            NotifyPropertyChanged("ColorButtonBackgroundValue")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ColorButtonInactiveBackground As Color
        Get
            Return _colorbuttoninactivebackground
        End Get
        Set(value As Color)
            _colorbuttoninactivebackground = value
            Application.Current.Resources("ColorButtonInactiveBackground") = New SolidColorBrush(_colorbuttoninactivebackground)
            NotifyPropertyChanged("ColorButtonInactiveBackground")
        End Set
    End Property

    <RegistrySerializerAlias("ColorButtonInactiveBackground")>
    Public Property ColorButtonInactiveBackgroundValue As String
        Get
            Return ColorButtonInactiveBackground.ToString
        End Get
        Set(value As String)
            ColorButtonInactiveBackground = ColorConverter.ConvertFromString(If(String.IsNullOrEmpty(value), "#FFD2EBFB", value))
            NotifyPropertyChanged("ColorButtonInactiveBackgroundValue")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ColorListviewRow As Color
        Get
            Return _colorlistviewrow
        End Get
        Set(value As Color)
            _colorlistviewrow = value
            Application.Current.Resources("ColorListviewRow") = New SolidColorBrush(_colorlistviewrow)
            NotifyPropertyChanged("ColorListviewRow")
        End Set
    End Property

    <RegistrySerializerAlias("ColorListviewRow")>
    Public Property ColorListviewRowValue As String
        Get
            Return ColorListviewRow.ToString
        End Get
        Set(value As String)
            ColorListviewRow = ColorConverter.ConvertFromString(If(String.IsNullOrEmpty(value), Colors.White.ToString, value))
            NotifyPropertyChanged("ColorListviewRowValue")
        End Set
    End Property

    <RegistrySerializerIgnorable(True)>
    Public Property ColorListviewAlternationRow As Color
        Get
            Return _colorlistviewalternationrow
        End Get
        Set(value As Color)
            _colorlistviewalternationrow = value
            Application.Current.Resources("ColorListviewAlternationRow") = New SolidColorBrush(_colorlistviewalternationrow)
            NotifyPropertyChanged("ColorListviewAlternationRow")
        End Set
    End Property

    <RegistrySerializerAlias("ColorListviewAlternationRow")>
    Public Property ColorListviewAlternationRowValue As String
        Get
            Return ColorListviewAlternationRow.ToString
        End Get
        Set(value As String)
            ColorListviewAlternationRow = ColorConverter.ConvertFromString(If(String.IsNullOrEmpty(value), Colors.AliceBlue.ToString, value))
            NotifyPropertyChanged("ColorListviewAlternationRowValue")
        End Set
    End Property

    Public Property ExternalSoftware As ObservableCollection(Of clsExternalSoftware)
        Get
            Return _externalsoftware
        End Get
        Set(value As ObservableCollection(Of clsExternalSoftware))
            _externalsoftware = If(value, New ObservableCollection(Of clsExternalSoftware))
            NotifyPropertyChanged("ExternalSoftware")
        End Set
    End Property

    Public Property Favorites As ObservableCollection(Of clsDirectoryObject)
        Get
            Return _favorites
        End Get
        Set(value As ObservableCollection(Of clsDirectoryObject))
            _favorites = If(value, New ObservableCollection(Of clsDirectoryObject))
            NotifyPropertyChanged("Favorites")
        End Set
    End Property

    Public Property Filters As ObservableCollection(Of clsFilter)
        Get
            Return _filters
        End Get
        Set(value As ObservableCollection(Of clsFilter))
            _filters = If(value, New ObservableCollection(Of clsFilter))
            NotifyPropertyChanged("Filters")
        End Set
    End Property


End Class
