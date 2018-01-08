Imports System.ComponentModel

Public Class pgGroup
    Public Shared ReadOnly CurrentObjectProperty As DependencyProperty = DependencyProperty.Register("CurrentObject",
                                                            GetType(clsDirectoryObject),
                                                            GetType(pgGroup),
                                                            New FrameworkPropertyMetadata(Nothing, AddressOf CurrentObjectPropertyChanged))

    Private Property _currentobject As clsDirectoryObject
    Private Property _currentdomainobjects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Public Property CurrentObject() As clsDirectoryObject
        Get
            Return GetValue(CurrentObjectProperty)
        End Get
        Set(ByVal value As clsDirectoryObject)
            SetValue(CurrentObjectProperty, value)
        End Set
    End Property

    Private Shared Sub CurrentObjectPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As pgGroup = CType(d, pgGroup)
        With instance
            ._currentobject = CType(e.NewValue, clsDirectoryObject)
        End With
    End Sub

    Private Sub hlManagedBy_Click(sender As Object, e As RoutedEventArgs) Handles hlManagedBy.Click
        ShowDirectoryObjectProperties(CurrentObject.managedBy, Window.GetWindow(Me))
    End Sub

    Private Sub pgGroup_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
    End Sub

End Class

'Глобальная группа может быть членом другой глобальной группы, универсальной группы или локальной группы домена.
'Универсальная группа может быть членом другой универсальной группы или локальной группы домена, но не может быть членом глобальной группы.
'Локальная группа домена может быть членом только другой локальной группы домена.
'Локальную группу домена можно преобразовать в универсальную группу лишь в том случае, если эта локальная группа домена не содержит 
'    других членов локальной группы домена. Локальная группа домена не может быть членом универсальной группы.
'Глобальную группу можно преобразовать в универсальную лишь в том случае, если эта глобальная группа не входит в состав другой глобальной группы.
'    Универсальная группа не может быть членом глобальной группы.