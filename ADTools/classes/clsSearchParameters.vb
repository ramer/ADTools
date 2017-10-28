Imports System.Collections.ObjectModel

Public Class clsSearchParameters

    Private returncollection As clsThreadSafeObservableCollection(Of clsDirectoryObject)
    Private pattern As String
    Private domain As clsDomain
    Private parentobject As clsDirectoryObject

    Private specificdomains As ObservableCollection(Of clsDomain) = Nothing
    Private attributes As ObservableCollection(Of clsAttribute) = Nothing
    Private searchobjectclasses As clsSearchObjectClasses = Nothing
    Private freesearch As Boolean = False

End Class
