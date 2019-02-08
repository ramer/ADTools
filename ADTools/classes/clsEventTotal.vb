Imports System.Collections.ObjectModel

Public Class clsEventTotal
    Private _day As Date
    Private _first As Date?
    Private _last As Date?
    Private _events As New clsThreadSafeObservableCollection(Of clsEvent)

    Sub New()

    End Sub

    Sub New(evts As clsThreadSafeObservableCollection(Of clsEvent), day As Date)
        _events = evts
        _day = day
    End Sub

    Public ReadOnly Property Image() As BitmapImage
        Get
            Dim _image As String = ""

            If Diff.HasValue AndAlso Diff.Value.Hours >= 9 Then
                _image = "images/ok.png"
            ElseIf Diff.HasValue AndAlso Diff.Value.Hours < 9 Then
                _image = "images/warning.png"
            Else
                _image = "images/puzzle.png"
            End If

            Return New BitmapImage(New Uri("pack://application:,,,/" & _image))
        End Get
    End Property

    Public ReadOnly Property Day As Date
        Get
            Return _day
        End Get
    End Property

    Public ReadOnly Property First As Date?
        Get
            If _first IsNot Nothing Then Return _first

            For Each evt In _events
                If _first Is Nothing OrElse evt.TimeGenerated < _first Then _first = evt.TimeGenerated
            Next

            Return _first
        End Get
    End Property

    Public ReadOnly Property Last As Date?
        Get
            If _last IsNot Nothing Then Return _last

            For Each evt In _events
                If _last Is Nothing OrElse evt.TimeGenerated > _last Then _last = evt.TimeGenerated
            Next

            Return _last
        End Get
    End Property

    Public ReadOnly Property Diff As TimeSpan?
        Get
            Return If(First IsNot Nothing And Last IsNot Nothing, Last - First, Nothing)
        End Get
    End Property

End Class
