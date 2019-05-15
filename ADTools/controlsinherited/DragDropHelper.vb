Public Class DragDropHelper

    Public Shared ReadOnly IsDragOverProperty As DependencyProperty = DependencyProperty.RegisterAttached("IsDragOver", GetType(Boolean), GetType(DragDropHelper), New PropertyMetadata(False))

    Public Shared Sub SetIsDragOver(ByVal element As DependencyObject, ByVal value As Boolean)
        element.SetValue(IsDragOverProperty, value)
    End Sub

    Public Shared Function GetIsDragOver(ByVal element As DependencyObject) As Boolean
        Return CBool(element.GetValue(IsDragOverProperty))
    End Function

End Class