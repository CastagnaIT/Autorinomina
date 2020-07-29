Imports System.Windows.Controls.Primitives

Public Class YearDatePicker
    Inherits DatePicker
    Dim Txt As DatePickerTextBox



    Protected Overrides Sub OnCalendarOpened(e As RoutedEventArgs)
        Dim popup = TryCast(Me.Template.FindName("PART_Popup", Me), Popup)
        If popup IsNot Nothing AndAlso TypeOf popup.Child Is System.Windows.Controls.Calendar Then
            DirectCast(popup.Child, Calendar).DisplayMode = CalendarMode.Decade
        End If

        If IsDropDownOpen = False Then Me.IsDropDownOpen = True
        AddHandler DirectCast(popup.Child, Calendar).DisplayModeChanged, New EventHandler(Of CalendarModeChangedEventArgs)(AddressOf DatePickerCo_DisplayModeChanged)
    End Sub

    Private Sub DatePickerCo_DisplayModeChanged(sender As Object, e As CalendarModeChangedEventArgs)
        Dim popup = TryCast(Me.Template.FindName("PART_Popup", Me), Popup)
        If popup IsNot Nothing AndAlso TypeOf popup.Child Is System.Windows.Controls.Calendar Then
            Dim _calendar = TryCast(popup.Child, System.Windows.Controls.Calendar)

            If IsDropDownOpen Then
                Me.SelectedDate = New DateTime(_calendar.DisplayDate.Year, 1, 1)
                Me.IsDropDownOpen = False
            End If

            RemoveHandler DirectCast(popup.Child, Calendar).DisplayModeChanged, New EventHandler(Of CalendarModeChangedEventArgs)(AddressOf DatePickerCo_DisplayModeChanged)

            _calendar.DisplayMode = CalendarMode.Decade
        End If
    End Sub

End Class
