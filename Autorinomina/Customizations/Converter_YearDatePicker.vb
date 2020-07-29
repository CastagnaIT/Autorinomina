Public Class Converter_YearDatePicker
    Implements IValueConverter

    Public Function Convert(ByVal value As Object, ByVal targetType As System.Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements System.Windows.Data.IValueConverter.Convert
        If value Is Nothing Then Return Now.Date.Year

        Return CDate(value).Year
    End Function

    Public Function ConvertBack(ByVal value As Object, ByVal targetType As System.Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements System.Windows.Data.IValueConverter.ConvertBack
        Debug.Print("cc " & value)
        Dim data As Date
        Try
            If IsNumeric(value) = False Then
                data = Now
            Else
                data = New Date(CInt(value), 1, 1)
            End If
        Catch ex As Exception
            data = Now
        End Try

        Return data
    End Function
End Class
