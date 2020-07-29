Public Class Converter_SpaceToSymbol
    Implements IValueConverter

    Public Function Convert(ByVal value As Object, ByVal targetType As System.Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements System.Windows.Data.IValueConverter.Convert
        If value Is Nothing Then Return ""

        Return Text.RegularExpressions.Regex.Replace(value.ToString, "^(\s+)|(\s+)$", "▪")
    End Function

    Public Function ConvertBack(ByVal value As Object, ByVal targetType As System.Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements System.Windows.Data.IValueConverter.ConvertBack
        If value Is Nothing Then Return ""

        Return Text.RegularExpressions.Regex.Replace(value.ToString, "(▪+)", " ")
    End Function
End Class
