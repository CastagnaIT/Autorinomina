Imports System.Globalization

Public Class Converter_Chars
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim s = String.Empty

        If Not String.IsNullOrEmpty(value) Then
            s = value.ToString()

            If s.Contains("\t") Then
                s = s.Replace("\t", ControlChars.Tab)
            End If

            If s.Contains("\r\n") Then
                s = s.Replace("\r\n", Environment.NewLine)
            End If

            If s.Contains("\n") Then
                s = s.Replace("\n", Environment.NewLine)
            End If

            If s.Contains("&#x0a;&#x0d;") Then
                s = s.Replace("&#x0a;&#x0d;", Environment.NewLine)
            End If

            If s.Contains("&#x0a;") Then
                s = s.Replace("&#x0a;", Environment.NewLine)
            End If

            If s.Contains("&#x0d;") Then
                s = s.Replace("&#x0d;", Environment.NewLine)
            End If

            If s.Contains("&#10;&#13;") Then
                s = s.Replace("&#10;&#13;", Environment.NewLine)
            End If

            If s.Contains("&#10;") Then
                s = s.Replace("&#10;", Environment.NewLine)
            End If

            If s.Contains("&#13;") Then
                s = s.Replace("&#13;", Environment.NewLine)
            End If

            If s.Contains("<br />") Then
                s = s.Replace("<br />", Environment.NewLine)
            End If

            If s.Contains("<LineBreak />") Then
                s = s.Replace("<LineBreak />", Environment.NewLine)
            End If
        End If

        Return s
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
