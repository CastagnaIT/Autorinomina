Public Class Converter_InfoStatoToImage
    Implements IValueConverter

    Public Function Convert(ByVal value As Object, ByVal targetType As System.Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements System.Windows.Data.IValueConverter.Convert
        If value Is Nothing Then Return Nothing

        Select Case Integer.Parse(value)
            Case FileDataInfoStato.NULLO, FileDataInfoStato.RIPRISTINATO
                Return "pack://application:,,,/Immagini/iconeStato/status_nullo.png"
            Case FileDataInfoStato.ANTEPRIMA_OK
                Return "pack://application:,,,/Immagini/iconeStato/status_anteprima.png"
            Case FileDataInfoStato.CONFLITTO
                Return "pack://application:,,,/Immagini/iconeStato/status_warning.png"
            Case FileDataInfoStato.ERRORE, FileDataInfoStato.ANTEPRIMA_ERRORE
                Return "pack://application:,,,/Immagini/iconeStato/status_errore.png"
            Case FileDataInfoStato.RINOMINATO
                Return "pack://application:,,,/Immagini/iconeStato/status_ok.png"
        End Select

        Return Nothing
    End Function

    Public Function ConvertBack(ByVal value As Object, ByVal targetType As System.Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements System.Windows.Data.IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
