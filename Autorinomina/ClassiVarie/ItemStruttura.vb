Public Class ItemStruttura
    Implements System.ComponentModel.INotifyPropertyChanged

    Public Sub New(ByVal tipoDato As String, ByVal testo As String, ByVal opzioni As String, ByVal coloreTesto As Brush)
        _tipoDato = tipoDato
        _testo = testo
        _opzioni = opzioni
        _coloreTesto = coloreTesto
    End Sub

    Public Property TipoDato As String
        Get
            Return _tipoDato
        End Get
        Set(value As String)
            _tipoDato = value
            OnPropertyChanged("TipoDato")
        End Set
    End Property

    Public Property Testo As String
        Get
            Return _testo
        End Get
        Set(value As String)
            _testo = value
            OnPropertyChanged("Testo")
        End Set
    End Property

    Public Property Opzioni As String
        Get
            Return _opzioni
        End Get
        Set(value As String)
            _opzioni = value
            OnPropertyChanged("Opzioni")
        End Set
    End Property

    Public Property ColoreTesto As Brush
        Get
            Return _coloreTesto
        End Get
        Set(value As Brush)
            _coloreTesto = value
            OnPropertyChanged("ColoreTesto")
        End Set
    End Property

    Private _tipoDato As String
    Private _testo As String
    Private _opzioni As String
    Private _coloreTesto As Brush

    Protected Overridable Sub OnPropertyChanged(ByVal propertyName As String)
        Dim e As New System.ComponentModel.PropertyChangedEventArgs(propertyName)
        RaiseEvent PropertyChanged(Me, e)
    End Sub


    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

End Class

Public Class CollItemsStruttura
    Inherits RangeObservableCollection(Of ItemStruttura)
End Class
