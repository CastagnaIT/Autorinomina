<Serializable>
Public Class ItemLinguaTVDB
    Implements System.ComponentModel.INotifyPropertyChanged

    Public Sub New()
        'necessario per la serializzazione
    End Sub

    Public Sub New(NomeLingua As String, Abbr As String)
        _NomeLingua = NomeLingua
        _Abbr = Abbr
    End Sub

    Public Property NomeLingua As String
        Get
            Return _NomeLingua
        End Get
        Set(value As String)
            _NomeLingua = value
            OnPropertyChanged("NomeLingua")
        End Set
    End Property

    Public Property Abbr As String
        Get
            Return _Abbr
        End Get
        Set(value As String)
            _Abbr = value
            OnPropertyChanged("Abbr")
        End Set
    End Property

    Private _NomeLingua As String
    Private _Abbr As String


    Protected Overridable Sub OnPropertyChanged(ByVal propertyName As String)
        Dim e As New System.ComponentModel.PropertyChangedEventArgs(propertyName)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

End Class

<Serializable>
Public Class CollItemsLinguaTVDB
    Inherits RangeObservableCollection(Of ItemLinguaTVDB)
End Class
