<Serializable>
Public Class ItemTermine
    Implements System.ComponentModel.INotifyPropertyChanged

    Public Sub New()
        'necessario per la serializzazione
    End Sub

    Public Sub New(ByVal termine As String, ByVal termineSostituto As String, ByVal caseSensitive As String)
        _termine = termine
        _termineSostituto = termineSostituto
        _caseSensitive = caseSensitive
    End Sub

    Public Property Termine As String
        Get
            Return _termine
        End Get
        Set(value As String)
            _termine = value
            OnPropertyChanged("Termine")
        End Set
    End Property

    Public Property TermineSostituto As String
        Get
            Return _termineSostituto
        End Get
        Set(value As String)
            _termineSostituto = value
            OnPropertyChanged("TermineSostituto")
        End Set
    End Property

    Public Property CaseSensitive As String
        Get
            Return _caseSensitive
        End Get
        Set(value As String)
            _caseSensitive = value
            OnPropertyChanged("CaseSensitive")
        End Set
    End Property

    Private _termine As String
    Private _termineSostituto As String
    Private _caseSensitive As String


    Protected Overridable Sub OnPropertyChanged(ByVal propertyName As String)
        Dim e As New System.ComponentModel.PropertyChangedEventArgs(propertyName)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

End Class

<Serializable>
Public Class CollItemsTermini
    Inherits RangeObservableCollection(Of ItemTermine)
End Class
