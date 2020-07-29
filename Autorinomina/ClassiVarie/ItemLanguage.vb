Public Class ItemLanguage
    Implements System.ComponentModel.INotifyPropertyChanged

    Public Sub New(ByVal nome As String, ByVal autore As String, ByVal versione As String, ByVal englishName As String)
        _nome = nome
        _autore = autore
        _versione = versione
        _englishName = englishName
    End Sub

    Public Property Nome As String
        Get
            Return _nome
        End Get
        Set(value As String)
            _nome = value
            OnPropertyChanged("Nome")
        End Set
    End Property

    Public Property Autore As String
        Get
            Return _autore
        End Get
        Set(value As String)
            _autore = value
            OnPropertyChanged("Autore")
        End Set
    End Property

    Public Property Versione As String
        Get
            Return _versione
        End Get
        Set(value As String)
            _versione = value
            OnPropertyChanged("Versione")
        End Set
    End Property

    Public Property EnglishName As String
        Get
            Return _englishName
        End Get
        Set(value As String)
            _englishName = value
            OnPropertyChanged("EnglishName")
        End Set
    End Property

    Private _nome As String
    Private _autore As String
    Private _versione As String
    Private _englishName As String

    Protected Overridable Sub OnPropertyChanged(ByVal propertyName As String)
        Dim e As New System.ComponentModel.PropertyChangedEventArgs(propertyName)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

End Class

<Serializable>
Public Class CollItemsLanguages
    Inherits RangeObservableCollection(Of ItemLanguage)
End Class
