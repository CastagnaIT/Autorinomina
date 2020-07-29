Public Class ItemTVDBmanual
    Implements System.ComponentModel.INotifyPropertyChanged

    Public Sub New()
        'necessario per la serializzazione
    End Sub

    Public Sub New(ByVal IDSerieTv As String, ByVal NomeSerieTv As String, ByVal FirstAired As String, Network As String, Lingua As String, LinguaAbbr As String)
        _IDSerieTv = IDSerieTv
        _NomeSerieTv = NomeSerieTv
        _FirstAired = FirstAired
        _Network = Network
        _Lingua = Lingua
        _LinguaAbbr = LinguaAbbr
    End Sub


    Public Property IDSerieTv As String
        Get
            Return _IDSerieTv
        End Get
        Set(value As String)
            _IDSerieTv = value
            OnPropertyChanged("IDSerieTv")
        End Set
    End Property

    Public Property NomeSerieTv As String
        Get
            Return _NomeSerieTv
        End Get
        Set(value As String)
            _NomeSerieTv = value
            OnPropertyChanged("NomeSerieTv")
        End Set
    End Property

    Public Property FirstAired As String
        Get
            Return _FirstAired
        End Get
        Set(value As String)
            _FirstAired = value
            OnPropertyChanged("FirstAired")
        End Set
    End Property

    Public Property Network As String
        Get
            Return _Network
        End Get
        Set(value As String)
            _Network = value
            OnPropertyChanged("Network")
        End Set
    End Property

    Public Property Lingua As String
        Get
            Return _Lingua
        End Get
        Set(value As String)
            _Lingua = value
            OnPropertyChanged("Lingua")
        End Set
    End Property

    Public Property LinguaAbbr As String
        Get
            Return _LinguaAbbr
        End Get
        Set(value As String)
            _LinguaAbbr = value
            OnPropertyChanged("LinguaAbbr")
        End Set
    End Property

    Private _IDSerieTv As String
    Private _NomeSerieTv As String
    Private _FirstAired As String
    Private _Network As String
    Private _Lingua As String
    Private _LinguaAbbr As String

    Protected Overridable Sub OnPropertyChanged(ByVal propertyName As String)
        Dim e As New System.ComponentModel.PropertyChangedEventArgs(propertyName)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

End Class

Public Class CollItemsTVDBmanual
    Inherits ObjectModel.ObservableCollection(Of ItemTVDBmanual)
End Class
