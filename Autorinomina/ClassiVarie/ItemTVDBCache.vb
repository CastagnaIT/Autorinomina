Imports System.Collections.ObjectModel
Imports TvDbSharper.Clients
Imports TvDbSharper.Clients.Series.Json

Public Class ItemTVDBCache

    Public Sub New(ByVal ID_SerieTv As String, ByVal NStagione As Integer, ByVal Episodi As BasicEpisode(), ByVal Serie As Series)
        _ID_SerieTv = ID_SerieTv
        _NStagione = NStagione
        _Episodi = Episodi
        _Serie = Serie
    End Sub

    Public Property ID_SerieTv As String
        Get
            Return _ID_SerieTv
        End Get
        Set(value As String)
            _ID_SerieTv = value
        End Set
    End Property

    Public Property NStagione As Integer
        Get
            Return _NStagione
        End Get
        Set(value As Integer)
            _NStagione = value
        End Set
    End Property

    Public Property Episodi As BasicEpisode()
        Get
            Return _Episodi
        End Get
        Set(value As BasicEpisode())
            _Episodi = value
        End Set
    End Property

    Public Property Serie As Series
        Get
            Return _Serie
        End Get
        Set(value As Series)
            _Serie = value
        End Set
    End Property

    Private _ID_SerieTv As String
    Private _NStagione As Integer
    Private _Episodi As BasicEpisode()
    Private _Serie As Series

End Class

<Serializable>
Public Class CollItemsTVDBCache
    Inherits ObservableCollection(Of ItemTVDBCache)
End Class