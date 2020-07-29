Imports System.Collections.ObjectModel
Imports System.Runtime.InteropServices
Imports AutoRinomina.Interaction

Public Class ItemFileData
    Implements System.ComponentModel.INotifyPropertyChanged
    'non uso INotifyPropertyChanged non ho trovato un modo per fare un delay refresh UI quando si modifica tramite codice i dati
    'questo causerebbe un refresh continuo del datagrid

    Public Sub New(ByVal percorsoFile As String, index As Integer)

        Dim FI As New IO.FileInfo(percorsoFile)

        _percorso = FI.Directory.FullName
        _nomeFile = FI.Name
        _nomeFileRinominato = ""
        _dataCreazione = FI.CreationTime.ToString
        _dataUltimaModifica = FI.LastWriteTime.ToString
        _dimensione = FormatBYTE(FI.Length)
        Dim TipoFile As String = ""
        Try
            Dim shfi As New SHFILEINFO
            SHGetFileInfo(percorsoFile, 0, shfi, Marshal.SizeOf(shfi), SHGFI_TYPENAME)
            TipoFile = shfi.szTypeName
            If String.IsNullOrEmpty(TipoFile) Then TipoFile = "File " & FI.Extension.Remove(0, 1).ToUpper

        Catch ex As Exception
            TipoFile = "File " & FI.Extension.Remove(0, 1).ToUpper
        End Try
        _tipo = TipoFile

        Dim NomeCartella As String = ""
        Try
            NomeCartella = FI.Directory.Name
        Catch ex As Exception
            NomeCartella = IO.Path.GetDirectoryName(FI.Directory.FullName)
        End Try
        _nomeCartella = NomeCartella

        _stato = FileDataInfoStato.NULLO
        _statoInfo = ""
        _index = index
    End Sub

    Public Sub New(ByVal percorso As String, ByVal nomeFile As String, ByVal nomeFileRinominato As String, ByVal dataCreazione As String, ByVal dataUltimaModifica As String, ByVal dimensione As String, ByVal tipo As String, ByVal nomeCartella As String, ByVal stato As Integer, ByVal statoInfo As String, ByVal index As Integer)
        _percorso = percorso
        _nomeFile = nomeFile
        _nomeFileRinominato = nomeFileRinominato
        _dataCreazione = dataCreazione
        _dataUltimaModifica = dataUltimaModifica
        _dimensione = dimensione
        _tipo = tipo
        _nomeCartella = nomeCartella
        _stato = stato
        _statoInfo = statoInfo
        _index = index
    End Sub

    Public Property Percorso As String
        Get
            Return _percorso
        End Get
        Set(value As String)
            _percorso = value
            '   OnPropertyChanged("Percorso")
        End Set
    End Property

    Public Property NomeFile As String
        Get
            Return _nomeFile
        End Get
        Set(value As String)
            _nomeFile = value
            '  OnPropertyChanged("NomeFile")
        End Set
    End Property

    Public Property NomeFileRinominato As String
        Get
            Return _nomeFileRinominato
        End Get
        Set(value As String)
            _nomeFileRinominato = value
            CheckOnPropertyChanged("NomeFileRinominato")
            '  OnPropertyChanged("NomeFileRinominato")
        End Set
    End Property

    Public Property DataCreazione As String
        Get
            Return _dataCreazione
        End Get
        Set(value As String)
            _dataCreazione = value
            '   OnPropertyChanged("DataCreazione")
        End Set
    End Property

    Public Property DataUltimaModifica As String
        Get
            Return _dataUltimaModifica
        End Get
        Set(value As String)
            _dataUltimaModifica = value
            '  OnPropertyChanged("DataUltimaModifica")
        End Set
    End Property

    Public Property Dimensione As String
        Get
            Return _dimensione
        End Get
        Set(value As String)
            _dimensione = value
            '  OnPropertyChanged("Dimensione")
        End Set
    End Property

    Public Property Tipo As String
        Get
            Return _tipo
        End Get
        Set(value As String)
            _tipo = value
            '  OnPropertyChanged("Tipo")
        End Set
    End Property

    Public Property NomeCartella As String
        Get
            Return _nomeCartella
        End Get
        Set(value As String)
            _nomeCartella = value
            ' OnPropertyChanged("NomeCartella")
        End Set
    End Property

    Public Property Stato As Integer
        Get
            Return _stato
        End Get
        Set(value As Integer)
            _stato = value
            CheckOnPropertyChanged("Stato")
        End Set
    End Property

    Public Property StatoInfo As String
        Get
            Return _statoInfo
        End Get
        Set(value As String)
            _statoInfo = value
            CheckOnPropertyChanged("StatoInfo")
        End Set
    End Property

    Public Property Index As Integer
        Get
            Return _index
        End Get
        Set(value As Integer)
            _index = value
            CheckOnPropertyChanged("Index")
        End Set
    End Property

    Private _percorso As String
    Private _nomeFile As String
    Private _nomeFileRinominato As String
    Private _dataCreazione As String
    Private _dataUltimaModifica As String
    Private _dimensione As String
    Private _tipo As String
    Private _nomeCartella As String
    Private _stato As Integer
    Private _statoInfo As String
    Private _index As Integer

    Private Sub CheckOnPropertyChanged(ByVal propertyName As String)
        If Coll_Files_OnPropertyChanged_Enabled Then OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub OnPropertyChanged(ByVal propertyName As String)
        Dim e As New System.ComponentModel.PropertyChangedEventArgs(propertyName)
        RaiseEvent PropertyChanged(Me, e)
    End Sub
    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
End Class

Public Class CollItemsFilesData
    Inherits ObservableCollectionEx(Of ItemFileData)
End Class
