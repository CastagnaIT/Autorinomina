Imports System.ComponentModel
Imports System.Text.RegularExpressions

Public Class DLG_TheTVDB_Manuale
    WithEvents BW_Ricerca As New BackgroundWorker

    Public Property ResultData As String() = Nothing
    Dim NomeSerieTv As String = ""
    Private Collection_Temp As New ObjectModel.ObservableCollection(Of Coll_Temp)()
    Dim tvDB As TvdbLib.TvdbHandler

    Class Coll_Temp
        Public Sub New(_IDSerieTv As String, _NomeSerieTv As String, _Lingua As String, _LinguaAbbr As String)
            IDSerieTv = _IDSerieTv
            NomeSerieTv = _NomeSerieTv
            Lingua = _Lingua
            LinguaAbbr = _LinguaAbbr
        End Sub

        Public Property IDSerieTv As String
        Public Property NomeSerieTv As String
        Public Property Lingua As String
        Public Property LinguaAbbr As String
    End Class

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(_NomeSerieTv As String)

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().

        _NomeSerieTv = Regex.Replace(_NomeSerieTv.Trim, "[^a-zA-Z\s\d]", "") 'elimino caratteri speciali
        NomeSerieTv = _NomeSerieTv

        tvDB = New TvdbLib.TvdbHandler(New TvdbLib.Cache.XmlCacheProvider(DataPath & "Cache"), TVDB_APIKEY)
        tvDB.InitCache()

        BackgroundWait()
    End Sub

    Private Sub BackgroundWait()
        Dim VB As New VisualBrush
        VB.TileMode = TileMode.None
        VB.Stretch = Stretch.None
        Dim R As New Run
        R.Foreground = Brushes.Gray
        R.BaselineAlignment = BaselineAlignment.Center
        R.Text = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_TVDB_Desc4 & NomeSerieTv
        R.FontStyle = FontStyles.Italic
        R.FontSize = 13
        VB.Visual = New TextBlock(R)
        LV_ElencoRisultati.Background = VB
    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        TB_NomeSerieTv.Text = NomeSerieTv
        TB_CambiaRicerca.Text = NomeSerieTv

        Dim CVS As CollectionViewSource = FindResource("CVS_Risultati")
        CVS.Source = Collection_Temp
        CVS.SortDescriptions.Add(New SortDescription("NomeSerieTv", ListSortDirection.Ascending))
        CVS.SortDescriptions.Add(New SortDescription("Lingua", ListSortDirection.Ascending))

        BW_Ricerca.RunWorkerAsync(NomeSerieTv)
    End Sub



    Private Sub BTN_Ok_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Ok.Click
        If LV_ElencoRisultati.SelectedItem Is Nothing Then
            MsgBox(Localization.Resource_Common_Dialogs.DLG_TVDB_Msg_SelectResult, MsgBoxStyle.Exclamation, Me.Title)
            Return
        End If

        ResultData = {CType(LV_ElencoRisultati.SelectedItem, Coll_Temp).IDSerieTv, CType(LV_ElencoRisultati.SelectedItem, Coll_Temp).LinguaAbbr}

        Me.DialogResult = True
        Close()
    End Sub

    Private Sub BTN_Salta_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Salta.Click
        ResultData = {"-1", ""}
        Me.DialogResult = False
        Close()
    End Sub

    Private Sub BTN_Interrompi_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Interrompi.Click
        ResultData = {Nothing, ""}
        Me.DialogResult = False
        Close()
    End Sub

    Private Sub BTN_EseguiRicerca_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_EseguiRicerca.Click
        BTN_EseguiRicerca.IsEnabled = False
        Collection_Temp.Clear()
        BackgroundWait()
        BW_Ricerca.RunWorkerAsync(TB_CambiaRicerca.Text)
    End Sub


    Private Sub BW_Ricerca_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_Ricerca.DoWork
        Dim NomeSerieTv As String = e.Argument

        Dim ListaLingue As List(Of TvdbLib.Data.TvdbLanguage) = tvDB.Languages
        Dim LinguaPrimaria As TvdbLib.Data.TvdbLanguage = ListaLingue.Find(Function(x) x.Abbriviation = XMLSettings_Read("TVDB_LinguaPrimaria"))
        Dim LinguaSecondaria As TvdbLib.Data.TvdbLanguage = ListaLingue.Find(Function(x) x.Abbriviation = XMLSettings_Read("TVDB_LinguaSecondaria"))

        Dim ListaSerie As List(Of TvdbLib.Data.TvdbSearchResult) = tvDB.SearchSeries(NomeSerieTv, LinguaPrimaria)
        For Each item As TvdbLib.Data.TvdbSearchResult In ListaSerie
            If item.Language = LinguaPrimaria OrElse item.Language = LinguaSecondaria Then
                Dim item_S As TvdbLib.Data.TvdbSearchResult = item
                Dispatcher.BeginInvoke(Windows.Threading.DispatcherPriority.Normal, Sub()
                                                                                        Collection_Temp.Add(New Coll_Temp(item_S.Id, item_S.SeriesName, item_S.Language.Name, item_S.Language.Abbriviation))
                                                                                    End Sub)
            End If
        Next

        If LinguaPrimaria <> LinguaSecondaria Then

            ListaSerie = tvDB.SearchSeries(NomeSerieTv, LinguaSecondaria)
            For Each item As TvdbLib.Data.TvdbSearchResult In ListaSerie
                If item.Language = LinguaPrimaria OrElse item.Language = LinguaSecondaria Then
                    Dim item_S As TvdbLib.Data.TvdbSearchResult = item
                    Dispatcher.BeginInvoke(Windows.Threading.DispatcherPriority.Normal, Sub()
                                                                                            Collection_Temp.Add(New Coll_Temp(item_S.Id, item_S.SeriesName, item_S.Language.Name, item_S.Language.Abbriviation))
                                                                                        End Sub)
                End If
            Next
        End If
    End Sub

    Private Sub BW_Ricerca_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW_Ricerca.RunWorkerCompleted
        If Collection_Temp.Count = 0 Then
            Dim VB As New VisualBrush
            VB.TileMode = TileMode.None
            VB.Stretch = Stretch.None
            Dim R As New Run
            R.Foreground = Brushes.Gray
            R.BaselineAlignment = BaselineAlignment.Center
            R.Text = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_TVDB_NoResults
            R.FontStyle = FontStyles.Italic
            R.FontSize = 13
            VB.Visual = New TextBlock(R)
            LV_ElencoRisultati.Background = VB
        Else
            LV_ElencoRisultati.Background = Nothing
        End If

        BTN_EseguiRicerca.IsEnabled = True
    End Sub

    Private Sub TB_CambiaRicerca_PreviewKeyDown(sender As System.Object, e As System.Windows.Input.KeyEventArgs) Handles TB_CambiaRicerca.PreviewKeyDown
        If e.Key = Key.Enter Then
            If BTN_EseguiRicerca.IsEnabled Then
                BTN_EseguiRicerca_Click(Me, Nothing)
                e.Handled = True
            End If
        End If
    End Sub

End Class
