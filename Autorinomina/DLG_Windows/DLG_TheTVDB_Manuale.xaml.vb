Imports System.ComponentModel
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports TvDbSharper
Imports TvDbSharper.BaseSchemas
Imports TvDbSharper.Clients.Languages.Json
Imports TvDbSharper.Clients.Search.Json

Public Class DLG_TheTVDB_Manuale
    Public Property ResultData As String() = Nothing
    Dim NomeSerieTv As String = ""
    Dim CVS As CollectionViewSource
    Private Collection_Temp As New CollItemsTVDBmanual
    WithEvents BW_Ricerca As New BackgroundWorker
    Dim tvDB As New TvDbClient

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(_NomeSerieTv As String)

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().

        _NomeSerieTv = Regex.Replace(_NomeSerieTv.Trim, "[^a-zA-Z\s\d]", "") 'elimino caratteri speciali
        NomeSerieTv = _NomeSerieTv

        BackgroundWait(_NomeSerieTv)
    End Sub

    Private Sub BackgroundWait(bkgrd_txt As String)
        Dim VB As New VisualBrush
        VB.TileMode = TileMode.None
        VB.Stretch = Stretch.None
        Dim R As New Run
        R.Foreground = Brushes.Gray
        R.BaselineAlignment = BaselineAlignment.Center
        R.Text = Autorinomina.Localization.Resource_Common_Dialogs.DLG_TVDB_Desc4 & bkgrd_txt
        R.FontStyle = FontStyles.Italic
        R.FontSize = 13
        VB.Visual = New TextBlock(R)
        LV_ElencoRisultati.Background = VB
    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        TB_NomeSerieTv.Text = NomeSerieTv
        TB_CambiaRicerca.Text = NomeSerieTv

        CVS = FindResource("CVS_Risultati")
        CVS.Source = Collection_Temp
        ' CVS.SortDescriptions.Add(New SortDescription("Lingua", ListSortDirection.Ascending))
        CVS.SortDescriptions.Add(New SortDescription("NomeSerieTv", ListSortDirection.Ascending))


        tvDB.Authentication.AuthenticateAsync(TVDB_APIKEY).Wait()

        BW_Ricerca.RunWorkerAsync(NomeSerieTv)
    End Sub

    Private Sub BW_Ricerca_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_Ricerca.DoWork
        Dim NomeSerieTv As String = e.Argument
        Dim TempColl As New CollItemsTVDBmanual

        Try
            'definisco lingue da cercare
            Dim RichiestaElencoLingue As Task(Of TvDbResponse(Of Clients.Languages.Json.Language())) = tvDB.Languages.GetAllAsync()
            RichiestaElencoLingue.Wait()
            Dim ElencoLingue As Language() = RichiestaElencoLingue.Result.Data

            Dim LinguaPrimaria As Clients.Languages.Json.Language = ElencoLingue.FirstOrDefault(Function(i) i.Abbreviation = XMLSettings_Read("TVDB_LinguaPrimaria"))
            Dim LinguaSecondaria As Clients.Languages.Json.Language = ElencoLingue.FirstOrDefault(Function(i) i.Abbreviation = XMLSettings_Read("TVDB_LinguaSecondaria"))


            'cerco con lingua primaria
            tvDB.AcceptedLanguage = LinguaPrimaria.Abbreviation
            Dim Cerca_Lang_Primaria As Task(Of TvDbResponse(Of SeriesSearchResult())) = tvDB.Search.SearchSeriesByNameAsync(NomeSerieTv)
            Cerca_Lang_Primaria.Wait()
            Dim ListaSerieLngPrimaria As SeriesSearchResult() = Cerca_Lang_Primaria.Result.Data

            For Each item In ListaSerieLngPrimaria
                TempColl.Add(New ItemTVDBmanual(item.Id, item.SeriesName, FormattaData(item.FirstAired), item.Network, LinguaPrimaria.EnglishName, LinguaPrimaria.Abbreviation))
            Next


            'cerco con lingua secondaria
            If Not LinguaPrimaria.Abbreviation.Equals(LinguaSecondaria.Abbreviation) Then

                tvDB.AcceptedLanguage = LinguaSecondaria.Abbreviation
                Dim Cerca_Lang_Secondaria As Task(Of TvDbResponse(Of SeriesSearchResult())) = tvDB.Search.SearchSeriesByNameAsync(NomeSerieTv)
                Cerca_Lang_Secondaria.Wait()
                Dim ListaSerieLngSecondaria As SeriesSearchResult() = Cerca_Lang_Secondaria.Result.Data

                For Each item In ListaSerieLngSecondaria
                    TempColl.Add(New ItemTVDBmanual(item.Id, item.SeriesName, FormattaData(item.FirstAired), item.Network, LinguaSecondaria.EnglishName, LinguaSecondaria.Abbreviation))
                Next

            End If
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

        e.Result = TempColl
    End Sub

    Private Sub BW_Ricerca_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW_Ricerca.RunWorkerCompleted
        If (e.Error IsNot Nothing) Then
            'nothing
        Else
            If e.Cancelled Then
                'nothing
            Else
                Dim Result As CollItemsTVDBmanual = e.Result

                Using CVS.DeferRefresh
                    For Each item In Result
                        Collection_Temp.Add(item)
                    Next
                End Using

                If Collection_Temp.Count = 0 Then
                    Dim VB As New VisualBrush
                    VB.TileMode = TileMode.None
                    VB.Stretch = Stretch.None
                    Dim R As New Run
                    R.Foreground = Brushes.Gray
                    R.BaselineAlignment = BaselineAlignment.Center
                    R.Text = Autorinomina.Localization.Resource_Common_Dialogs.DLG_TVDB_NoResults
                    R.FontStyle = FontStyles.Italic
                    R.FontSize = 13
                    VB.Visual = New TextBlock(R)
                    LV_ElencoRisultati.Background = VB
                Else
                    LV_ElencoRisultati.Background = Nothing
                End If

                BTN_EseguiRicerca.IsEnabled = True
            End If
        End If
    End Sub


    Private Sub BTN_Ok_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Ok.Click
        If LV_ElencoRisultati.SelectedItem Is Nothing Then
            MsgBox(Localization.Resource_Common_Dialogs.DLG_TVDB_Msg_SelectResult, MsgBoxStyle.Exclamation, Me.Title)
            Return
        End If

        ResultData = {CType(LV_ElencoRisultati.SelectedItem, ItemTVDBmanual).IDSerieTv, CType(LV_ElencoRisultati.SelectedItem, ItemTVDBmanual).LinguaAbbr}

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

        BackgroundWait(TB_CambiaRicerca.Text)

        BW_Ricerca.RunWorkerAsync(TB_CambiaRicerca.Text)
    End Sub

    Private Function FormattaData(DataOriginale As String) As String
        If Not String.IsNullOrEmpty(DataOriginale) Then
            Try
                Dim Data As Date = Date.ParseExact(DataOriginale, "yyyy-MM-dd", New CultureInfo("en-US"))
                Return Data.ToLocalTime.ToShortDateString
            Catch ex As Exception
                Debug.Print(ex.Message)
                Return DataOriginale
            End Try
        End If

        Return DataOriginale
    End Function

    Private Sub TB_CambiaRicerca_PreviewKeyDown(sender As System.Object, e As System.Windows.Input.KeyEventArgs) Handles TB_CambiaRicerca.PreviewKeyDown
        If e.Key = Key.Enter Then
            If BTN_EseguiRicerca.IsEnabled Then
                BTN_EseguiRicerca_Click(Me, Nothing)
                e.Handled = True
            End If
        End If
    End Sub

End Class
