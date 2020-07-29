Imports System.Data
Imports System.Text.RegularExpressions
Imports System.Windows.Threading
Imports TvDbSharper
Imports TvDbSharper.BaseSchemas
Imports TvDbSharper.Clients.Languages.Json

Public Class WND_OpzioniAvanzate
    Private Coll_Regex_SerieTv_Numerazione As New ObjectModel.ObservableCollection(Of Item_Regex)()
    Private Coll_Regex_TVDBLingue_Primaria As New CollItemsLinguaTVDB
    Private Coll_Regex_TVDBLingue_Secondaria As New CollItemsLinguaTVDB

    Dim tvdb_languages_file_path As String = DataPath & "\tvdb_languages.xml"
    Dim AvoidSave As String = True

    Class Item_Regex
        Public Sub New(_Regex As String, _Note As String)
            Regex = _Regex
            Note = _Note
        End Sub

        Public Property Regex As String
        Public Property Note As String
    End Class

    Private Sub WND_OpzioniAvanzate_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Select Case XMLSettings_Read("CategoriaSelezionata")
            Case "CATEGORIA_serietv"
                TC_OpzioniAvanzate.SelectedIndex = 0
            Case "CATEGORIA_video"
                TC_OpzioniAvanzate.SelectedIndex = 1
            Case "CATEGORIA_audio"
                TC_OpzioniAvanzate.SelectedIndex = 2
            Case "CATEGORIA_immagini"
                TC_OpzioniAvanzate.SelectedIndex = 3
            Case "CATEGORIA_generica"
                TC_OpzioniAvanzate.SelectedIndex = 4
        End Select


        'SerieTv
        CB_SerieTv_ModalitaFiltro.SelectedIndex = Integer.Parse(XMLSettings_Read("CATEGORIA_serietv_config_SceltaFiltro"))
        CB_SerieTv_ErroreFiltro.SelectedIndex = Integer.Parse(XMLSettings_Read("CATEGORIA_serietv_config_ErroreFiltro"))
        CB_SerieTv_RecognizeLinkedEpisode.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_serietv_config_RecognizeLinkedEpisode"))
        CB_SerieTv_EliminaParentesiEP.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_serietv_config_EliminaParentesiTitoloEpisodio"))
        CB_SerieTv_EliminaParentesiSerie.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_serietv_config_EliminaParentesiTitoloSerieTv"))
        CB_SerieTv_ConvertiNumRomani.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_serietv_config_ConvertiNumRomani"))

        Dim RegexNumerazione() As String = XMLSettings_Read("CATEGORIA_serietv_config_regex").Split(";")
        Dim RegexNumerazione_Note() As String = XMLSettings_Read("CATEGORIA_serietv_config_regex_note").Split(";")

        For x As Integer = 0 To RegexNumerazione.Length - 1
            Coll_Regex_SerieTv_Numerazione.Add(New Item_Regex(RegexNumerazione(x), RegexNumerazione_Note(x)))
        Next
        LB_SerieTv_Regex.ItemsSource = Coll_Regex_SerieTv_Numerazione

        'Video
        CB_Video_EliminaParentesi.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_video_config_EliminaParentesi"))
        CB_Video_WeakFilter.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_video_config_WeakFilter"))

        'Audio
        CB_Audio_NumerazMancante.SelectedIndex = Integer.Parse(XMLSettings_Read("CATEGORIA_audio_config_FallbackNumerazione"))
        CB_Audio_Numerazione_EscludiCaratteri.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_audio_config_EscludiCaratteriNumerazione"))
        CB_Audio_CercaNumerazInversa.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_audio_config_cercaNumerazioneRightToLeft"))
        CB_Audio_EliminaParentesi.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_audio_config_EliminaParentesi"))
        CB_Audio_WeakFilter.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_audio_config_WeakFilter"))
        CB_Audio_ID3Missing.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_audio_config_ID3TagMissing"))

        'Immagini
        CB_Immagini_NumerazMancante.SelectedIndex = Integer.Parse(XMLSettings_Read("CATEGORIA_immagini_config_FallbackNumerazione"))
        CB_Immagini_CercaNumerazInversa.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_immagini_config_cercaNumerazioneRightToLeft"))
        CB_Immagini_EliminaParentesi.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_immagini_config_EliminaParentesi"))
        CB_Immagini_WeakFilter.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_immagini_config_WeakFilter"))

        'Generica
        CB_Generica_NumerazMancante.SelectedIndex = Integer.Parse(XMLSettings_Read("CATEGORIA_generica_config_FallbackNumerazione"))
        CB_Generica_CercaNumerazInversa.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_generica_config_cercaNumerazioneRightToLeft"))
        CB_Generica_EliminaParentesi.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_generica_config_EliminaParentesi"))
        CB_Generica_WeakFilter.IsChecked = Boolean.Parse(XMLSettings_Read("CATEGORIA_generica_config_WeakFilter"))

        'TheTVDB
        CB_TVDB_Lingua_Principale.ItemsSource = Coll_Regex_TVDBLingue_Primaria
        CB_TVDB_Lingua_Secondaria.ItemsSource = Coll_Regex_TVDBLingue_Secondaria


        If IO.File.Exists(tvdb_languages_file_path) Then
            Try
                Coll_Regex_TVDBLingue_Primaria.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemLinguaTVDB)(tvdb_languages_file_path, "tvdb_languages"))
                Coll_Regex_TVDBLingue_Secondaria.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemLinguaTVDB)(tvdb_languages_file_path, "tvdb_languages"))

                CB_TVDB_Lingua_Principale.SelectedItem = Coll_Regex_TVDBLingue_Primaria.FirstOrDefault(Function(i) i.Abbr.Equals(XMLSettings_Read("TVDB_LinguaPrimaria")))
                CB_TVDB_Lingua_Secondaria.SelectedItem = Coll_Regex_TVDBLingue_Secondaria.FirstOrDefault(Function(i) i.Abbr.Equals(XMLSettings_Read("TVDB_LinguaSecondaria")))

                CB_TVDB_Lingua_Principale.IsEnabled = True
                CB_TVDB_Lingua_Secondaria.IsEnabled = True
                LB_TVDB_Lingua_Principale.IsEnabled = True
                LB_TVDB_Lingua_Secondaria.IsEnabled = True

            Catch ex As Exception
                Debug.Print(ex.Message)
                IO.File.Delete(tvdb_languages_file_path)
                HY_RicaricaLingue_Click(Me, Nothing)
            End Try
        Else
            HY_RicaricaLingue_Click(Me, Nothing)
        End If


        CB_TVDB_MetodoRicerca.SelectedIndex = IIf(XMLSettings_Read("TVDB_RicercaManuale").Equals("True"), 0, 1)
        CB_TVDB_NessunRisultato.SelectedIndex = IIf(XMLSettings_Read("TVDB_Risultati").Equals("True"), 0, 1)
        CB_TVDB_Lingua_Fallback.IsChecked = Boolean.Parse(XMLSettings_Read("TVDB_LinguaFallBack"))
        CB_TVDB_AssociazioneStagioni.SelectedIndex = Integer.Parse(XMLSettings_Read("TVDB_AssociazioneStagioni"))

        AvoidSave = False
    End Sub



    Private Sub HY_RicaricaLingue_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        If Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() Then
            TB_TVDB_RicaricaLingue.IsEnabled = False
            CB_TVDB_Lingua_Principale.IsEnabled = False
            CB_TVDB_Lingua_Secondaria.IsEnabled = False
            LB_TVDB_Lingua_Principale.IsEnabled = False
            LB_TVDB_Lingua_Secondaria.IsEnabled = False
            Coll_Regex_TVDBLingue_Primaria.Clear()
            Coll_Regex_TVDBLingue_Secondaria.Clear()

            Dim start As System.Threading.ThreadStart = Sub()
                                                            Try
                                                                'recupero lingue da tvdb
                                                                Dim tvDB As New TvDbClient()
                                                                tvDB.Authentication.AuthenticateAsync(TVDB_APIKEY).Wait()

                                                                Dim Cerca_Lingue As Task(Of TvDbResponse(Of Language())) = tvDB.Languages.GetAllAsync()
                                                                Cerca_Lingue.Wait()
                                                                Dim ListaLingue As Language() = Cerca_Lingue.Result.Data

                                                                For Each item In ListaLingue
                                                                    Dispatcher.Invoke(DispatcherPriority.Normal, Sub()
                                                                                                                     Coll_Regex_TVDBLingue_Primaria.Add(New ItemLinguaTVDB(item.EnglishName, item.Abbreviation))
                                                                                                                     Coll_Regex_TVDBLingue_Secondaria.Add(New ItemLinguaTVDB(item.EnglishName, item.Abbreviation))
                                                                                                                 End Sub)
                                                                Next


                                                                'eseguo fallback (se necessario) , carico preferenze, salvo su xml
                                                                Dispatcher.Invoke(DispatcherPriority.Normal, Sub()

                                                                                                                 If Coll_Regex_TVDBLingue_Primaria.Count = 0 Then 'fallback in caso di errore o connessione assente
                                                                                                                     Coll_Regex_TVDBLingue_Primaria.Add(New ItemLinguaTVDB("italian", "it"))
                                                                                                                     Coll_Regex_TVDBLingue_Secondaria.Add(New ItemLinguaTVDB("italian", "it"))

                                                                                                                     Coll_Regex_TVDBLingue_Primaria.Add(New ItemLinguaTVDB("english", "en"))
                                                                                                                     Coll_Regex_TVDBLingue_Secondaria.Add(New ItemLinguaTVDB("english", "en"))
                                                                                                                 End If

                                                                                                                 CB_TVDB_Lingua_Principale.SelectedItem = Coll_Regex_TVDBLingue_Primaria.FirstOrDefault(Function(i) i.Abbr.Equals(XMLSettings_Read("TVDB_LinguaPrimaria")))
                                                                                                                 CB_TVDB_Lingua_Secondaria.SelectedItem = Coll_Regex_TVDBLingue_Secondaria.FirstOrDefault(Function(i) i.Abbr.Equals(XMLSettings_Read("TVDB_LinguaSecondaria")))

                                                                                                                 CB_TVDB_Lingua_Principale.IsEnabled = True
                                                                                                                 CB_TVDB_Lingua_Secondaria.IsEnabled = True
                                                                                                                 LB_TVDB_Lingua_Principale.IsEnabled = True
                                                                                                                 LB_TVDB_Lingua_Secondaria.IsEnabled = True

                                                                                                                 XmlSerialization.WriteToXmlFile(tvdb_languages_file_path, Coll_Regex_TVDBLingue_Primaria, "tvdb_languages", False)
                                                                                                             End Sub)

                                                            Catch ex As Exception
                                                                Debug.Print(ex.Message)

                                                            Finally
                                                                Dispatcher.Invoke(DispatcherPriority.Normal, Sub()
                                                                                                                 TB_TVDB_RicaricaLingue.IsEnabled = True
                                                                                                             End Sub)
                                                            End Try
                                                        End Sub

            Dim Thread1 As New System.Threading.Thread(start)
            Thread1.Start()

        End If
    End Sub

    Private Sub BTN_Salva_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Salva.Click
        'SerieTv
        XMLSettings_Save("CATEGORIA_serietv_config_SceltaFiltro", CB_SerieTv_ModalitaFiltro.SelectedIndex.ToString)
        XMLSettings_Save("CATEGORIA_serietv_config_ErroreFiltro", CB_SerieTv_ErroreFiltro.SelectedIndex.ToString)
        XMLSettings_Save("CATEGORIA_serietv_config_RecognizeLinkedEpisode", CB_SerieTv_RecognizeLinkedEpisode.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_serietv_config_EliminaParentesiTitoloEpisodio", CB_SerieTv_EliminaParentesiEP.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_serietv_config_EliminaParentesiTitoloSerieTv", CB_SerieTv_EliminaParentesiSerie.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_serietv_config_ConvertiNumRomani", CB_SerieTv_ConvertiNumRomani.IsChecked.ToString)

        Dim RegexNumerazione As String = ""
        Dim RegexNumerazione_Note As String = ""
        For x As Integer = 0 To Coll_Regex_SerieTv_Numerazione.Count - 1
            Dim SplitChar As String = IIf((Coll_Regex_SerieTv_Numerazione.Count - 1) = x, "", ";")
            RegexNumerazione &= Coll_Regex_SerieTv_Numerazione(x).Regex & SplitChar
            RegexNumerazione_Note &= Coll_Regex_SerieTv_Numerazione(x).Note & SplitChar
        Next

        XMLSettings_Save("CATEGORIA_serietv_config_regex", RegexNumerazione)
        XMLSettings_Save("CATEGORIA_serietv_config_regex_note", RegexNumerazione_Note)

        'Video
        XMLSettings_Save("CATEGORIA_video_config_EliminaParentesi", CB_Video_EliminaParentesi.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_video_config_WeakFilter", CB_Video_WeakFilter.IsChecked.ToString)

        'Audio
        XMLSettings_Save("CATEGORIA_audio_config_FallbackNumerazione", CB_Audio_NumerazMancante.SelectedIndex.ToString)
        XMLSettings_Save("CATEGORIA_audio_config_EscludiCaratteriNumerazione", CB_Audio_Numerazione_EscludiCaratteri.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_audio_config_cercaNumerazioneRightToLeft", CB_Audio_CercaNumerazInversa.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_audio_config_EliminaParentesi", CB_Audio_EliminaParentesi.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_audio_config_WeakFilter", CB_Audio_WeakFilter.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_audio_config_ID3TagMissing", CB_Audio_ID3Missing.IsChecked.ToString)

        'Immagini
        XMLSettings_Save("CATEGORIA_immagini_config_FallbackNumerazione", CB_Immagini_NumerazMancante.SelectedIndex.ToString)
        XMLSettings_Save("CATEGORIA_immagini_config_cercaNumerazioneRightToLeft", CB_Immagini_CercaNumerazInversa.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_immagini_config_EliminaParentesi", CB_Immagini_EliminaParentesi.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_immagini_config_WeakFilter", CB_Immagini_WeakFilter.IsChecked.ToString)

        'Generica
        XMLSettings_Save("CATEGORIA_generica_config_FallbackNumerazione", CB_Generica_NumerazMancante.SelectedIndex.ToString)
        XMLSettings_Save("CATEGORIA_generica_config_cercaNumerazioneRightToLeft", CB_Generica_CercaNumerazInversa.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_generica_config_EliminaParentesi", CB_Generica_EliminaParentesi.IsChecked.ToString)
        XMLSettings_Save("CATEGORIA_generica_config_WeakFilter", CB_Generica_WeakFilter.IsChecked.ToString)

        'TVDB
        If Not String.IsNullOrEmpty(CB_TVDB_Lingua_Principale.Text) AndAlso Not CB_TVDB_Lingua_Principale.Text.ToLower.Equals("unknown") Then
            XMLSettings_Save("TVDB_LinguaPrimaria", CType(CB_TVDB_Lingua_Principale.SelectedItem, ItemLinguaTVDB).Abbr)
        End If
        If Not String.IsNullOrEmpty(CB_TVDB_Lingua_Secondaria.Text) AndAlso Not CB_TVDB_Lingua_Secondaria.Text.ToLower.Equals("unknown") Then
            XMLSettings_Save("TVDB_LinguaSecondaria", CType(CB_TVDB_Lingua_Secondaria.SelectedItem, ItemLinguaTVDB).Abbr)
        End If


        XMLSettings_Save("TVDB_RicercaManuale", IIf(CB_TVDB_MetodoRicerca.SelectedIndex = 0, "True", "False"))
        XMLSettings_Save("TVDB_Risultati", IIf(CB_TVDB_NessunRisultato.SelectedIndex = 0, "True", "False"))
        XMLSettings_Save("TVDB_LinguaFallBack", CB_TVDB_Lingua_Fallback.IsChecked.ToString)
        XMLSettings_Save("TVDB_AssociazioneStagioni", CB_TVDB_AssociazioneStagioni.SelectedIndex.ToString)

        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Annulla.Click
        Close()
    End Sub

    Private Sub BTN_SerieTv_Aggiungi_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SerieTv_Aggiungi.Click
        Dim result_Pattern = InputBox(Localization.Resource_OpzioniAvanzate.Msg_InsertPattern, BTN_SerieTv_Aggiungi.Content.ToString, "")
        Dim result_note = InputBox(Localization.Resource_OpzioniAvanzate.Msg_InsertPatternNote, BTN_SerieTv_Aggiungi.Content.ToString, "")

        If Not String.IsNullOrEmpty(result_Pattern.Trim) Then
            If Coll_Regex_SerieTv_Numerazione.FirstOrDefault(Function(i) i.Regex.Equals(result_Pattern)) Is Nothing Then
                Coll_Regex_SerieTv_Numerazione.Add(New Item_Regex(result_Pattern, result_note))
            End If
        End If

    End Sub

    Private Sub BTN_SerieTv_Elimina_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SerieTv_Elimina.Click
        If LB_SerieTv_Regex.SelectedItem IsNot Nothing Then
            Dim Result As MsgBoxResult = MsgBox(String.Format(Localization.Resource_OpzioniAvanzate.Msg_DeletePatternRegex, LB_SerieTv_Regex.SelectedItem.Regex), MsgBoxStyle.Question + MsgBoxStyle.YesNo, BTN_SerieTv_Elimina.Content.ToString)

            If Result = MsgBoxResult.Yes Then
                Coll_Regex_SerieTv_Numerazione.Remove(LB_SerieTv_Regex.SelectedItem)
            End If
        End If
    End Sub

    Private Sub BTN_SerieTv_Copia_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SerieTv_Copia.Click
        If LB_SerieTv_Regex.SelectedItem IsNot Nothing Then
            Clipboard.SetText(LB_SerieTv_Regex.SelectedItem.Regex)
        End If
    End Sub

    Private Sub BTN_SerieTv_Prova_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SerieTv_Prova.Click
        If LB_SerieTv_Regex.SelectedItem IsNot Nothing Then
            'Immetti il nome del file su cui eseguire il regex
            Dim Filename As String = InputBox(Localization.Resource_OpzioniAvanzate.Msg_PatternTest, BTN_SerieTv_Prova.Content.ToString, "")

            Dim Stagione As Integer = 0
            Dim Episodio As Integer = 0
            Dim EpisodioConcat As Integer = 0
            Try
                'Cerco la numerazione
                If Regex.IsMatch(Filename, LB_SerieTv_Regex.SelectedItem.Regex, RegexOptions.IgnoreCase) Then

                    Dim m As Match = Regex.Match(Filename, LB_SerieTv_Regex.SelectedItem.Regex, RegexOptions.IgnoreCase)
                    If m.Groups("season").Value() <> "" Then Stagione = m.Groups("season").Value()
                    If m.Groups("episode").Value() <> "" Then Episodio = m.Groups("episode").Value()
                    If m.Groups("episode2").Value() <> "" Then EpisodioConcat = m.Groups("episode2").Value()

                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, BTN_SerieTv_Prova.Content.ToString)
            Finally

                Dim Result As String = "Result:" & vbCrLf &
                    "Season:" & Stagione & vbCrLf &
                    "Episode:" & Episodio & vbCrLf &
                    "Next Episode (if exist):" & EpisodioConcat
                If Stagione = 0 AndAlso Episodio = 0 AndAlso EpisodioConcat = 0 Then
                    Result = "No result. Check regex and/or the filename."
                End If

                MsgBox(Result, MsgBoxStyle.Information, BTN_SerieTv_Prova.Content.ToString)
            End Try
        End If
    End Sub

    Private Sub BTN_SerieTv_SpostaSu_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SerieTv_SpostaSu.Click
        If LB_SerieTv_Regex.SelectedIndex > 0 AndAlso LB_SerieTv_Regex.SelectedItems.Count = 1 Then

            Dim index As Integer = LB_SerieTv_Regex.SelectedIndex

            Coll_Regex_SerieTv_Numerazione.Move(index, index - 1)
            LB_SerieTv_Regex.ScrollIntoView(LB_SerieTv_Regex.SelectedItem)
        End If
    End Sub

    Private Sub BTN_SerieTv_SpostaGiu_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SerieTv_SpostaGiu.Click
        If LB_SerieTv_Regex.SelectedIndex < LB_SerieTv_Regex.Items.Count - 1 AndAlso LB_SerieTv_Regex.SelectedItems.Count = 1 Then

            Dim index As Integer = LB_SerieTv_Regex.SelectedIndex

            Coll_Regex_SerieTv_Numerazione.Move(index, index + 1)
            LB_SerieTv_Regex.ScrollIntoView(LB_SerieTv_Regex.SelectedItem)
        End If
    End Sub

    Private Sub BTN_SerieTv_Pattern_Help_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SerieTv_Pattern_Help.Click
        Try
            System.Diagnostics.Process.Start("http://www.autorinomina.it/index.php/tutorials/elenco-dei-tutorial/17-regex-per-riconoscimento-numerazione-serie-tv")
        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, BTN_SerieTv_Pattern_Help.Content.ToString)
        End Try
    End Sub
End Class
