Imports System.Text.RegularExpressions

Public  Class AR_Core_SerieTv
    Inherits AR_Core

    Dim Coll_TerminiBlackList As New CollItemsTermini
    Dim Coll_TerminiSostituzioni As New CollItemsTermini
    Dim tvDB As TvdbLib.TvdbHandler

    Dim TVDB_Dizionario_Trovati As New Dictionary(Of String, String())

    Dim Filtro_Tipo As String = "#"
    Dim Filtro_ApplyOnAll As Boolean = False

    Private Enum FINFO_TVDB
        TITOLOSERIETV
        TITOLOEPISODIO
        DATAPRIMATV
        CREATOR
        DIRECTOR
        NUMEROEPISODI
        NUMEROSTAGIONI
        GENERE
        NETWORK
        Count
    End Enum

    Public Sub New(_OwnerWindow As Window)
        MyBase.New(_OwnerWindow)

        Coll_TerminiBlackList.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & "BlackList_CATEGORIA_serietv.xml", "BlackList"))
        Coll_TerminiSostituzioni.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & "Sostituzioni_CATEGORIA_serietv.xml", "Sostituzioni"))
        Debug.Print("load")
        If (InfoRichiesteStruttura.Contains("TVDB_")) Then tvDB = New TvdbLib.TvdbHandler(New TvdbLib.Cache.XmlCacheProvider(IO.Path.Combine(DataPath, "cache")), TVDB_APIKEY)
    End Sub

    Private Delegate Function Delegate_GeneralFunc(ByVal ParameterToPass As String) As String()

    Private Function Open_DLGSceltaFiltro(ByVal PathFile As String) As String()
        Dim WND As New DLG_SceltaManuale(PathFile)
        WND.ShowDialog()
        Return WND.ResultData
    End Function

    Private Function Open_DLGModificaManuale(ByVal PathFile As String) As String()
        Dim WND As New DLG_ModificaManuale(PathFile)
        WND.ShowDialog()
        Return WND.ResultData
    End Function

    Private Function OpenDLGtvDB_Manuale(ByVal PathFile As String) As String()
        Dim WND As New DLG_TheTVDB_Manuale(PathFile)
        WND.ShowDialog()
        Return WND.ResultData
    End Function


    Public Overrides Function RunCore(ByVal PathFile As String, ByVal NumeroSequenziale As Integer, ByVal LivePreview As Boolean, ByVal PaddingLenght As Integer, ByVal Optional overrideTipoFiltro As String = Nothing) As String
        'InfoFiltrate index informazioni: 0=Numerazione | 1=Titolo episodio | 2=Titolo serietv | 3=data | 4=validazione
        Dim InfoFiltrate() As String = Nothing

        Dim Info_MI() As String = Nothing
        Dim Info_PF() As String = Nothing
        Dim Info_TVDB() As String = Nothing
        Dim Estensione As String = IO.Path.GetExtension(PathFile)

        'Scelta del filtro
        Dim config_SceltaFiltro As String = IIf(LivePreview, 0, XMLSettings_Read("CATEGORIA_serietv_config_SceltaFiltro"))
        If Filtro_ApplyOnAll OrElse Not InfoRichiesteStruttura.Contains("AR_") Then config_SceltaFiltro = "0"

        Select Case config_SceltaFiltro
            Case "1", "2" 'Consenti scelta del filtro solo in caso di errore, oppure, consenti sempre (case 2)
                InfoFiltrate = Filtro_SerieTv(IO.Path.GetFileName(PathFile), Filtro_Tipo)

                If InfoFiltrate(4).Equals("False") OrElse config_SceltaFiltro.Equals("2") Then
                    Dim del As New Delegate_GeneralFunc(AddressOf Open_DLGSceltaFiltro)
                    Dim Result() As String = OwnerWindow.Dispatcher.Invoke(del, PathFile)

                    If String.IsNullOrEmpty(Result(0)) Then Throw New UserInterruptedException

                    Filtro_ApplyOnAll = Boolean.Parse(Result(2))
                    If (Filtro_ApplyOnAll) Then Filtro_Tipo = Result(1)

                    InfoFiltrate = Filtro_SerieTv(IO.Path.GetFileName(PathFile), Result(1))
                End If

            Case Else 'Scelta filtro automatica
                InfoFiltrate = Filtro_SerieTv(IO.Path.GetFileName(PathFile), IIf(String.IsNullOrEmpty(overrideTipoFiltro), Filtro_Tipo, overrideTipoFiltro))
        End Select

        'Nel caso il filtro non passasse la validazione
        If InfoRichiesteStruttura.Contains("AR_") AndAlso InfoFiltrate(4).Equals("False") Then
            If XMLSettings_Read("CATEGORIA_serietv_config_ErroreFiltro").Equals("0") AndAlso LivePreview = False Then
                'Apri finestra per correggere manualmente il file
                Dim del As New Delegate_GeneralFunc(AddressOf Open_DLGModificaManuale)
                Dim Result() As String = OwnerWindow.Dispatcher.Invoke(del, PathFile)

                If String.IsNullOrEmpty(Result(0)) Then Throw New UserInterruptedException
                Return Result(1)
            Else
                'Non rinominare il file
                Return ""
            End If
        End If

        If InfoRichiesteStruttura.Contains("MI_") Then Info_MI = GetVideoInfo(PathFile)
        If InfoRichiesteStruttura.Contains("PF_") Then Info_PF = GetPropertyInfo(PathFile)

        If InfoRichiesteStruttura.Contains("TVDB_") AndAlso LivePreview = False Then

            'Estrapolo il numero della serie tv e dell'episodio
            Dim TvDB_DatiSelezionati As KeyValuePair(Of String, String()) = New KeyValuePair(Of String, String())("-1", {"-1", ""})
            Dim RisultatoDLG() As String = {Nothing}
            Dim Stagione As Integer = 1
            Dim Episodio As Integer = 1
            If InfoFiltrate(0).Contains("x") Then
                Stagione = Integer.Parse(Regex.Replace(InfoFiltrate(0).Split("x")(0), "[^0-9]", ""))
                Episodio = Integer.Parse(Regex.Replace(InfoFiltrate(0).Split("x")(1), "[^0-9]", ""))
            Else
                Episodio = Integer.Parse(Regex.Replace(InfoFiltrate(0), "[^0-9]", ""))
            End If

            'Scelta dei risultati manuale, si apre solo quando la serietv cercata non è già stata aggiunta nel dizionario
            If XMLSettings_Read("TVDB_RicercaManuale").Equals("True") Then
                If Not TVDB_Dizionario_Trovati.ContainsKey(InfoFiltrate(2)) Then
                    Dim Result() As String = {"-1", ""}
                    'se l'id serie non è presente
                    'scelta manuale dai risultati di TheTVDB
                    If LivePreview = False Then
                        Dim del As New Delegate_GeneralFunc(AddressOf OpenDLGtvDB_Manuale)
                        Result = OwnerWindow.Dispatcher.Invoke(del, InfoFiltrate(2))
                    End If
                    'Se in Result(0) l'ID è -1 == ricerca automatica tbDB

                    If String.IsNullOrEmpty(Result(0)) Then Throw New UserInterruptedException
                    TvDB_DatiSelezionati = New KeyValuePair(Of String, String())(InfoFiltrate(2), Result)
                Else
                    'l'id è già presente, recupero le info
                    TvDB_DatiSelezionati = New KeyValuePair(Of String, String())(InfoFiltrate(2), TVDB_Dizionario_Trovati.Item(InfoFiltrate(2)))
                End If

            End If

            Dim IT As ItemStruttura = Coll_Struttura.FirstOrDefault(Function(i) i.TipoDato.Equals("TVDB_TitoloEpisodio"))

            Dim ordineEpisodi As Integer = 0
            If Not IT Is Nothing Then ordineEpisodi = LeggiStringOpzioni("OrdineEpisodi", IT.Opzioni)

            Info_TVDB = Get_TVDB_Info(TvDB_DatiSelezionati.Value(0), TvDB_DatiSelezionati.Value(1), Episodio, Stagione, InfoFiltrate(2), ordineEpisodi)

            If Info_TVDB Is Nothing AndAlso XMLSettings_Read("TVDB_Risultati").Equals("False") Then
                'nel caso non ci siano risultati su TheTVDB non rinominare
                Return ""
            End If

        ElseIf LivePreview Then
            'dati falsi per la live preview
            Dim Val(FINFO_TVDB.Count - 1) As String

            Val(FINFO_TVDB.TITOLOEPISODIO) = Localization.Resource_Struttura.TVDB_TitoloEpisodio
            Val(FINFO_TVDB.TITOLOSERIETV) = Localization.Resource_Struttura.TVDB_TitoloSerieTv
            Val(FINFO_TVDB.DATAPRIMATV) = Localization.Resource_Struttura.TVDB_DataPrimaTv
            Val(FINFO_TVDB.CREATOR) = Localization.Resource_Struttura.TVDB_Creator
            Val(FINFO_TVDB.DIRECTOR) = Localization.Resource_Struttura.TVDB_Director
            Val(FINFO_TVDB.NUMEROEPISODI) = Localization.Resource_Struttura.TVDB_NumeroEpisodi
            Val(FINFO_TVDB.NUMEROSTAGIONI) = Localization.Resource_Struttura.TVDB_NumeroStagioni
            Val(FINFO_TVDB.GENERE) = Localization.Resource_Struttura.TVDB_Genere
            Val(FINFO_TVDB.NETWORK) = Localization.Resource_Struttura.TVDB_Network

            Info_TVDB = Val
        End If


        Dim StileMaiuscoleMinuscole As Integer = Integer.Parse(XMLSettings_Read("CATEGORIA_serietv_config_StileMaiuscoleMinuscole"))
        Dim TestoPrecedeUnSeparatore As Boolean = False 'Serve per conversione maiuscole/minuscole modalità 4 (Prima lettera maiuscola dopo ogni separatore resto minuscolo) 
        Dim PrimaFraseOparola As Boolean = True 'Serve per capire qual'è la parola iniziale da fare la capitalizzazione
        Dim NomeRinominato As String = ""

        For Each IT As ItemStruttura In Coll_Struttura
            Dim NuovoContenuto As String = ""
            Dim IgnoraConversioneMM As Boolean = False 'Ignora conversione maiuscole/minuscole

            Select Case IT.TipoDato
                Case "STD_Testo"
                    NuovoContenuto = LeggiStringOpzioni("Testo", IT.Opzioni)

                Case "STD_Separatore"
                    If LeggiStringOpzioni("SeparatoreSpaziatura", IT.Opzioni).Equals("True") Then
                        NuovoContenuto = Space(1) & LeggiStringOpzioni("SeparatoreCarattere", IT.Opzioni) & Space(1)
                    Else
                        NuovoContenuto = LeggiStringOpzioni("SeparatoreCarattere", IT.Opzioni)
                    End If
                    IgnoraConversioneMM = True

                Case "STD_ElencoSequenziale"
                    If ContenutoElencoSequenziale.Length >= (NumeroSequenziale + 1) Then
                        NuovoContenuto = ContenutoElencoSequenziale(NumeroSequenziale)
                    Else
                        Throw New ARCoreException(Autorinomina.Localization.Resource_Common.AR_Core_Error_SeqTextIncomplete)
                    End If

                Case "STD_NumerazioneSequenziale"
                    'Crea una nuova numerazione
                    Dim _PaddingLenght As Integer = PaddingLenght
                    If LeggiStringOpzioni("NumerazioneAutoPadding", IT.Opzioni).Equals("False") Then
                        _PaddingLenght = Integer.Parse(LeggiStringOpzioni("NumerazionePadding", IT.Opzioni))
                    Else
                        _PaddingLenght = IIf(_PaddingLenght <= 1, 2, _PaddingLenght)
                    End If

                    Dim NewBaseNum As String = "-1"
                    Select Case LeggiStringOpzioni("NumerazioneStile", IT.Opzioni, "5")
                        Case "5", "6", "7", "8"
                            NewBaseNum = LeggiStringOpzioni("NumerazioneStagione", IT.Opzioni) & "x" & (Integer.Parse(LeggiStringOpzioni("NumerazioneIniziaDa", IT.Opzioni)) + NumeroSequenziale)
                        Case Else
                            NewBaseNum = (Integer.Parse(LeggiStringOpzioni("NumerazioneIniziaDa", IT.Opzioni)) + NumeroSequenziale).ToString
                    End Select

                    NuovoContenuto = CreaNuovaNumerazione(NewBaseNum, LeggiStringOpzioni("NumerazioneStile", IT.Opzioni, "5"), _PaddingLenght)
                    NuovoContenuto = LeggiStringOpzioni("NumerazionePrefisso", IT.Opzioni) & NuovoContenuto & LeggiStringOpzioni("NumerazioneSuffisso", IT.Opzioni)
                    IgnoraConversioneMM = True

                Case "AR_Numerazione"
                    'Filtra numerazione dal file (padding automatico)
                    Dim _PaddingLenght As Integer = IIf(PaddingLenght < 10, 2, PaddingLenght)
                    If InfoFiltrate(0).Contains("x") Then
                        Dim tmp As String() = InfoFiltrate(0).Split("x")

                        If tmp(1).Contains("-") Then
                            'Contiene un doppio episodio
                            _PaddingLenght = tmp(1).Split("-")(0).Length
                        Else
                            _PaddingLenght = tmp(1).Length
                        End If
                    End If

                    NuovoContenuto = CreaNuovaNumerazione(InfoFiltrate(0), LeggiStringOpzioni("NumerazioneStile", IT.Opzioni, "5"), _PaddingLenght)
                    IgnoraConversioneMM = True

                Case "AR_TitoloEpisodio"
                    If LeggiStringOpzioni("ConcatenaTitoloSerieTv", IT.Opzioni).Equals("True") Then
                        NuovoContenuto = IIf(String.IsNullOrEmpty(InfoFiltrate(2).Trim), InfoFiltrate(1), InfoFiltrate(2) & Space(1) & InfoFiltrate(1))
                    Else
                        NuovoContenuto = InfoFiltrate(1)
                    End If

                Case "AR_TitoloSerieTv"
                    If LeggiStringOpzioni("ConcatenaTitoloSerieTv", IT.Opzioni).Equals("False") Then
                        NuovoContenuto = InfoFiltrate(2)
                    End If

                Case "AR_Data"
                    If LeggiStringOpzioni("RemoveInsteadInsert", IT.Opzioni).Equals("False") Then
                        NuovoContenuto = FormattaData(InfoFiltrate(3), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    End If
                    IgnoraConversioneMM = True

                Case "AR_RegexPattern"
                    NuovoContenuto = EstrapolaInfoRegexPattern(IO.Path.GetFileNameWithoutExtension(PathFile), LeggiStringOpzioni("Pattern", IT.Opzioni))

                Case "MI_RisoluzioneFilmato"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.DIMENSIONI)
                    IgnoraConversioneMM = True

                Case "MI_AspectRatio"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.ASPECTRATIO)
                    IgnoraConversioneMM = True

                Case "MI_FrameRate"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.FRAMERATE)
                    IgnoraConversioneMM = True

                Case "MI_Durata"
                    NuovoContenuto = ConvertiMillisecondToTime(Info_MI(FINFO_VIDEO.DURATA), LeggiStringOpzioni("FormatoDurataStile", IT.Opzioni))
                    IgnoraConversioneMM = True

                Case "MI_CodecVideo"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.CODECVIDEO)
                    IgnoraConversioneMM = True

                Case "MI_CodecAudio"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.CODECAUDIO)
                    IgnoraConversioneMM = True

                Case "MI_CodecAudioLingua"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.CODECAUDIOLINGUA)
                    IgnoraConversioneMM = True

                Case "MI_Lingue"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.LINGUE)
                    IgnoraConversioneMM = True

                Case "MI_LingueSottotitoli"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.LINGUESOTTOTITOLI)
                    IgnoraConversioneMM = True

                Case "MI_NumeroCapitoli"
                    NuovoContenuto = Info_MI(FINFO_VIDEO.NUMEROCAPITOLI)
                    IgnoraConversioneMM = True

                Case "TVDB_TitoloEpisodio"
                    NuovoContenuto = Info_TVDB(FINFO_TVDB.TITOLOEPISODIO)
                    If String.IsNullOrEmpty(NuovoContenuto) AndAlso XMLSettings_Read("TVDB_Risultati").Equals("True") Then NuovoContenuto = InfoFiltrate(1)

                Case "TVDB_TitoloSerieTv"
                    NuovoContenuto = Info_TVDB(FINFO_TVDB.TITOLOSERIETV)
                    If String.IsNullOrEmpty(NuovoContenuto) AndAlso XMLSettings_Read("TVDB_Risultati").Equals("True") Then NuovoContenuto = InfoFiltrate(2)

                Case "TVDB_DataPrimaTv"
                    NuovoContenuto = FormattaData(Info_TVDB(FINFO_TVDB.DATAPRIMATV), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), True)
                    IgnoraConversioneMM = True

                Case "TVDB_Creator"
                    NuovoContenuto = Info_TVDB(FINFO_TVDB.CREATOR)

                Case "TVDB_Director"
                    NuovoContenuto = Info_TVDB(FINFO_TVDB.DIRECTOR)

                Case "TVDB_NumeroEpisodi"
                    NuovoContenuto = Info_TVDB(FINFO_TVDB.NUMEROEPISODI)

                Case "TVDB_NumeroStagioni"
                    NuovoContenuto = Info_TVDB(FINFO_TVDB.NUMEROSTAGIONI)

                Case "TVDB_Genere"
                    NuovoContenuto = Info_TVDB(FINFO_TVDB.GENERE)

                Case "TVDB_Network"
                    NuovoContenuto = Info_TVDB(FINFO_TVDB.NETWORK)

                Case "PF_NomeFile"
                    NuovoContenuto = Info_PF(FINFO_PROPRIETA.NOMEFILE)

                Case "PF_DataCreazione"
                    NuovoContenuto = FormattaData(Info_PF(FINFO_PROPRIETA.DATACREAZIONE), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)

                Case "PF_DataUltimaModifica"
                    NuovoContenuto = FormattaData(Info_PF(FINFO_PROPRIETA.DATAULTIMAMODIFICA), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)

                Case "PF_DataUltimoAccesso"
                    NuovoContenuto = FormattaData(Info_PF(FINFO_PROPRIETA.DATAULTIMOACCESSO), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)

                Case "EXT_Opzioni"
                    If Not String.IsNullOrEmpty(LeggiStringOpzioni("EXTSostituisci", IT.Opzioni)) Then
                        Estensione = "." & LeggiStringOpzioni("EXTSostituisci", IT.Opzioni)
                    End If
                    Estensione = ConversioneMaiuscoleMinuscole_EXT(Estensione, LeggiStringOpzioni("EXTMaiuscoloMinisucolo", IT.Opzioni))
                    IgnoraConversioneMM = True

            End Select

            If Not IgnoraConversioneMM Then
                If PrimaFraseOparola Then TestoPrecedeUnSeparatore = True : PrimaFraseOparola = False 'nel caso 3 di capitalizza frase

                If StileMaiuscoleMinuscole = 3 OrElse StileMaiuscoleMinuscole = 4 Then ' caso 3/4
                    If (Coll_Struttura.IndexOf(IT) - 1) >= 0 Then
                        Dim IT_Precedente As ItemStruttura = Coll_Struttura((Coll_Struttura.IndexOf(IT) - 1))

                        If IT_Precedente.TipoDato.Equals("STD_NumerazioneSequenziale") OrElse IT_Precedente.TipoDato.Equals("AR_Numerazione") Then
                            'Capitalizzo quando precede una numerazione
                            TestoPrecedeUnSeparatore = True
                        End If
                        If StileMaiuscoleMinuscole = 4 AndAlso IT_Precedente.TipoDato.Equals("STD_Separatore") Then
                            'Capitalizzo anche dopo ogni separatore
                            TestoPrecedeUnSeparatore = True
                        End If
                    End If
                End If

                NuovoContenuto = ConversioneMaiuscoleMinuscole(NuovoContenuto, StileMaiuscoleMinuscole, TestoPrecedeUnSeparatore)
            End If
            'reset
            TestoPrecedeUnSeparatore = False

            NomeRinominato &= NuovoContenuto
        Next

        'sostituzione termini
        If XMLSettings_Read("CATEGORIA_serietv_config_SostituzioneTermini").Equals("True") Then
            NomeRinominato = SostituzioneTermini(NomeRinominato, Coll_TerminiSostituzioni)
        End If

        'Check filename lenght limit
        CheckFilenameLenghtLimit(NomeRinominato)

        'Controllo eventuali caratteri non ammessi
        NomeRinominato = ValidateChar_Replace(NomeRinominato)

        NomeRinominato &= Estensione

        Return NomeRinominato
    End Function



    Private Function Get_TVDB_Info(ID_SerieTv As String, linguaAbbrev As String, numeroEpisodio As Integer, numeroStagione As Integer, nomeSerieTv As String, ordineEpisodi As Integer) As String()
        Dim ListaLingue As List(Of TvdbLib.Data.TvdbLanguage) = tvDB.Languages
        tvDB.InitCache()

        If ID_SerieTv.Equals("-1") Then
            Try
                'Ricerca automatica del risultato

                If Not TVDB_Dizionario_Trovati.ContainsKey(nomeSerieTv) Then

                    'Per migliorare la ricerca pulisco dal nome eventuali caratteri superflui
                    Dim nomeSerieTvPulita As String = Regex.Replace(nomeSerieTv, "\s{2,}", Space(1)) 'elimino più spazi vuoti consecutivi
                    nomeSerieTvPulita = Regex.Replace(nomeSerieTvPulita, "[^a-zA-Z\s\d]", "") 'elimino caratteri speciali

                    Dim LinguaPrimaria As TvdbLib.Data.TvdbLanguage = ListaLingue.Find(Function(x) x.Abbriviation = XMLSettings_Read("TVDB_LinguaPrimaria"))
                    Dim LinguaSecondaria As TvdbLib.Data.TvdbLanguage = ListaLingue.Find(Function(x) x.Abbriviation = XMLSettings_Read("TVDB_LinguaSecondaria"))

                    '>> ricerca con lingua primaria

                    Dim Risultati As List(Of TvdbLib.Data.TvdbSearchResult) = tvDB.SearchSeries(nomeSerieTvPulita, LinguaPrimaria)
                    Dim ricercaConclusa As Boolean = False

                    For Each series As TvdbLib.Data.TvdbSearchResult In Risultati
                        For Each parola As String In nomeSerieTvPulita.Split(Space(1))
                            If parola.Length <= 3 Then Continue For
                            If series.SeriesName.Contains(parola) Then
                                ID_SerieTv = series.Id
                                linguaAbbrev = LinguaPrimaria.Abbriviation
                                ricercaConclusa = True
                            End If
                        Next
                        If ricercaConclusa Then Exit For
                    Next

                    '>> ricerca con lingua secondaria
                    If Not ricercaConclusa Then
                        Risultati = tvDB.SearchSeries(nomeSerieTvPulita, LinguaSecondaria)

                        For Each series As TvdbLib.Data.TvdbSearchResult In Risultati
                            For Each parola As String In nomeSerieTvPulita.Split(Space(1))
                                If parola.Length <= 3 Then Continue For
                                If series.SeriesName.Contains(parola) Then
                                    ID_SerieTv = series.Id
                                    linguaAbbrev = LinguaSecondaria.Abbriviation
                                    ricercaConclusa = True
                                End If
                            Next
                            If ricercaConclusa Then Exit For
                        Next

                    End If

                    If ID_SerieTv.Equals("-1") Then
                        Return Nothing
                    Else
                        TVDB_Dizionario_Trovati.Add(nomeSerieTv, {ID_SerieTv, linguaAbbrev})
                    End If
                End If
            Catch ex As Exception
                Throw New Exception(Autorinomina.Localization.Resource_Common.AR_Core_Error_TVDB_Error, ex.InnerException)
            End Try
        Else
            If Not TVDB_Dizionario_Trovati.ContainsKey(nomeSerieTv) Then TVDB_Dizionario_Trovati.Add(nomeSerieTv, {ID_SerieTv, linguaAbbrev})
        End If

        'Recupero informazioni
        Dim titoloEpisodio As String = ""
        Dim titoloSerieTv As String = ""
        Dim dataPrimaTv As String = ""

        Try
            Dim LinguaScelta As TvdbLib.Data.TvdbLanguage = ListaLingue.Find(Function(x) x.Abbriviation = linguaAbbrev)
            Dim Serie As TvdbLib.Data.TvdbSeries = tvDB.GetSeries(ID_SerieTv, LinguaScelta, True, False, True, True)

            Dim ListEpisodi As List(Of TvdbLib.Data.TvdbEpisode) = Serie.GetEpisodes(numeroStagione)

            Select Case ordineEpisodi
                Case 1 'ordine dvd (episodi con l'ordine originale del dvd)

                    If ListEpisodi(0).DvdEpisodeNumber = Nothing Then Throw New ARCoreException(Localization.Resource_Common.AR_Core_Error_TVDB_InexistentDVDnumerazion)

                    ListEpisodi = ListEpisodi.OrderBy(Function(a) a.DvdEpisodeNumber,
                                     Comparer(Of Double).Create(Function(key1, key2)
                                                                    Dim Ep1N As Double = 0
                                                                    Dim Ep2N As Double = 0

                                                                    If Not key1 = Nothing Then
                                                                        Ep1N = key1
                                                                        Ep2N = key2
                                                                    End If

                                                                    If Ep1N = Ep2N Then Return 0
                                                                    Return IIf(Ep1N < Ep2N, -1, 1)
                                                                End Function)).ToList

                Case 2, 3
                    '2 ordine dvd combinato (episodi con l'ordine originale del dvd, ma con eventuali puntate unificate. ES: il primo episodio contiene le prime 3 puntate)
                    '3 ordine dvd combinato con titoli (stesso del caso 2, ma qui vengono concatenate nel nome dell'episodio tutti i titoli delle 3 puntate)

                    If ListEpisodi(0).DvdEpisodeNumber = Nothing Then Throw New ARCoreException(Localization.Resource_Common.AR_Core_Error_TVDB_InexistentDVDnumerazion)

                    ListEpisodi = ListEpisodi.OrderBy(Function(a) a.DvdEpisodeNumber,
                                   Comparer(Of Double).Create(Function(key1, key2)
                                                                  Dim Ep1N As Double = 0
                                                                  Dim Ep2N As Double = 0

                                                                  If Not key1 = Nothing Then
                                                                      Ep1N = key1
                                                                      Ep2N = key2
                                                                  End If

                                                                  If Ep1N = Ep2N Then Return 0
                                                                  Return IIf(Ep1N < Ep2N, -1, 1)
                                                              End Function)).ToList

                    Dim tempLE As New List(Of TvdbLib.Data.TvdbEpisode)
                    Dim tempNepisodio As Integer = 0
                    Dim tempNCombinati As Integer = 0
                    Dim tempTitoliEpisodi As String = ""

                    For x As Integer = 0 To ListEpisodi.Count - 1
                        Dim episode As TvdbLib.Data.TvdbEpisode = ListEpisodi(x)
                        Dim Nepisodio As Integer = Integer.Parse(Regex.Split(episode.DvdEpisodeNumber.ToString, "\.")(0))

                        If Not tempNepisodio = Nepisodio Then
                            If tempLE.Count > 1 AndAlso tempNCombinati > 1 Then
                                If ordineEpisodi = 2 Then
                                    tempLE(tempLE.Count - 1).EpisodeName = tempLE(tempLE.Count - 1).EpisodeName & " (" & tempNCombinati & Space(1) & Localization.Resource_Common.AR_Core_Episodes & ")"
                                Else
                                    tempLE(tempLE.Count - 1).EpisodeName = tempTitoliEpisodi
                                End If

                                tempLE.Add(episode)

                                tempTitoliEpisodi = episode.EpisodeName
                                tempNCombinati = 1
                            Else

                                tempTitoliEpisodi &= ", " & episode.EpisodeName
                                tempNCombinati += 1

                                If tempLE.Count - 1 = x AndAlso tempNCombinati > 1 Then
                                    If ordineEpisodi = 2 Then
                                        tempLE(tempLE.Count - 1).EpisodeName = tempLE(tempLE.Count - 1).EpisodeName & " (" & tempNCombinati & Space(1) & Localization.Resource_Common.AR_Core_Episodes & ")"
                                    Else
                                        tempLE(tempLE.Count - 1).EpisodeName = tempTitoliEpisodi
                                    End If
                                End If
                            End If
                        End If

                        tempNepisodio = Nepisodio
                    Next

                    ListEpisodi = tempLE

                Case Else 'ordine di default di the tvdb

                    ListEpisodi = ListEpisodi.OrderBy(Function(a) a.EpisodeNumber,
                                  Comparer(Of Integer).Create(Function(key1, key2)
                                                                  If key1 = key2 Then Return 0
                                                                  Return IIf(key1 < key2, -1, 1)
                                                              End Function)).ToList
            End Select

            Dim Val(FINFO_TVDB.Count - 1) As String

            If Not ListEpisodi Is Nothing AndAlso ListEpisodi.Count >= numeroEpisodio Then
                Dim episode As TvdbLib.Data.TvdbEpisode = ListEpisodi(numeroEpisodio - 1)

                Val(FINFO_TVDB.TITOLOEPISODIO) = episode.EpisodeName
                Val(FINFO_TVDB.TITOLOSERIETV) = Serie.SeriesName
                Val(FINFO_TVDB.DATAPRIMATV) = episode.FirstAired

                Dim CreatorString As String = ""
                For Each name In episode.Writer
                    CreatorString &= name & ", "
                Next
                If CreatorString.Length > 1 Then Val(FINFO_TVDB.CREATOR) = CreatorString.Remove(CreatorString.Length - 2, 2)

                Dim DirectorString As String = ""
                For Each name In episode.Directors
                    DirectorString &= name & ", "
                Next
                If DirectorString.Length > 1 Then Val(FINFO_TVDB.DIRECTOR) = DirectorString.Remove(DirectorString.Length - 2, 2)

                Val(FINFO_TVDB.NUMEROEPISODI) = ListEpisodi.Count
                Val(FINFO_TVDB.NUMEROSTAGIONI) = Serie.NumSeasons

                Dim GenereString As String = ""
                For Each name In Serie.Genre
                    GenereString &= name & ", "
                Next
                If GenereString.Length > 1 Then Val(FINFO_TVDB.GENERE) = GenereString.Remove(GenereString.Length - 2, 2)

                Val(FINFO_TVDB.NETWORK) = Serie.Network
            End If

            Return Val

        Catch ex As Exception
            Throw New Exception(Autorinomina.Localization.Resource_Common.AR_Core_Error_TVDB_Error, ex.InnerException)
        Finally
            tvDB.CloseCache()
        End Try
    End Function



    Private Function Filtro_SerieTv(NomeFile As String, CarattereSplit As String) As String()
        Dim Risultato_TestoContenuto As String = "" 'titolo episodio definitivo
        Dim Risultato_Numerazione As String = ""
        Dim Risultato_TitoloSerie As String = ""
        Dim Risultato_Data As String = ""

        Dim Temp_TitoloSerie As String = ""
        Dim Temp_TitoloSerie_Alternativo As String = ""
        Dim NomeFile_Filtrato As String = ""

        If CarattereSplit = "" Then Return {"", "", "", "", "False"}

        NomeFile_Filtrato = RimozioneTerminiBlackList(IO.Path.GetFileNameWithoutExtension(NomeFile), Coll_TerminiBlackList)

        'Cerco la numerazione
        Dim Stagione As String = ""
        Dim Episodio As String = "-1"
        Dim EpisodioConcat As String = "0"

        If PatternRegexNumerazioni.Count > 0 Then
            For x As Integer = 0 To PatternRegexNumerazioni.Count - 1

                If Regex.IsMatch(NomeFile_Filtrato, PatternRegexNumerazioni(x), RegexOptions.IgnoreCase) Then
                    If XMLSettings_Read("CATEGORIA_serietv_config_RecognizeLinkedEpisode").Equals("False") Then
                        If PatternRegexNumerazioni(x).ToString.Contains("episode2") Then Continue For
                    End If

                    Dim m As Match = Regex.Match(NomeFile_Filtrato, PatternRegexNumerazioni(x), RegexOptions.IgnoreCase)
                    If m.Groups("season").Value() <> "" Then Stagione = m.Groups("season").Value()
                    If m.Groups("episode").Value() <> "" Then Episodio = m.Groups("episode").Value()
                    If m.Groups("episode2").Value() <> "" Then EpisodioConcat = m.Groups("episode2").Value()

                    'eseguo uno split nel caso sia presente il titolo della serietv PRIMA della numerazione
                    'regex flag case insensitive -> fix per 0X00 al posto di 0x00
                    Dim tmp() As String = Regex.Split(NomeFile_Filtrato, Regex.Escape(m.Value), RegexOptions.IgnoreCase)

                    Temp_TitoloSerie = tmp(0)
                    If tmp.Length > 0 Then NomeFile_Filtrato = tmp(1)

                    Exit For
                End If
            Next
        End If

        'Definisco il nuovo stile di numerazione temporaneo, che sarà successivamente cambiato in base alla scelta dell'utente in un secondo momento
        If Not Episodio.Equals("-1") Then
            If Stagione = "" Then
                Risultato_Numerazione = Episodio
            Else
                Risultato_Numerazione = Stagione & "x" & Episodio &
                    IIf(Not EpisodioConcat.Equals("0"), "-" & EpisodioConcat, "")
            End If
        End If

        'funzione per determinare il carattere da usare come split
        If CarattereSplit.Equals("_.") Then
            NomeFile_Filtrato = NomeFile_Filtrato.Replace("_", ".")
            CarattereSplit = "#"
        End If

        If CarattereSplit.Equals("#") Then
            Dim ContenutoRipulito As String = NomeFile_Filtrato

            'elimino eventuali (...) oppure [...] e ... per un riconoscimento piu preciso del carattere da usare nello split
            ContenutoRipulito = Regex.Replace(ContenutoRipulito, "\(.*?\)", "")
            ContenutoRipulito = Regex.Replace(ContenutoRipulito, "\[.*?\]", "")
            ContenutoRipulito = Regex.Replace(ContenutoRipulito, "\.{3}", "")

            'determino la quantità  di caratteri speciali presenti
            Dim QtaCharSpec() As Integer = {0, 0, 0}
            QtaCharSpec(0) = Regex.Split(ContenutoRipulito, "\.").Length
            QtaCharSpec(1) = Regex.Split(ContenutoRipulito, "_").Length
            QtaCharSpec(2) = Regex.Split(ContenutoRipulito, "\-").Length

            Dim IndexCharSpec As Integer = GetIndexOfBiggestArrayItem(QtaCharSpec)
            If QtaCharSpec(IndexCharSpec) > 2 Then
                Select Case IndexCharSpec
                    Case 0
                        CarattereSplit = "."
                        'fix per il tipo:    ReGenesis-2x10-IL.Selvaggio.e.l'innocente-[D.Tv_ita]-by.L@MI@.avi
                        If QtaCharSpec(0) >= QtaCharSpec(2) Then NomeFile_Filtrato = NomeFile_Filtrato.Replace("-", ".")

                    Case 1
                        CarattereSplit = "_"
                        If (QtaCharSpec(1) >= QtaCharSpec(0)) Then NomeFile_Filtrato = NomeFile_Filtrato.Replace(".", "_")
                        If (QtaCharSpec(1) >= QtaCharSpec(2)) Then NomeFile_Filtrato = NomeFile_Filtrato.Replace("-", "_")

                    Case 2
                        CarattereSplit = "-"
                        If (QtaCharSpec(2) >= QtaCharSpec(0)) Then NomeFile_Filtrato = NomeFile_Filtrato.Replace(".", "-")
                        If (QtaCharSpec(2) >= QtaCharSpec(1)) Then NomeFile_Filtrato = NomeFile_Filtrato.Replace("_", "-")
                End Select


            Else
                If QtaCharSpec(0) = 2 Then
                    'FIX per questo tipo: "Supernatural.4x02.Sei lì, Dio, sono io, Dean Winchester.FFT.sat.ITA.avi"
                    'contengono 2 punti che causano errore

                    CarattereSplit = "."
                    NomeFile_Filtrato = NomeFile_Filtrato.Replace(" ", ".")
                Else
                    CarattereSplit = Space(1)
                End If
            End If

            If QtaCharSpec(1) > 2 AndAlso Not CarattereSplit = "_" Then NomeFile_Filtrato = NomeFile_Filtrato.Replace("_", ".")
        End If

        'Rimozione ()[] anche con contenuti inclusi dal titolo dell'episodio
        If XMLSettings_Read("CATEGORIA_serietv_config_EliminaParentesiTitoloEpisodio").Equals("True") Then
            NomeFile_Filtrato = RimozioneParentesi(NomeFile_Filtrato, True).Trim
        End If

        'esecuzione dello split
        Dim ContenutoNomeFileSplittato As New ArrayList
        Dim StrSplitted() As String = Regex.Split(NomeFile_Filtrato, Regex.Escape(CarattereSplit))

        For x As Integer = 0 To StrSplitted.Length - 1
            If Not StrSplitted(x).Trim.Length = 0 Then ContenutoNomeFileSplittato.Add(StrSplitted(x))
        Next

        'eseguo un split in base al carattereSplit del titolo episodio 
        For x As Integer = 0 To ContenutoNomeFileSplittato.Count - 1
            If (ContenutoNomeFileSplittato.Count - 1) >= 1 AndAlso Temp_TitoloSerie_Alternativo.Length = 0 AndAlso Temp_TitoloSerie.Length = 0 Then
                Temp_TitoloSerie_Alternativo = ContenutoNomeFileSplittato(x).ToString.Trim

            Else

                Select Case ContenutoNomeFileSplittato(x).ToString.Trim
                    Case "-", "_", ".", ""
                    Case Else

                        Dim Result As String = ContenutoNomeFileSplittato(x).ToString
                        If Regex.IsMatch(ContenutoNomeFileSplittato(x).ToString, "\.{3}", RegexOptions.IgnoreCase) Then
                            Result = Result.Replace("_", " ").Trim
                        Else
                            Result = Result.Replace("_", " ")
                            If CarattereSplit.Equals(".") Then Result = Result.Replace(".", " ")
                            Result = Result.Trim
                        End If

                        If Result.Length > 0 Then
                            Risultato_TestoContenuto &= IIf(Risultato_TestoContenuto.Length > 0, Space(1), "") & Result
                        End If
                End Select
            End If
        Next

        'Elimino questi eventuali caratteri rimanenti .-_ dall'inizio e dalla fine
        Dim RegexPatterns As String() = {"\s*\-+\s*$", "\s*\.+\s*$", "\s*_+\s*$", "^\s*\-+\s*", "^\s*\.+\s*", "^\s*_+\s*"}
        For Each Pattern As String In RegexPatterns
            If Regex.IsMatch(Risultato_TestoContenuto, Pattern) = True Then
                Risultato_TestoContenuto = Regex.Replace(Risultato_TestoContenuto, Pattern, "")
            End If
        Next

        'Elimino eventuali parentesi vuote
        Risultato_TestoContenuto = RimozioneParentesi(Risultato_TestoContenuto, False)


        '-----------------------------------------------------------------------> RICONOSCIMENTO TITOLO SERIETV
        Dim Result_TitoloSerie As String = Temp_TitoloSerie
        Dim ContenutoRipulitoBis As String = Temp_TitoloSerie

        'elimino eventuali (...) oppure [...] e ... per un riconoscimento piu preciso del carattere da usare nello split
        ContenutoRipulitoBis = Regex.Replace(ContenutoRipulitoBis, "\(.*?\)", "")
        ContenutoRipulitoBis = Regex.Replace(ContenutoRipulitoBis, "\[.*?\]", "")
        ContenutoRipulitoBis = Regex.Replace(ContenutoRipulitoBis, "\.{3}", "")


        'determino la quantità  di caratteri speciali presenti
        Dim QtaCharSpecBis() As Integer = {0, 0, 0}
        QtaCharSpecBis(0) = Regex.Split(ContenutoRipulitoBis, "\.").Length
        QtaCharSpecBis(1) = Regex.Split(ContenutoRipulitoBis, "_").Length
        QtaCharSpecBis(2) = Regex.Split(ContenutoRipulitoBis, "\-").Length

        Dim IndexCharSpecBIS As Integer = GetIndexOfBiggestArrayItem(QtaCharSpecBis)
        If QtaCharSpecBis(IndexCharSpecBIS) > 2 Then
            Select Case IndexCharSpecBIS
                Case 0
                    CarattereSplit = "."
                Case 1
                    CarattereSplit = "_"
                Case 2
                    CarattereSplit = "-"
            End Select
        Else
            CarattereSplit = Space(1)
        End If

        Result_TitoloSerie = Result_TitoloSerie.Replace(CarattereSplit, " ")

        'Elimino eventuali parentesi vuote o con questi caratteri - . o spazi
        Result_TitoloSerie = Regex.Replace(Result_TitoloSerie, "\[(\-|\.|\s)+\]", "")
        Result_TitoloSerie = Regex.Replace(Result_TitoloSerie, "\((\-|\.|\s)+\)", "")

        'Elimino - e . dall'inizio e dalla fine
        Result_TitoloSerie = Regex.Replace(Result_TitoloSerie, "^\s*\-", "")
        Result_TitoloSerie = Regex.Replace(Result_TitoloSerie, "\-\s*$", "")

        Result_TitoloSerie = Regex.Replace(Result_TitoloSerie, "^\s*\.", "")
        Result_TitoloSerie = Regex.Replace(Result_TitoloSerie, "\.\s*$", "")

        'Fix per quei casi in cui non c'è il titolo della serie nel nome del file
        If Temp_TitoloSerie_Alternativo.Trim = "-" Then Temp_TitoloSerie_Alternativo = ""
        If Temp_TitoloSerie_Alternativo.Trim = "_" Then Temp_TitoloSerie_Alternativo = ""
        If Temp_TitoloSerie_Alternativo.Trim = "." Then Temp_TitoloSerie_Alternativo = ""

        If Temp_TitoloSerie_Alternativo.Trim.Length > 0 Then
            Risultato_TitoloSerie = Temp_TitoloSerie_Alternativo.Trim
        Else
            Risultato_TitoloSerie = Result_TitoloSerie.Trim
        End If


        'Cerco la data
        Risultato_Data = EstrapolaData(Risultato_TestoContenuto)

        'Rimozione ()[] anche con contenuti inclusi dal titolo della serie 
        If XMLSettings_Read("CATEGORIA_serietv_config_EliminaParentesiTitoloSerieTv").Equals("True") Then
            Risultato_TitoloSerie = RimozioneParentesi(Risultato_TitoloSerie, True).Trim
        End If

        'Eventuali Trim degli spazi all'interno delle parentesi "(  esempio" --> "(esempio"
        Risultato_TitoloSerie = TrimInternoParentesi(Risultato_TitoloSerie)

        'Converto i - ii - iii in> I - II - III | Solo se presente all'interno di () es: (parte ii)
        If XMLSettings_Read("CATEGORIA_serietv_config_ConvertiNumRomani").Equals("True") Then
            Dim ElencoTXTRicerca() As String = {"i", "ii", "iii"}
            Dim ElencoTXTSostituisci() As String = {"1", "2", "3"}

            For x As Integer = 0 To ElencoTXTRicerca.Length - 1
                'Cerca: (parte ii)
                Dim rg As New Regex("\(.*?\W" & ElencoTXTRicerca(x) & "\)", RegexOptions.IgnoreCase)
                If rg.IsMatch(Risultato_TestoContenuto) Then
                    Risultato_TestoContenuto = Regex.Replace(Risultato_TestoContenuto, "\W" & ElencoTXTRicerca(x) & "\)", " " & ElencoTXTSostituisci(x) & ")", RegexOptions.IgnoreCase)
                End If

                'Cerca: (ii parte)
                Dim rgF As New Regex("\(" & ElencoTXTRicerca(x) & "\W.*?\)", RegexOptions.IgnoreCase)
                If rgF.IsMatch(Risultato_TestoContenuto) Then
                    Risultato_TestoContenuto = Regex.Replace(Risultato_TestoContenuto, "\(" & ElencoTXTRicerca(x) & "\W", "(" & ElencoTXTSostituisci(x) & " ", RegexOptions.IgnoreCase)
                End If
            Next
        End If


        Dim ValidaRisultato As Boolean = ((Risultato_TestoContenuto.Length > 1) OrElse (Risultato_TitoloSerie.Length > 1)) AndAlso (Risultato_Numerazione.Length > 0)

        Return {Risultato_Numerazione, Risultato_TestoContenuto, Risultato_TitoloSerie, Risultato_Data, ValidaRisultato.ToString}
    End Function

End Class
