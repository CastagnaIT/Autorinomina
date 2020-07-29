Public Class AR_Core_Audio
    Inherits AR_Core

    Dim Coll_TerminiBlackList As New CollItemsTermini
    Dim Coll_TerminiSostituzioni As New CollItemsTermini

    Public Sub New(_OwnerWindow As Window)
        MyBase.New(_OwnerWindow)

        Coll_TerminiBlackList.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & "BlackList_CATEGORIA_audio.xml", "BlackList"))
        Coll_TerminiSostituzioni.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & "Sostituzioni_CATEGORIA_audio.xml", "Sostituzioni"))
    End Sub

    Public Overrides Function RunCore(ByVal PathFile As String, ByVal NumeroSequenziale As Integer, ByVal LivePreview As Boolean, ByVal PaddingLenght As Integer, ByVal Optional overrideTipoFiltro As String = Nothing) As String
        'InfoFiltrate index informazioni: 0 = numerazione | 1 = nome file filtrato | 2 = data | 3 = anno
        Dim InfoFiltrate() As String = Nothing
        Dim Info_MI() As String = Nothing
        Dim Info_PF() As String = Nothing
        Dim Estensione As String = IO.Path.GetExtension(PathFile)

        If InfoRichiesteStruttura.Contains("AR_") Then InfoFiltrate = Filtro_Generico(PathFile, "audio", Coll_TerminiBlackList, Coll_TerminiSostituzioni)
        If InfoRichiesteStruttura.Contains("MI_") Then Info_MI = GetAudioInfo(PathFile)
        If InfoRichiesteStruttura.Contains("PF_") Then Info_PF = GetPropertyInfo(PathFile)

        Dim StileMaiuscoleMinuscole As Integer = Integer.Parse(XMLSettings_Read("CATEGORIA_audio_config_StileMaiuscoleMinuscole"))
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
                    Select Case LeggiStringOpzioni("NumerazioneStile", IT.Opzioni)
                        Case "5", "6", "7", "8"
                            NewBaseNum = LeggiStringOpzioni("NumerazioneStagione", IT.Opzioni) & "x" & (Integer.Parse(LeggiStringOpzioni("NumerazioneIniziaDa", IT.Opzioni)) + NumeroSequenziale)
                        Case Else
                            NewBaseNum = (Integer.Parse(LeggiStringOpzioni("NumerazioneIniziaDa", IT.Opzioni)) + NumeroSequenziale).ToString
                    End Select

                    NuovoContenuto = CreaNuovaNumerazione(NewBaseNum, LeggiStringOpzioni("NumerazioneStile", IT.Opzioni), _PaddingLenght)
                    NuovoContenuto = LeggiStringOpzioni("NumerazionePrefisso", IT.Opzioni) & NuovoContenuto & LeggiStringOpzioni("NumerazioneSuffisso", IT.Opzioni)
                    IgnoraConversioneMM = True

                Case "AR_Numerazione", "MI_AudioTagNumero"
                    If IT.TipoDato.Equals("AR_Numerazione") Then
                        'Filtra numerazione dal file
                        NuovoContenuto = InfoFiltrate(0)

                    Else
                        'Utilizza numerazione dal Tag
                        NuovoContenuto = Info_MI(FINFO_AUDIO.TAG_NUMERO)

                        If String.IsNullOrEmpty(NuovoContenuto) Then
                            Select Case Integer.Parse(XMLSettings_Read("CATEGORIA_audio_config_FallbackNumerazione"))
                                Case 0 'Utilizza filtro
                                    If InfoFiltrate Is Nothing Then InfoFiltrate = Filtro_Generico(PathFile, "audio", Coll_TerminiBlackList, Coll_TerminiSostituzioni)
                                    NuovoContenuto = InfoFiltrate(0)

                                Case 1 'Utilizza numero sequenziale
                                    NuovoContenuto = NumeroSequenziale

                                Case 3 'Non rinominare il file
                                    Return ""
                            End Select
                        End If
                    End If

                    Dim _PaddingLenght As Integer = IIf(PaddingLenght <= 1, 2, PaddingLenght) 'in questo caso la numerazione è totalmente automatica il padding deve iniziare sempre da due zeri 01 02 03..
                    NuovoContenuto = CreaNuovaNumerazione(NuovoContenuto, LeggiStringOpzioni("NumerazioneStile", IT.Opzioni), _PaddingLenght)
                    IgnoraConversioneMM = True

                Case "AR_NomeDelFile"
                    NuovoContenuto = InfoFiltrate(1)

                Case "AR_Data"
                    If LeggiStringOpzioni("RemoveInsteadInsert", IT.Opzioni).Equals("False") Then
                        NuovoContenuto = FormattaData(InfoFiltrate(2), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    End If
                    IgnoraConversioneMM = True

                Case "AR_Anno"
                    If LeggiStringOpzioni("RemoveInsteadInsert", IT.Opzioni).Equals("False") Then NuovoContenuto = InfoFiltrate(3)
                    IgnoraConversioneMM = True

                Case "AR_RegexPattern"
                    NuovoContenuto = EstrapolaInfoRegexPattern(IO.Path.GetFileNameWithoutExtension(PathFile), LeggiStringOpzioni("Pattern", IT.Opzioni))

                Case "MI_AudioDurata"
                    NuovoContenuto = ConvertiMillisecondToTime(Info_MI(FINFO_AUDIO.DURATA), LeggiStringOpzioni("FormatoDurataStile", IT.Opzioni))
                    IgnoraConversioneMM = True

                Case "MI_AudioTagArtista"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.TAG_ARTISTA)

                Case "MI_AudioTagAlbum"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.TAG_ALBUM)

                Case "MI_AudioTagTitolo"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.TAG_TITOLO)

                Case "MI_AudioTagCompositore"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.TAG_COMPOSITORE)

                Case "MI_AudioTagPubliscer"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.TAG_PUBLISCER)

                Case "MI_AudioTagGenere"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.TAG_GENERE)

                Case "MI_AudioTagDataRilascio"
                    NuovoContenuto = FormattaData(Info_MI(FINFO_AUDIO.TAG_DATA_RILASCIO), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    IgnoraConversioneMM = True

                Case "MI_AudioTagDataRegistrazione"
                    NuovoContenuto = FormattaData(Info_MI(FINFO_AUDIO.TAG_DATA_REGISTRAZIONE), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    IgnoraConversioneMM = True

                Case "MI_AudioFrequenza"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.FREQUENZA)
                    IgnoraConversioneMM = True

                Case "MI_AudioBitRate"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.BITRATE)
                    IgnoraConversioneMM = True

                Case "MI_AudioModalitaBitRate"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.MODALITA_BITRATE)
                    IgnoraConversioneMM = True

                Case "MI_AudioNumeroCanali"
                    NuovoContenuto = Info_MI(FINFO_AUDIO.NUMERO_CANALI)
                    IgnoraConversioneMM = True

                Case "PF_NomeFile"
                    NuovoContenuto = Info_PF(FINFO_PROPRIETA.NOMEFILE)

                Case "PF_DataCreazione"
                    NuovoContenuto = FormattaData(Info_PF(FINFO_PROPRIETA.DATACREAZIONE), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    IgnoraConversioneMM = True

                Case "PF_DataUltimaModifica"
                    NuovoContenuto = FormattaData(Info_PF(FINFO_PROPRIETA.DATAULTIMAMODIFICA), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    IgnoraConversioneMM = True

                Case "PF_DataUltimoAccesso"
                    NuovoContenuto = FormattaData(Info_PF(FINFO_PROPRIETA.DATAULTIMOACCESSO), LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    IgnoraConversioneMM = True

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
        If XMLSettings_Read("CATEGORIA_audio_config_SostituzioneTermini").Equals("True") Then
            NomeRinominato = SostituzioneTermini(NomeRinominato, Coll_TerminiSostituzioni)
        End If

        'Check filename lenght limit
        CheckFilenameLenghtLimit(NomeRinominato)

        'Controllo eventuali caratteri non ammessi
        NomeRinominato = ValidateChar_Replace(NomeRinominato)

        NomeRinominato &= Estensione

        Return NomeRinominato
    End Function

End Class
