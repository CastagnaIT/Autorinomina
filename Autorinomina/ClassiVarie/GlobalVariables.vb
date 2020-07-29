Module GlobalVariables

    Public DataPath As String

    Public Coll_Files As New CollItemsFilesData
    Public Coll_Files_OnPropertyChanged_Enabled As Boolean = False

    Public Coll_Struttura As New CollItemsStruttura

    Public Coll_BlackList_New As New List(Of String)

    Public Const TVDB_APIKEY As String = "188970691D51D336"


    Public FallbackDefaultDataStruttura As New Dictionary(Of String, String)
    Public FallbackDefaultXMLSettings As New Dictionary(Of String, String)

    Public Sub CaricaValoriFallback()
        'FALLBACK VALORI STRUTTURA
        FallbackDefaultDataStruttura.Add("Testo", "")
        FallbackDefaultDataStruttura.Add("SeparatoreCarattere", "-")
        FallbackDefaultDataStruttura.Add("SeparatoreSpaziatura", "True")
        FallbackDefaultDataStruttura.Add("NumerazioneIniziaDa", "1")
        FallbackDefaultDataStruttura.Add("NumerazioneStagione", "1")
        FallbackDefaultDataStruttura.Add("NumerazioneStile", "0") '0 = stile numerico normale
        FallbackDefaultDataStruttura.Add("NumerazioneAutoPadding", "True")
        FallbackDefaultDataStruttura.Add("NumerazionePadding", "2")
        FallbackDefaultDataStruttura.Add("NumerazionePrefisso", "")
        FallbackDefaultDataStruttura.Add("NumerazioneSuffisso", "")
        FallbackDefaultDataStruttura.Add("FormatoDurataStile", Autorinomina.Localization.Resource_BalloonToolTip.Grid_FormatLenght_TimeFormat_Ex4) 'formato tempo durata audio e video
        FallbackDefaultDataStruttura.Add("FormatoDataStile", Autorinomina.Localization.Resource_BalloonToolTip.Grid_FormatDateTime_Ex2)
        FallbackDefaultDataStruttura.Add("EstensioniSostituisci", "")
        FallbackDefaultDataStruttura.Add("EstensioniMinuscole", "False")
        FallbackDefaultDataStruttura.Add("OrdineEpisodi", "0")
        FallbackDefaultDataStruttura.Add("ConcatenaTitoloSerieTv", "False")
        FallbackDefaultDataStruttura.Add("TVDBOrdineEpisodi", "0") 'ordine episodi TheTVDB 0=default
        FallbackDefaultDataStruttura.Add("EXTMaiuscoloMinisucolo", "0")
        FallbackDefaultDataStruttura.Add("EXTSostituisci", "")
        FallbackDefaultDataStruttura.Add("UnitMeasure", "CM")
        FallbackDefaultDataStruttura.Add("UnitMeasureRounded", "False")
        FallbackDefaultDataStruttura.Add("SizeStyle", Autorinomina.Localization.Resource_BalloonToolTip.Grid_StyleSize_Format_Ex1)
        FallbackDefaultDataStruttura.Add("Pattern", "")
        FallbackDefaultDataStruttura.Add("RemoveInsteadInsert", "False")

        'FALLBACK VALORI SETTINGS XML
        '(PER CHIAVI INSERITE SU NUOVE VERSIONI DI AR, MANTIENE COMPATIBILITA' SUGLI UPGRADE DEGLI UTENTI PERCHE' IL SETTINGS.XML NON VIENE AGGIORNATO)
        FallbackDefaultXMLSettings.Add("TVDB_LinguaFallBack", "True")
        FallbackDefaultXMLSettings.Add("TVDB_AssociazioneStagioni", "0")
        FallbackDefaultXMLSettings.Add("PreviewHintsNewBlackListWords", "True")
        FallbackDefaultXMLSettings.Add("PreviewHintsNewBlackListWords_Sensibility", "90")

    End Sub


    Public Enum ENUM_PREVIEW_RESULT
        ERRORI 'errori vari gestiti
        NOANTEPRIMA 'da impostazioni
        CONFLITTI 'nomi di file già presenti (localmente o fisicamente nelle cartelle)
        Count
    End Enum

    Public Enum ENUM_RENAME_RESULT
        ERRORI 'file non più esistenti / nuovo nome non è presente
        NONRINOMINATI_RIPRISTINATI 'non rinominato/ripristinato da impostazioni
        NONRINOMINATI_RIPRISTINATI_BYUSER 'non rinominato/ripristinato scelta utente
        CONFLITTI 'nomi di file già presenti (fisicamente nelle cartelle)
        Count
    End Enum
End Module
