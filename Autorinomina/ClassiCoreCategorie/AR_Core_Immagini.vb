Imports System.Text.RegularExpressions

Public Class AR_Core_Immagini
    Inherits AR_Core

    Dim Coll_TerminiBlackList As New CollItemsTermini
    Dim Coll_TerminiSostituzioni As New CollItemsTermini
    Dim ExtensionMetadataAllowed() As String

    Public Sub New(_OwnerWindow As Window)
        MyBase.New(_OwnerWindow)

        Coll_TerminiBlackList.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & "BlackList_CATEGORIA_immagini.xml", "BlackList"))
        Coll_TerminiSostituzioni.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & "Sostituzioni_CATEGORIA_immagini.xml", "Sostituzioni"))

        ExtensionMetadataAllowed = XMLSettings_Read("Extensions_images_metadata").Split(",")
    End Sub


    Public Overrides Function RunCore(ByVal PathFile As String, ByVal NumeroSequenziale As Integer, ByVal LivePreview As Boolean, ByVal PaddingLenght As Integer, ByVal Optional overrideTipoFiltro As String = Nothing) As String
        'InfoFiltrate index informazioni: 0 = numerazione | 1 = nome file filtrato | 2 = data | 3 = anno
        Dim InfoFiltrate() As String = Nothing
        Dim Info_PF() As String = Nothing
        Dim Info_Exif_Main As FreeImageAPI.Metadata.MDM_EXIF_MAIN = Nothing
        Dim Info_Exif As FreeImageAPI.Metadata.MDM_EXIF_EXIF = Nothing
        Dim Info_IPTC As FreeImageAPI.Metadata.MDM_IPTC = Nothing
        Dim Estensione As String = IO.Path.GetExtension(PathFile)

        Dim Img_dib As FreeImageAPI.FIBITMAP = Nothing
        If InfoRichiesteStruttura.Contains("II_") OrElse InfoRichiesteStruttura.Contains("EXIF_") OrElse InfoRichiesteStruttura.Contains("IPTC_") Then
            Img_dib = Get_Image_dib(PathFile)

            If ExtensionMetadataAllowed.Contains(Estensione.ToLower.Replace(".", "")) Then

                Dim MDM As Object() = Get_Image_MDMExif(Img_dib)
                Info_Exif_Main = MDM(0)
                Info_Exif = MDM(1)

                If InfoRichiesteStruttura.Contains("IPTC_") Then Info_IPTC = Get_Image_IPTC(Img_dib)
            End If
        End If

        If InfoRichiesteStruttura.Contains("AR_") Then InfoFiltrate = Filtro_Generico(PathFile, "immagini", Coll_TerminiBlackList, Coll_TerminiSostituzioni)
        If InfoRichiesteStruttura.Contains("PF_") Then Info_PF = GetPropertyInfo(PathFile)

        Dim StileMaiuscoleMinuscole As Integer = Integer.Parse(XMLSettings_Read("CATEGORIA_immagini_config_StileMaiuscoleMinuscole"))
        Dim TestoPrecedeUnSeparatore As Boolean = False 'Serve per conversione maiuscole/minuscole modalità 4 (Prima lettera maiuscola dopo ogni separatore resto minuscolo) 
        Dim PrimaFraseOparola As Boolean = True 'Serve per capire qual'è la parola iniziale da fare la capitalizzazione
        Dim NomeRinominato As String = ""

        For Each IT As ItemStruttura In Coll_Struttura
            Dim NuovoContenuto As String = ""
            Dim IgnoraConversioneMM As Boolean = False 'Ignora conversione maiuscole/minuscole

            If IT.TipoDato.Contains("EXIF") Then
                If IT.TipoDato.Contains("MAIN") Then
                    If Info_Exif_Main IsNot Nothing Then NuovoContenuto = Get_Exif_Info(Info_Exif_Main, IT.TipoDato)
                Else
                    If Info_Exif IsNot Nothing Then NuovoContenuto = Get_Exif_Info(Info_Exif, IT.TipoDato)
                End If

                If IT.TipoDato.Contains("Date") AndAlso Not String.IsNullOrEmpty(NuovoContenuto) Then
                    NuovoContenuto = FormattaData(NuovoContenuto, LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    IgnoraConversioneMM = True
                End If

            End If

            If IT.TipoDato.Contains("IPTC") Then
                If Info_IPTC IsNot Nothing Then NuovoContenuto = Get_Exif_Info(Info_IPTC, IT.TipoDato)

                If IT.TipoDato.Contains("Date") AndAlso Not String.IsNullOrEmpty(NuovoContenuto) Then
                    NuovoContenuto = FormattaData(NuovoContenuto, LeggiStringOpzioni("FormatoDataStile", IT.Opzioni), False)
                    IgnoraConversioneMM = True
                End If
            End If

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

                Case "AR_Numerazione"
                    'Filtra numerazione dal file
                    NuovoContenuto = InfoFiltrate(0)

                    If String.IsNullOrEmpty(NuovoContenuto) Then
                        Select Case Integer.Parse(XMLSettings_Read("CATEGORIA_immagini_config_FallbackNumerazione"))
                            Case 0 'Utilizza numero sequenziale
                                NuovoContenuto = NumeroSequenziale

                            Case 2 'Non rinominare il file
                                Return ""
                        End Select
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

                Case "II_Dimensione"
                    'Calcolando il DPI manualmente si ottiene un valore molto più preciso, ottenendo dimensioni esatte dell'immagine
                    'Invece il metodo >> FreeImageAPI.FreeImage.GetResolutionX(Img_dib) << restituisce un valore DPI UInteger pertanto il calcolo delle dimensioni dell'immagine risulta impreciso
                    Dim DPIX As Double = 0.0254 * FreeImageAPI.FreeImage.GetDotsPerMeterX(Img_dib)
                    Dim DPIY As Double = 0.0254 * FreeImageAPI.FreeImage.GetDotsPerMeterY(Img_dib)

                    If DPIX = 0 Then DPIX = 72 '< nel caso di dpi non impostati, imposto i dpi standard
                    If DPIY = 0 Then DPIY = 72

                    Dim Larghezza As Double = (FreeImageAPI.FreeImage.GetWidth(Img_dib) * 2.54) / DPIX 'Calcolo la base in CM
                    Dim Altezza As Double = (FreeImageAPI.FreeImage.GetHeight(Img_dib) * 2.54) / DPIY 'Calcolo la base in CM

                    Select Case LeggiStringOpzioni("UnitMeasure", IT.Opzioni)
                        Case "PX"
                            Larghezza = FreeImageAPI.FreeImage.GetWidth(Img_dib)
                            Altezza = FreeImageAPI.FreeImage.GetHeight(Img_dib)
                        Case "MM"
                            Larghezza = Larghezza * 10
                            Altezza = Altezza * 10
                        Case "INCH"
                            Larghezza = Larghezza * 0.393701
                            Altezza = Altezza * 0.393701
                    End Select

                    If Boolean.Parse(LeggiStringOpzioni("UnitMeasureRounded", IT.Opzioni)) Then
                        NuovoContenuto = Regex.Replace(LeggiStringOpzioni("SizeStyle", IT.Opzioni), "(%H)", CInt(Altezza).ToString)
                        NuovoContenuto = Regex.Replace(NuovoContenuto, "(%W)", CInt(Larghezza).ToString)
                    Else
                        NuovoContenuto = Regex.Replace(LeggiStringOpzioni("SizeStyle", IT.Opzioni), "(%H)", Math.Round(Altezza, 2, MidpointRounding.AwayFromZero).ToString)
                        NuovoContenuto = Regex.Replace(NuovoContenuto, "(%W)", Math.Round(Larghezza, 2, MidpointRounding.AwayFromZero).ToString)
                    End If

                    IgnoraConversioneMM = True

                Case "II_AspectRatio"
                    Dim Temp_Ratio As Decimal
                    Dim Larghezza As Integer = FreeImageAPI.FreeImage.GetWidth(Img_dib)
                    Dim Altezza As Integer = FreeImageAPI.FreeImage.GetHeight(Img_dib)

                    If Larghezza <= Altezza Then
                        Temp_Ratio = Format(Altezza / Larghezza, "Fixed")
                    Else
                        Temp_Ratio = Format(Larghezza / Altezza, "Fixed")
                    End If
                    NuovoContenuto = Temp_Ratio
                    IgnoraConversioneMM = True

                Case "II_DPI"
                    'Calcolando il DPI manualmente si ottiene un valore molto più preciso, ottenendo dimensioni esatte
                    'Il metodo >>> FreeImageAPI.FreeImage.GetResolutionX(Img_dib) restituisce un valore DPI UInteger pertanto il calcolo delle dimensioni dell'immagine resulta impreciso
                    Dim DPIX As Double = 0.0254 * FreeImageAPI.FreeImage.GetDotsPerMeterX(Img_dib)
                    Dim DPIY As Double = 0.0254 * FreeImageAPI.FreeImage.GetDotsPerMeterY(Img_dib)

                    If DPIX = 0 Then DPIX = 72 '< nel caso di dpi non impostati, imposto i dpi standard
                    If DPIY = 0 Then DPIY = 72

                    If Boolean.Parse(LeggiStringOpzioni("UnitMeasureRounded", IT.Opzioni)) Then
                        NuovoContenuto = Regex.Replace(LeggiStringOpzioni("SizeStyle", IT.Opzioni), "(%H)", CInt(DPIY).ToString)
                        NuovoContenuto = Regex.Replace(NuovoContenuto, "(%W)", CInt(DPIX).ToString)
                    Else
                        NuovoContenuto = Regex.Replace(LeggiStringOpzioni("SizeStyle", IT.Opzioni), "(%H)", Math.Round(DPIY, 2, MidpointRounding.AwayFromZero).ToString)
                        NuovoContenuto = Regex.Replace(NuovoContenuto, "(%W)", Math.Round(DPIX, 2, MidpointRounding.AwayFromZero).ToString)
                    End If

                    IgnoraConversioneMM = True

                Case "II_BitsPerPixel"
                    NuovoContenuto = FreeImageAPI.FreeImage.GetBPP(Img_dib)
                    IgnoraConversioneMM = True

                Case "II_UniqueColors"
                    NuovoContenuto = FreeImageAPI.FreeImage.GetUniqueColors(Img_dib)
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

        FreeImageAPI.FreeImage.UnloadEx(Img_dib)

        'sostituzione termini
        If XMLSettings_Read("CATEGORIA_immagini_config_SostituzioneTermini").Equals("True") Then
            NomeRinominato = SostituzioneTermini(NomeRinominato, Coll_TerminiSostituzioni)
        End If

        'Check filename lenght limit
        CheckFilenameLenghtLimit(NomeRinominato)

        'Controllo eventuali caratteri non ammessi
        NomeRinominato = ValidateChar_Replace(NomeRinominato)

        NomeRinominato &= Estensione

        Return NomeRinominato
    End Function



    Private Function Get_Image_dib(PathFile As String) As FreeImageAPI.FIBITMAP
        Dim dib As FreeImageAPI.FIBITMAP = FreeImageAPI.FreeImage.LoadEx(PathFile)

        If dib.IsNull Then
            Throw New ARCoreException(Localization.Resource_Common.AR_Core_Immagini_OpenFileError)
        Else
            Return dib
        End If
    End Function

    Private Function Get_Image_MDMExif(dib As FreeImageAPI.FIBITMAP) As Object()
        Dim iMetaData As FreeImageAPI.Metadata.ImageMetadata = New FreeImageAPI.Metadata.ImageMetadata(dib)

        Dim iMDM_Exif_Main As FreeImageAPI.Metadata.MDM_EXIF_MAIN = iMetaData.Item(FreeImageAPI.FREE_IMAGE_MDMODEL.FIMD_EXIF_MAIN)
        Dim iMDM_Exif_Exif As FreeImageAPI.Metadata.MDM_EXIF_EXIF = iMetaData.Item(FreeImageAPI.FREE_IMAGE_MDMODEL.FIMD_EXIF_EXIF)

        Return {iMDM_Exif_Main, iMDM_Exif_Exif}
    End Function

    Private Function Get_Image_IPTC(dib As FreeImageAPI.FIBITMAP) As FreeImageAPI.Metadata.MDM_IPTC
        Dim iMetaData As FreeImageAPI.Metadata.ImageMetadata = New FreeImageAPI.Metadata.ImageMetadata(dib)

        Dim iMDM_IPTC As FreeImageAPI.Metadata.MDM_IPTC = iMetaData.Item(FreeImageAPI.FREE_IMAGE_MDMODEL.FIMD_IPTC)

        Return iMDM_IPTC
    End Function

    Private Function Get_Exif_Info(MDM As Object, TipoDatoExif As String) As String
        Try
            Dim Tipo As Type = MDM.GetType
            Dim PI As Reflection.PropertyInfo = Tipo.GetProperty(TipoDatoExif.Substring(TipoDatoExif.LastIndexOf("_") + 1))
            Dim Valore As String = PI.GetValue(MDM, Nothing)

            If String.IsNullOrEmpty(Valore) Then
                Return ""
            Else

                Select Case TipoDatoExif
                    Case "EXIF_MAIN_Orientation"
                        If Not String.IsNullOrEmpty(Valore) Then
                            Dim Risultato As Integer
                            If Integer.TryParse(Valore, Risultato) Then
                                If 1 <= Risultato <= 8 Then
                                    Valore = Localization.Resource_Struttura.ResourceManager.GetString("EXIF_MAIN_Orientation_Value_" & Risultato)
                                End If
                            End If
                        End If
                End Select

                Return Valore
            End If

        Catch ex As Exception
            Debug.Print(ex.Message)
            Return ""
        End Try
    End Function

End Class
