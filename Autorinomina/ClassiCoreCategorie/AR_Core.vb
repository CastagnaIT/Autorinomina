Imports System.Globalization
Imports System.Text.RegularExpressions

Public Class AR_Core
    Protected InfoRichiesteStruttura As New ArrayList()
    Protected ContenutoElencoSequenziale As String() = {}
    Protected OwnerWindow As Window
    Protected PatternRegexNumerazioni As New ArrayList
    Dim PatternRegexData As New ArrayList

    Protected Enum FINFO_VIDEO
        DIMENSIONI
        ASPECTRATIO
        FRAMERATE
        CODECVIDEO
        ENCODEDLIBRARYNAME
        DURATA
        CODECAUDIO
        CODECAUDIOLINGUA
        LINGUE
        LINGUESOTTOTITOLI
        NUMEROCAPITOLI
        Count
    End Enum

    Protected Enum FINFO_AUDIO
        DURATA
        TAG_NUMERO
        TAG_ARTISTA
        TAG_ALBUM
        TAG_TITOLO
        TAG_COMPOSITORE
        TAG_PUBLISCER
        TAG_GENERE
        TAG_DATA_RILASCIO
        TAG_DATA_REGISTRAZIONE
        FREQUENZA
        BITRATE
        MODALITA_BITRATE
        NUMERO_CANALI
        Count
    End Enum

    Protected Enum FINFO_PROPRIETA
        NOMEFILE
        DATACREAZIONE
        DATAULTIMAMODIFICA
        DATAULTIMOACCESSO
        Count
    End Enum

    Public Sub New(_OwnerWindow As Window)
        OwnerWindow = _OwnerWindow

        Try
            Dim FilePath As String = DataPath & "\" & XMLSettings_Read("CategoriaSelezionata") & "_elencosequenziale.txt"
            If IO.File.Exists(FilePath) Then
                Dim reader As IO.StreamReader = New IO.StreamReader(FilePath, System.Text.ASCIIEncoding.Unicode)
                Dim ContenutoTXT = reader.ReadToEnd()
                ContenutoElencoSequenziale = ContenutoTXT.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                reader.Close()
            End If
        Catch ex As Exception
            ContenutoElencoSequenziale = {}
        End Try

        'Genere di informazioni richieste dalla struttura scelta
        For Each item In Coll_Struttura
            If item.TipoDato.Contains("AR_") AndAlso Not InfoRichiesteStruttura.Contains("AR_") Then InfoRichiesteStruttura.Add("AR_")
            If item.TipoDato.Contains("TVDB_") AndAlso Not InfoRichiesteStruttura.Contains("TVDB_") Then InfoRichiesteStruttura.Add("TVDB_")
            If item.TipoDato.Contains("MI_") AndAlso Not InfoRichiesteStruttura.Contains("MI_") Then InfoRichiesteStruttura.Add("MI_")
            If item.TipoDato.Contains("PF_") AndAlso Not InfoRichiesteStruttura.Contains("PF_") Then InfoRichiesteStruttura.Add("PF_")
            If item.TipoDato.Contains("EXIF_") AndAlso Not InfoRichiesteStruttura.Contains("EXIF_") Then InfoRichiesteStruttura.Add("EXIF_")
            If item.TipoDato.Contains("II_") AndAlso Not InfoRichiesteStruttura.Contains("II_") Then InfoRichiesteStruttura.Add("II_")
        Next

        PatternRegexNumerazioni.AddRange(XMLSettings_Read("CATEGORIA_serietv_config_regex").Split(";"))

        PatternRegexData.Add("[0-9]{4}(\W)(0[1-9]|1[0-2])(\W)(0[1-9]|[1-2][0-9]|3[0-1])") 'formato yyyy-mm-gg
        PatternRegexData.Add("(0[1-9]|[1-2][0-9]|3[0-1])(\W)(0[1-9]|1[0-2])(\W)[0-9]{4}") 'formato gg-mm-yyyy
    End Sub

    Overridable Function RunCore(ByVal PathFile As String, ByVal NumeroSequenziale As Integer, ByVal LivePreview As Boolean, ByVal PaddingLenght As Integer, ByVal Optional overrideTipoFiltro As String = Nothing) As String
        Return ""
    End Function


    Protected Function SostituzioneTermini(Testo As String, ByRef CollTermini As CollItemsTermini) As String

        For Each termine In CollTermini
            Dim ResultMatches As MatchCollection

            If termine.CaseSensitive.Equals("True") Then

                'If SerieTvVariantMatch Then
                ''se il termine da cercare inizia con almeno 1 spazio es: " L " --> " L'" diventa "^L " --> "^L'"
                ''workaround per cercarlo ad esempio su "L ora della verita" essendo la serie tv divisa in due stringhe, si deve cercare anche a inizio stringa
                'If Regex.IsMatch(termine.Termine, "^\s{1}", RegexOptions.Compiled) Then
                ' 'sostituisco il primo spazio con ^ per sostituire parola a inizio stringa
                ' Dim PatternVariante As String = Regex.Replace(termine.Termine, "^\s{1}", "")
                ' Dim TermineSostituto As String = Regex.Replace(termine.TermineSostituto, "^\s{1}", "")
                ' Dim ResultMatchesVariante As MatchCollection = Regex.Matches(Testo, "^" & Regex.Escape(PatternVariante), RegexOptions.Compiled)
                '
                '                For n As Integer = ResultMatchesVariante.Count - 1 To 0 Step -1
                '                Testo = Testo.Remove(ResultMatchesVariante(n).Index, ResultMatchesVariante(n).Length).Insert(ResultMatchesVariante(n).Index, TermineSostituto)
                '        Next
                '            End If
                '            End If

                ResultMatches = Regex.Matches(Testo, Regex.Escape(termine.Termine), RegexOptions.None)
            Else
                ' If SerieTvVariantMatch Then
                ' If Regex.IsMatch(termine.Termine, "^\s{1}", RegexOptions.Compiled) Then
                ' 'sostituisco il primo spazio con ^ per sostituire parola a inizio stringa
                ' Dim PatternVariante As String = Regex.Replace(termine.Termine, "^\s{1}", "")
                ' Dim TermineSostituto As String = Regex.Replace(termine.TermineSostituto, "^\s{1}", "")
                ' Dim ResultMatchesVariante As MatchCollection = Regex.Matches(Testo, "^" & Regex.Escape(PatternVariante), RegexOptions.Compiled + RegexOptions.IgnoreCase)
                '
                '                For n As Integer = ResultMatchesVariante.Count - 1 To 0 Step -1
                '                Testo = Testo.Remove(ResultMatchesVariante(n).Index, ResultMatchesVariante(n).Length).Insert(ResultMatchesVariante(n).Index, TermineSostituto)
                '        Next
                '            End If
                '            End If

                ResultMatches = Regex.Matches(Testo, Regex.Escape(termine.Termine), RegexOptions.IgnoreCase)
            End If

            For n As Integer = ResultMatches.Count - 1 To 0 Step -1
                Testo = Testo.Remove(ResultMatches(n).Index, ResultMatches(n).Length).Insert(ResultMatches(n).Index, termine.TermineSostituto)
            Next
        Next

        Return Testo
    End Function

    Protected Function RimozioneTerminiBlackList(Testo As String, ByRef CollTermini As CollItemsTermini, Optional LasciaSpazi As Boolean = False) As String
        Dim TrePuntiConsecutivi As Boolean = Regex.IsMatch(Testo, "\.{3}", RegexOptions.IgnoreCase)

        For Each termine In CollTermini

            Dim RegexPattern As String = "([^a-z|A-Z|^0-9]|^|\||\G)" & Regex.Escape(termine.Termine).Replace("\ ", ".") & "([^a-z|A-Z|^0-9]|$|\|)"
            Dim RegexOptions As RegexOptions = RegexOptions.None ' RegexOptions.Compiled rimosso perchè triplica il tempo di esecuzione

            If CBool(termine.CaseSensitive) = False Then RegexOptions += RegexOptions.IgnoreCase

            If Regex.IsMatch(Testo, RegexPattern, RegexOptions) Then
                Dim ResultMatches As MatchCollection = Regex.Matches(Testo, RegexPattern, RegexOptions)

                For mIndex As Integer = ResultMatches.Count - 1 To 0 Step -1
                    Testo = Testo.Remove(ResultMatches(mIndex).Index, ResultMatches(mIndex).Length)
                    Testo = Testo.Insert(ResultMatches(mIndex).Index, "|") 'aggiungere questo corregge il riconoscimento per i termini sucessivi
                    'esempio:
                    ' Breaking.Bad.5x06.Buonuscita.ITA.BDRip.x264-NovaRip
                    ' Breaking.Bad.5x06.Buonuscita.ITA.BDRip|    <<<<< rimosso ITA e sostituito con |
                    ' Breaking.Bad.5x06.Buonuscita|BDRip|    <<<<< senza | risulterebbe: Breaking.Bad.5x06.BuonuscitaBDRip    >>>>> e la parola BDRip non sarebbe rimossa perchè adiacente al testo
                    ' Breaking.Bad.5x06.Buonuscita|
                Next
            End If
        Next

        If LasciaSpazi Then
            Testo = Testo.Replace("|", " ")
        Else
            Testo = Testo.Replace("|", "")
        End If

        'rimozione parentesi vuote
        Dim ResultMatchesP As MatchCollection
        Dim RegexPatternP As String = "\[\]"
        ResultMatchesP = Regex.Matches(Testo, RegexPatternP, RegexOptions.None)
        For Each match As Match In ResultMatchesP
            Testo = Testo.Remove(match.Index, match.Length)
        Next

        RegexPatternP = "\(\)"
        ResultMatchesP = Regex.Matches(Testo, RegexPatternP, RegexOptions.None)
        For Each match As Match In ResultMatchesP
            Testo = Testo.Remove(match.Index, match.Length)
        Next

        If Not TrePuntiConsecutivi Then
            'rimuovo i caratteri "." rimasti dalla rimozione dei termini BL
            Testo = Regex.Replace(Testo, "\.{4,20}", "")
        End If

        Return Testo
    End Function

    Protected Function Filtro_Generico(NomeFile As String, nomecategoria As String, ByRef Coll_TerminiBlackList As CollItemsTermini, ByRef Coll_TerminiSostituzioni As CollItemsTermini, Optional NonCercareNumerazione As Boolean = False) As String()
        Dim Risultato_Numerazione As String = ""
        Dim Risultato_TestoContenuto As String = ""
        Dim Risultato_Data As String = ""
        Dim Risultato_Anno As String = ""
        Dim CarattereSplit As String = Space(1)
        Dim WeakFilter As Boolean = XMLSettings_Read("CATEGORIA_" & nomecategoria & "_config_WeakFilter").Equals("True")

        Dim NomeFile_Filtrato As String = IO.Path.GetFileNameWithoutExtension(NomeFile)
        NomeFile_Filtrato = RimozioneTerminiBlackList(NomeFile_Filtrato, Coll_TerminiBlackList, True)

        'Cerco la numerazione
        If Not NonCercareNumerazione Then
            Dim Modified_PatternRegexNumerazioni As New ArrayList
            Modified_PatternRegexNumerazioni.AddRange(PatternRegexNumerazioni)
            Modified_PatternRegexNumerazioni.Add("[0-9]+")

            Dim RegexCustomOptions As New RegexOptions
            RegexCustomOptions = IIf(XMLSettings_Read("CATEGORIA_" & nomecategoria & "_config_cercaNumerazioneRightToLeft").Equals("True"), RegexOptions.IgnoreCase + RegexOptions.RightToLeft, RegexOptions.IgnoreCase)

            For Each Pattern As String In Modified_PatternRegexNumerazioni
                If Regex.IsMatch(NomeFile_Filtrato, "\s*" & Pattern & "\s*", RegexCustomOptions) Then

                    Dim m As Match = Regex.Match(NomeFile_Filtrato, "\s*" & Pattern & "\s*", RegexCustomOptions)
                    Risultato_Numerazione = m.Value

                    NomeFile_Filtrato = NomeFile_Filtrato.Remove(m.Index, m.Length)
                    NomeFile_Filtrato = NomeFile_Filtrato.Insert(m.Index, Space(1)).Trim

                    Exit For
                End If
            Next
        End If

        'elimino eventuali (...) oppure [...] e ... per un riconoscimento piu preciso del carattere da usare nello split
        NomeFile_Filtrato = Regex.Replace(NomeFile_Filtrato, "\(.*?\)", "")
        NomeFile_Filtrato = Regex.Replace(NomeFile_Filtrato, "\[.*?\]", "")
        NomeFile_Filtrato = Regex.Replace(NomeFile_Filtrato, "\.{3}", "")

        If WeakFilter = False Then
            'determino la quantità  di caratteri speciali presenti
            Dim QtaCharSpec() As Integer = {0, 0, 0}
            QtaCharSpec(0) = Regex.Split(NomeFile_Filtrato, "\.").Length
            QtaCharSpec(1) = Regex.Split(NomeFile_Filtrato, "_").Length
            QtaCharSpec(2) = Regex.Split(NomeFile_Filtrato, "\-").Length

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

        'Rimozione ()[] anche con contenuti inclusi
        If XMLSettings_Read("CATEGORIA_" & nomecategoria & "_config_EliminaParentesi").Equals("True") Then
            NomeFile_Filtrato = RimozioneParentesi(NomeFile_Filtrato, True).Trim
        End If

        If WeakFilter = False Then
            'esecuzione dello split
            Dim ContenutoNomeFileSplittato As New ArrayList
            Dim StrSplitted() As String = Regex.Split(NomeFile_Filtrato, "\" & CarattereSplit)

            For x As Integer = 0 To StrSplitted.Length - 1
                If Not StrSplitted(x).Trim.Length = 0 Then ContenutoNomeFileSplittato.Add(StrSplitted(x))
            Next

            'altre operazioni per filtrare il contenuto
            For x As Integer = 0 To ContenutoNomeFileSplittato.Count - 1
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
            Next

        Else

            Risultato_TestoContenuto = NomeFile_Filtrato
        End If

        'Elimino questi eventuali caratteri rimanenti .-_ dall'inizio e dalla fine
        Dim RegexPatterns As String() = {"\s*\-+\s*$", "\s*\.+\s*$", "\s*_+\s*$", "^\s*\-+\s*", "^\s*\.+\s*", "^\s*_+\s*"}
        For Each Pattern As String In RegexPatterns
            If Regex.IsMatch(Risultato_TestoContenuto, Pattern) = True Then
                Risultato_TestoContenuto = Regex.Replace(Risultato_TestoContenuto, Pattern, "")
            End If
        Next

        'Cerco la data
        Risultato_Data = EstrapolaData(Risultato_TestoContenuto)

        'Cerco l'anno, se prensente
        If Regex.IsMatch(Risultato_TestoContenuto, "(\s*\d{4}\s*)", RegexOptions.IgnoreCase + RegexOptions.RightToLeft) = True Then
            If Coll_Struttura.FirstOrDefault(Function(i) i.TipoDato.Contains("AR_Anno")) IsNot Nothing Then
                Dim MatchResult As Match = Regex.Match(Risultato_TestoContenuto, "(\s*\d{4}\s*)", RegexOptions.RightToLeft)
                Risultato_Anno = MatchResult.Value

                'elimino l'anno dal nome del file
                Risultato_TestoContenuto = Risultato_TestoContenuto.Remove(MatchResult.Index, MatchResult.Length)
                Risultato_TestoContenuto = Risultato_TestoContenuto.Insert(MatchResult.Index, Space(1)).Trim
            End If
        End If

        'Elimino eventuali parentesi vuote
        Risultato_TestoContenuto = RimozioneParentesi(Risultato_TestoContenuto, False)

        'Eventuali Trim degli spazi all'interno delle parentesi "(  esempio" --> "(esempio"
        Risultato_TestoContenuto = TrimInternoParentesi(Risultato_TestoContenuto)


        Return {Risultato_Numerazione, Risultato_TestoContenuto, Risultato_Data, Risultato_Anno}
    End Function

    Protected Function CreaNuovaNumerazione(Numerazione As String, Stile As Integer, PaddingZeros As Integer) As String
        Dim NumeroPos1 As Integer
        Dim NumeroPos2 As Integer = IIf(IsNumeric(Numerazione), Numerazione, 0) 'serve nel caso: numero provenienza "6" stile serie tv ---> 0x06
        Dim NumeroPos2Concat As Integer
        Dim NumeroSingolo As Integer

        If String.IsNullOrEmpty(Numerazione) Then Return ""

        If Numerazione.Contains("x") Then
            'formato 0x00
            Dim tmp As String() = Numerazione.Split("x")
            NumeroPos1 = Integer.Parse(tmp(0))

            If tmp(1).Contains("-") Then
                'Contiene un doppio episodio
                NumeroPos2 = Integer.Parse(tmp(1).Split("-")(0))
                NumeroPos2Concat = Integer.Parse(tmp(1).Split("-")(1))
            Else
                NumeroPos2 = Integer.Parse(tmp(1))
            End If
            NumeroSingolo = NumeroPos2
        Else
            'Contiene solo un numero
            If Integer.TryParse(Numerazione, NumeroSingolo) = False Then
                Return Numerazione
            End If
        End If


        Select Case Stile
            Case 0 'Stile normale  1,2,3 ...
                Return NumeroSingolo.ToString("D0")

            Case 1 'Stile normale con 0;  01,02,03 ...
                Return NumeroSingolo.ToString("D" & PaddingZeros)

            Case 2 'Stile romano  I,II,III ...
                Return New NumeriRomaniConverter().ToRomani(NumeroSingolo)

            Case 3 'Stile alfabetico AA,AB,AC ...
                Return Converti_NumeroToAlfabetico(NumeroSingolo, True)

            Case 4 'Stile alfabetico AA BB CC
                Return Converti_NumeroToAlfabetico(NumeroSingolo, False)

            Case 5 'Stile serietv europeo 0x00
                Return NumeroPos1.ToString & "x" & NumeroPos2.ToString("D" & PaddingZeros) & IIf(Not NumeroPos2Concat = 0, "-" & NumeroPos2Concat.ToString("D2"), "")

            Case 6 'Stile serietv europeo 0.00
                Return NumeroPos1.ToString & "." & NumeroPos2.ToString("D" & PaddingZeros) & IIf(Not NumeroPos2Concat = 0, "-" & NumeroPos2Concat.ToString("D2"), "")

            Case 7 'Stile serietv americano S01E01
                Return "S" & NumeroPos1.ToString("D2") & "E" & NumeroPos2.ToString("D" & PaddingZeros) & IIf(Not NumeroPos2Concat = 0, "-" & NumeroPos2Concat.ToString("D" & PaddingZeros), "")

            Case 8 'Serie Tv Numerico
                Return NumeroPos1.ToString("D2") & NumeroPos2.ToString("D" & PaddingZeros) & IIf(Not NumeroPos2Concat = 0, NumeroPos2Concat.ToString("D" & PaddingZeros), "")

            Case Else
                Throw New Exception("Style number '" & Stile & "' not recognized.")
        End Select
    End Function

    Protected Function EstrapolaInfoRegexPattern(Filename As String, Pattern As String) As String
        Try
            If Regex.IsMatch(Filename, Pattern) Then
                Dim m As Match = Regex.Match(Filename, Pattern)
                Return m.Value
            Else
                Return ""
            End If
        Catch ex As Exception
            Return "RegexError"
        End Try
    End Function

    Protected Function EstrapolaData(ByRef Filename As String) As String
        If PatternRegexData.Count > 0 Then
            For x As Integer = 0 To PatternRegexData.Count - 1

                If Regex.IsMatch(Filename, PatternRegexData(x), RegexOptions.IgnoreCase) Then
                    Dim m As Match = Regex.Match(Filename, PatternRegexData(x), RegexOptions.IgnoreCase)

                    Filename = Filename.Remove(m.Index, m.Length)
                    Filename = Filename.Insert(m.Index, Space(1)).Trim

                    Return m.Value
                End If
            Next
        End If
        Return ""
    End Function

    Protected Function GetVideoInfo(PathFileName As String) As String()
        Try
            Dim Val(FINFO_VIDEO.Count - 1) As String
            Dim MI As MediaInfo = New MediaInfo

            MI.Open(PathFileName)
            Dim Temp As String

            'Dimensioni
            Temp = MI.Get_(StreamKind.Video, 0, "Width")
            If Not String.IsNullOrEmpty(Temp) Then
                Val(FINFO_VIDEO.DIMENSIONI) = Temp & "x" & MI.Get_(StreamKind.Video, 0, "Height")
            End If

            'Aspect ratio
            Temp = MI.Get_(StreamKind.Video, 0, "DisplayAspectRatio/String")
            If Not String.IsNullOrEmpty(Temp) Then
                Val(FINFO_VIDEO.ASPECTRATIO) = Temp.Replace(":", ".")
            End If

            'Framerate
            Temp = MI.Get_(StreamKind.Video, 0, "FrameRate/String")
            If Not String.IsNullOrEmpty(Temp) Then
                Try
                    Temp = Format(CDbl(Temp.Replace(".", ",").ToLower.Replace("fps", "")), "0.###")
                Catch ex As Exception
                End Try

                Val(FINFO_VIDEO.FRAMERATE) = Temp
            End If

            'Codec video
            Temp = MI.Get_(StreamKind.Video, 0, "Format") 'vecchia versione dll utilizzava "CodecID/Hint"
            If Not String.IsNullOrEmpty(Temp) Then
                Val(FINFO_VIDEO.CODECVIDEO) = Temp
            End If

            'Libreria di codifica
            Temp = MI.Get_(StreamKind.Video, 0, "Encoded_Library_Name")
            If Not String.IsNullOrEmpty(Temp) Then
                Val(FINFO_VIDEO.ENCODEDLIBRARYNAME) = Temp
            End If

            'Durata (in ms)
            Temp = MI.Get_(StreamKind.Video, 0, "Duration")
            If String.IsNullOrEmpty(Temp) Then
                Val(FINFO_VIDEO.DURATA) = "0"
            Else
                Val(FINFO_VIDEO.DURATA) = Temp.Split(".")(0) 'formato 1243960.000000
            End If

            'Codec audio
            Temp = MI.Get_(StreamKind.General, 0, "Audio_Format_WithHint_List")
            If Not String.IsNullOrEmpty(Temp) Then
                Dim Temp2 As String = ""
                For Each righe As String In Temp.Split("/")
                    If Not String.IsNullOrEmpty(righe.Trim) Then
                        Temp2 &= righe.Trim & "-"
                    End If
                Next

                If Temp2.Length > 0 Then
                    Val(FINFO_VIDEO.CODECAUDIO) = Temp2.Substring(0, Temp2.Length - 1)
                End If
            End If

            'Codec audio con lingua
            Temp = MI.Get_(StreamKind.General, 0, "Audio_Format_WithHint_List")
            If Not String.IsNullOrEmpty(Temp) Then
                Dim Temp2 As String = MI.Get_(StreamKind.General, 0, "Audio_Language_List")
                Dim Temp3 As String = ""
                Dim ARStrCA() As String = Temp.Split("/")
                Dim ARStrLNG() As String = Temp2.Split("/")

                For x As Integer = 0 To ARStrCA.Length - 1
                    If (ARStrLNG.Length - 1) > x Then
                        Temp3 &= ARStrCA(x).Trim & Space(1) & ARStrLNG(x).Trim.Substring(0, 3).ToUpper & "-"
                    Else
                        Temp3 &= ARStrCA(x).Trim & Space(1) & "-"
                    End If
                Next

                If Not String.IsNullOrEmpty(Temp3) Then
                    Val(FINFO_VIDEO.CODECAUDIOLINGUA) = Temp3.Substring(0, Temp3.Length - 1)
                End If
            End If

            'Lingue tracce audio
            Temp = MI.Get_(StreamKind.General, 0, "Audio_Language_List")
            If Not String.IsNullOrEmpty(Temp) Then
                Dim Temp2 As String = ""
                For Each righe As String In Temp.Split("/")
                    If Not String.IsNullOrEmpty(righe.Trim) Then
                        Temp2 &= righe.Trim.Substring(0, 3).ToUpper & "-"
                    End If
                Next

                If Not String.IsNullOrEmpty(Temp2) Then
                    Val(FINFO_VIDEO.LINGUE) = Temp2.Substring(0, Temp2.Length - 1)
                End If
            End If

            'Lingue sottotitoli
            Temp = MI.Get_(StreamKind.General, 0, "Text_Language_List")
            If Not String.IsNullOrEmpty(Temp) Then
                Dim Temp2 As String = ""
                For Each righe As String In Temp.Split("/")
                    If Not String.IsNullOrEmpty(righe.Trim) Then
                        Temp2 &= righe.Trim.Substring(0, 3).ToUpper & "-"
                    End If
                Next

                If Not String.IsNullOrEmpty(Temp2) Then
                    Val(FINFO_VIDEO.LINGUE) = Temp2.Substring(0, Temp2.Length - 1)
                End If
            End If

            'Numero dei capitoli
            Temp = MI.Get_(StreamKind.General, 0, "ChaptersCount")
            If String.IsNullOrEmpty(Temp) Then
                Val(FINFO_VIDEO.NUMEROCAPITOLI) = "0"
            Else
                Val(FINFO_VIDEO.NUMEROCAPITOLI) = Temp
            End If

            MI.Close()

            Return Val
        Catch ex As Exception
            Throw New Exception("Error read file (GetVideoInfo)" & IO.Path.GetFileName(PathFileName), ex.InnerException)
        End Try
    End Function

    Protected Function GetAudioInfo(PathFileName As String) As String()
        Try
            Dim Val(FINFO_AUDIO.Count - 1) As String
            Dim MI As MediaInfo = New MediaInfo

            MI.Open(PathFileName)
            Dim Temp As String

            'Durata
            Temp = MI.Get_(StreamKind.Audio, 0, "Duration")
            If String.IsNullOrEmpty(Temp) Then
                Val(FINFO_AUDIO.DURATA) = "0"
            Else
                Val(FINFO_AUDIO.DURATA) = Temp
            End If

            'Tag numero
            Temp = MI.Get_(StreamKind.General, 0, "Track/Position")
            If Not String.IsNullOrEmpty(Temp) Then
                If Temp.Contains("/") Then 'fix per numerazione con numero complessivo es: 5/12   6/12
                    Temp = Temp.Split("/")(0)
                End If
                If XMLSettings_Read("CATEGORIA_audio_config_EscludiCaratteriNumerazione").Equals("True") Then
                    If Regex.IsMatch(Temp, "[-+]?([0-9]*\.)?[0-9]+([eE][-+]?[0-9]+)?", RegexOptions.IgnoreCase) Then
                        Temp = Regex.Match(Temp, "[-+]?([0-9]*\.)?[0-9]+([eE][-+]?[0-9]+)?", RegexOptions.IgnoreCase).Value.ToString
                    End If
                End If
                Val(FINFO_AUDIO.TAG_NUMERO) = Temp
            End If

            'Tag artista
            Temp = MI.Get_(StreamKind.General, 0, "Performer")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.TAG_ARTISTA) = Temp

            'Tag album
            Temp = MI.Get_(StreamKind.General, 0, "Album")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.TAG_ALBUM) = Temp

            'Tag titolo
            Temp = MI.Get_(StreamKind.General, 0, "Track")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.TAG_TITOLO) = Temp

            'Tag compositore
            Temp = MI.Get_(StreamKind.General, 0, "Composer")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.TAG_COMPOSITORE) = Temp

            'Tag publiscer
            Temp = MI.Get_(StreamKind.General, 0, "Publisher")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.TAG_PUBLISCER) = Temp

            'Tag genere
            Temp = MI.Get_(StreamKind.General, 0, "Genre")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.TAG_GENERE) = Temp

            'Tag data rilascio
            Temp = MI.Get_(StreamKind.General, 0, "Released_Date")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.TAG_DATA_RILASCIO) = Temp

            'Tag data registrazione
            Temp = MI.Get_(StreamKind.General, 0, "Recorded_Date")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.TAG_DATA_REGISTRAZIONE) = Temp

            'Frequenza
            Temp = MI.Get_(StreamKind.Audio, 0, "SamplingRate/String")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.FREQUENZA) = Temp

            'Bitrate
            Temp = MI.Get_(StreamKind.Audio, 0, "BitRate/String")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.BITRATE) = Temp

            'Modalita bitrate
            Temp = MI.Get_(StreamKind.Audio, 0, "BitRate_Mode")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.MODALITA_BITRATE) = Temp

            'Numero canali
            Temp = MI.Get_(StreamKind.Audio, 0, "Channel(s)")
            If Not String.IsNullOrEmpty(Temp) Then Val(FINFO_AUDIO.NUMERO_CANALI) = Temp

            If Val.Count(Function(s) String.IsNullOrEmpty(s)) >= 8 Then
                'se ci sono almeno 8 campi vuoti significa che non ci sono tag nel file
                If XMLSettings_Read("CATEGORIA_audio_config_ID3TagMissing").Equals("True") Then
                    Throw New ARCoreException(Localization.Resource_Common.AR_Core_Error_ID3MetadataMissing)
                End If
            End If

            MI.Close()
            Return Val
        Catch ex_core As ARCoreException
            Throw New ARCoreException(ex_core.Message)
        Catch ex As Exception
            Throw New Exception("Error read file (GetAudioInfo)" & IO.Path.GetFileName(PathFileName), ex.InnerException)
        End Try
    End Function

    Protected Function GetPropertyInfo(PathFileName As String) As String()
        Try
            Dim Val(FINFO_PROPRIETA.Count - 1) As String
            Dim FI As New IO.FileInfo(PathFileName)

            Val(FINFO_PROPRIETA.NOMEFILE) = IO.Path.GetFileNameWithoutExtension(PathFileName)
            Val(FINFO_PROPRIETA.DATACREAZIONE) = FI.CreationTime.ToString
            Val(FINFO_PROPRIETA.DATAULTIMAMODIFICA) = FI.LastWriteTime.ToString
            Val(FINFO_PROPRIETA.DATAULTIMOACCESSO) = FI.LastAccessTime.ToString

            Return Val
        Catch ex As Exception
            Throw New Exception("Error read file (GetPropertyInfo)" & IO.Path.GetFileName(PathFileName), ex.InnerException)
        End Try
    End Function

    Protected Function RimozioneParentesi(Testo As String, AncheContenuto As Boolean) As String
        Dim RegexPatterns As New ArrayList

        If AncheContenuto Then
            RegexPatterns.Add("\s*\[.*\]\s*")
            RegexPatterns.Add("\s*\(.*\)\s*")
        Else
            RegexPatterns.Add("\s*\[\s*\]\s*")
            RegexPatterns.Add("\s*\(\s*\)\s*")
        End If

        For Each Pattern As String In RegexPatterns

            Dim ResultMatches As MatchCollection = Regex.Matches(Testo, Pattern, RegexOptions.IgnoreCase)
            For n As Integer = ResultMatches.Count - 1 To 0 Step -1
                Testo = Testo.Remove(ResultMatches(n).Index, ResultMatches(n).Length)
                Testo = Testo.Insert(ResultMatches(n).Index, Space(1))
            Next
        Next

        Return Testo
    End Function

    Protected Function TrimInternoParentesi(Testo As String) As String
        Dim RegexPatterns() As String = {"\(\s*", "\s*\)", "\[\s*", "\s*\]"}

        For Each Pattern As String In RegexPatterns
            Testo = Regex.Replace(Testo, Pattern, "")
        Next

        Return Testo
    End Function

    Protected Function FormattaData(Data As String, Pattern As String, TVDB_Date As Boolean) As String
        Dim _Data As Date = Nothing
        Dim Result As Boolean = False


        Result = Date.TryParse(Data, _Data)
        If Not Result Then Result = Date.TryParseExact(Data, "dd-MM-yyyy", New CultureInfo("en-US"), DateTimeStyles.None, _Data)
        If Not Result Then Result = Date.TryParseExact(Data, "yyyy-MM-dd", New CultureInfo("en-US"), DateTimeStyles.None, _Data)
        If Not Result Then Result = Date.TryParseExact(Data, "yyyy", New CultureInfo("en-US"), DateTimeStyles.None, _Data)

        If Result Then
            Try
                Return Format(_Data, Pattern)
            Catch ex As Exception
                Debug.Print(ex.Message)
                Return Data
            End Try
        Else
            Return Data
        End If
    End Function

    Public Function ConversioneMaiuscoleMinuscole(Testo As String, Stile As String, SeparatorePresente As Boolean) As String
        If String.IsNullOrEmpty(Testo) Then Return Testo
        Select Case Stile
            Case 0 'Invariato
                Return Testo

            Case 1 'Capitalizza normalmente
                Return StrConv(Testo, VbStrConv.ProperCase)

            Case 2 'Capitalizza Mix
                Testo = Testo.ToLower
                Dim ResultMatches As MatchCollection = Regex.Matches(Testo, "\b\S{4,}\b", RegexOptions.IgnoreCase)
                For Each match As Match In ResultMatches
                    'prima rimuovo la parola da modificare
                    Testo = Testo.Remove(match.Index, match.Length)
                    'reinserisco la parola modificata
                    Testo = Testo.Insert(match.Index, StrConv(match.Value, VbStrConv.ProperCase))
                Next
                Return Testo

            Case 3, 4 'Solo prima lettera maiuscola dell'intera stringa (4 = prima lettera maiuscola anche dopo ogni separatore scelto)
                If SeparatorePresente Then
                    Testo = Testo.ToLower
                    Dim ResultMatches As MatchCollection = Regex.Matches(Testo, "\b\S*\b", RegexOptions.IgnoreCase)
                    If ResultMatches.Count >= 1 Then
                        'prima rimuovo la parola da modificare
                        Testo = Testo.Remove(ResultMatches(0).Index, ResultMatches(0).Length)
                        'reinserisco la parola modificata
                        Testo = Testo.Insert(ResultMatches(0).Index, StrConv(ResultMatches(0).Value, VbStrConv.ProperCase))
                    End If

                    Return Testo
                Else
                    Return Testo.ToLower
                End If

            Case 5 'Tutto minuscolo
                Return StrConv(Testo, VbStrConv.Lowercase)

            Case 6 'Tutto MAIUSCOLO
                Return StrConv(Testo, VbStrConv.Uppercase)

            Case Else
                Return Testo
        End Select
    End Function

    Protected Function ConvertiMillisecondToTime(MS_Time As Integer, Stile As String) As String
        Try
            Dim TS As New TimeSpan(0, 0, 0, 0, MS_Time)
            Dim Ore As Integer = TS.Hours
            Dim Minuti As Integer = TS.Minutes
            Dim Secondi As Integer = TS.Seconds

            Stile = Regex.Replace(Stile, "(%H)", IIf(Ore = 0, "", Ore.ToString))
            Stile = Regex.Replace(Stile, "(%M)", Minuti.ToString)
            Stile = Regex.Replace(Stile, "(%S)", Secondi.ToString)

            Return Stile
        Catch ex As Exception
            Debug.Print(ex.Message)
            Return ""
        End Try
    End Function

    Protected Function ConversioneMaiuscoleMinuscole_EXT(Testo As String, Stile As String) As String
        If String.IsNullOrEmpty(Testo) Then Return Testo
        Select Case Stile
            Case 0 'Invariato
                Return Testo

            Case 1 'Tutto minuscolo
                Return StrConv(Testo, VbStrConv.Lowercase)

            Case 2 'Tutto MAIUSCOLO
                Return StrConv(Testo, VbStrConv.Uppercase)

            Case Else
                Return Testo
        End Select
    End Function

    Private Function CN(ByVal Numero As Integer, ByVal Elenco() As String) As String
        Dim temp As Integer = Numero

        While (temp > 26)
            temp = (temp - 26)
        End While

        Return Elenco(temp).ToUpper
    End Function

    Private Function Converti_NumeroToAlfabetico(Numero As Integer, Modalita_1 As Boolean) As String
        Dim Temp As Integer = 0
        Dim Temp_Valore As String = ""
        Dim ElencoCaratteri() As String = {"", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}

        If Modalita_1 Then
            'AA, AB, AC, AD ...
            Temp = IIf(((Numero) / 26) = Math.Floor((Numero) / 26), Math.Floor((Numero) / 26) - 1, Math.Floor((Numero) / 26)) 'arrotondo x difetto
            Temp_Valore = CN((Numero), ElencoCaratteri)
            While Temp > 26
                Temp_Valore = CN(Temp, ElencoCaratteri) & Temp_Valore
                Temp = IIf(((Numero) / 26) = Math.Floor((Numero) / 26), Math.Floor((Numero) / 26) - 1, Math.Floor((Numero) / 26))
            End While
            If Temp > 0 Then Temp_Valore = CN(Temp, ElencoCaratteri) & Temp_Valore
        Else
            'A, AA, AAA, AAA ...
            Temp = IIf(((Numero) / 26) = Math.Floor((Numero) / 26), Math.Floor((Numero) / 26) - 1, Math.Floor((Numero) / 26))

            For x = 0 To Temp
                Temp_Valore &= CN((Numero), ElencoCaratteri)
            Next
        End If

        Return Temp_Valore
    End Function

    Protected Function GetIndexOfBiggestArrayItem(TheArray() As Integer) As Integer
        'This function gives max value of int array without sorting an array
        Dim MaxIntegersIndex As Integer = 0

        For i As Integer = 1 To UBound(TheArray)
            If TheArray(i) > TheArray(MaxIntegersIndex) Then
                MaxIntegersIndex = i
            End If
        Next
        'index of max value
        Return MaxIntegersIndex
    End Function

End Class
