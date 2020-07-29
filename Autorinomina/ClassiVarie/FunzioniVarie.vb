Imports System.IO
Imports System.Text.RegularExpressions

Module FunzioniVarie

    Public Sub AggiungiDatoStrutturaCollection(ByVal TipoDato As String, ByVal Contenuto As String)
        AggiungiDatoStrutturaCollection(TipoDato, Contenuto, True)
    End Sub

    Public Sub AggiungiDatoStrutturaCollection(ByVal TipoDato As String, ByVal Contenuto As String, ByVal InSequenza As Boolean)
        Dim Colore As Brush = Brushes.Black
        Dim Testo = Localization.Resource_Struttura.ResourceManager.GetString(TipoDato)

        If TipoDato.Contains("TVDB_") Then

            Colore = New BrushConverter().ConvertFromString("#10B01A")
        ElseIf TipoDato.Contains("MI_") Then

            Colore = New BrushConverter().ConvertFromString("#FF1410")
        ElseIf TipoDato.Contains("AR_") Then

            Colore = New BrushConverter().ConvertFromString("#2B66C3")
        ElseIf TipoDato.Contains("PF_") Then

            Colore = New BrushConverter().ConvertFromString("#FF00CE")
        ElseIf TipoDato.Contains("EXIF_") Then

            Colore = New BrushConverter().ConvertFromString("#D9801C")
        ElseIf TipoDato.Contains("IPTC_") Then

            Colore = New BrushConverter().ConvertFromString("#25D136")
        ElseIf TipoDato.Contains("II_") Then

            Colore = New BrushConverter().ConvertFromString("#3F4CB6")
        ElseIf TipoDato.Contains("STD_Testo") Then

            Testo = LeggiStringOpzioni("Testo", Contenuto)
        ElseIf TipoDato.Contains("STD_Separatore") Then

            Testo = LeggiStringOpzioni("SeparatoreCarattere", Contenuto)
        ElseIf TipoDato.Contains("STD_ElencoSequenziale") Then

            Colore = New BrushConverter().ConvertFromString("#4E2769")
        ElseIf TipoDato.Contains("STD_NumerazioneSequenziale") Then

            Colore = New BrushConverter().ConvertFromString("#4E2769")
        ElseIf TipoDato.Contains("EXT_Opzioni") Then

            Colore = New BrushConverter().ConvertFromString("#FF00CE")
        End If

        Coll_Struttura.Add(New ItemStruttura(TipoDato, Testo, Contenuto, Colore))
    End Sub

    Public Sub RiabilitaMenuCategoria(ByRef CM As ContextMenu)
        For Each MI As MenuItem In CM.Items

            If MI.Items.Count > 0 Then
                For Each Sub_MI As MenuItem In MI.Items
                    If MI.Tag Is Nothing Then Continue For
                    If Sub_MI.IsEnabled = False Then
                        Sub_MI.IsEnabled = True

                        Dim TB As TextBlock
                        If TypeOf Sub_MI.Header Is TextBlock Then
                            TB = Sub_MI.Header
                        Else
                            TB = New TextBlock With {.Text = Sub_MI.Header.ToString}
                            Sub_MI.Header = TB
                        End If

                        TB.TextDecorations = Nothing
                    End If
                Next

            Else
                If MI.Tag Is Nothing Then Continue For
                If MI.IsEnabled = False Then
                    MI.IsEnabled = True

                    Dim TB As TextBlock
                    If TypeOf MI.Header Is TextBlock Then
                        TB = MI.Header
                    Else
                        TB = New TextBlock With {.Text = MI.Header.ToString}
                        MI.Header = TB
                    End If

                    TB.TextDecorations = Nothing
                End If
            End If
        Next
    End Sub

    Public Sub RiabilitaMenuCategoria(ByVal UserData As String, ByVal abilita As Boolean, ByRef CM As ContextMenu)
        For Each MI As MenuItem In CM.Items

            If MI.Items.Count > 0 Then
                For Each Sub_MI As MenuItem In MI.Items
                    If Sub_MI.Tag Is Nothing Then Continue For
                    If Sub_MI.Tag.ToString.Equals(UserData) Then
                        Sub_MI.IsEnabled = abilita

                        Dim TB As TextBlock
                        If TypeOf Sub_MI.Header Is TextBlock Then
                            TB = Sub_MI.Header
                        Else
                            TB = New TextBlock With {.Text = Sub_MI.Header.ToString}
                            Sub_MI.Header = TB
                        End If

                        If abilita Then
                            TB.TextDecorations = Nothing
                        Else
                            TB.TextDecorations = TextDecorations.Strikethrough
                        End If

                        Return
                    End If
                Next

            Else
                If MI.Tag Is Nothing Then Continue For
                If MI.Tag.ToString.Equals(UserData) Then
                    MI.IsEnabled = abilita

                    Dim TB As TextBlock
                    If TypeOf MI.Header Is TextBlock Then
                        TB = MI.Header
                    Else
                        TB = New TextBlock With {.Text = MI.Header.ToString}
                        MI.Header = TB
                    End If

                    If abilita Then
                        TB.TextDecorations = Nothing
                    Else
                        TB.TextDecorations = TextDecorations.Strikethrough
                    End If

                    Return
                End If
            End If
        Next
    End Sub




    Function ValidateChar_Replace(ByVal Stringa As String, Optional ByVal EscludiAncheParentesi As Boolean = False) As String
        If EscludiAncheParentesi Then
            Return Regex.Replace(Stringa, "\[|\]|\(|\)|\\|\/|\:|\*|\?|\""|\<|\>|\|", ".")
        Else
            Return Regex.Replace(Stringa, "\\|\/|\:|\*|\?|\""|\<|\>|\|", ".")
        End If
    End Function

    Function ValidateNoSpecialChar(ByVal Stringa As String, ByVal EscludiAncheParentesi As Boolean) As Boolean
        If EscludiAncheParentesi Then
            Return Regex.IsMatch(Stringa, "\[|\]|\(|\)|\\|\/|\:|\*|\?|\""|\<|\>|\|")
        Else
            Return Regex.IsMatch(Stringa, "\\|\/|\:|\*|\?|\""|\<|\>|\|")
        End If
    End Function


    Public Function GetExtensions(_CategoriaSelezionata As String) As Object()
        Dim Exts As String = "*;"
        Dim ListExts As New List(Of String)
        Dim Name As String = ""
        Dim Result As String = ""

        Select Case _CategoriaSelezionata
            Case "CATEGORIA_serietv", "CATEGORIA_video"
                Exts = XMLSettings_Read("Extensions_video")
                Name = Localization.Resource_WND_Impostazioni.Extensions_Video
            Case "CATEGORIA_immagini"
                Exts = XMLSettings_Read("Extensions_images")
                Name = Localization.Resource_WND_Impostazioni.Extensions_Images
            Case "CATEGORIA_audio"
                Exts = XMLSettings_Read("Extensions_audio")
                Name = Localization.Resource_WND_Impostazioni.Extensions_Audio
        End Select

        For Each ext As String In Exts.Split(",")
            Result &= "*." & ext & ";"
            ListExts.Add(ext)
        Next

        Return {Result, ListExts, Name}
    End Function

    Public Function GetFolderSize(ByVal DirPath As String, Optional IncludeSubFolders As Boolean = True) As Long
        Dim lngDirSize As Long
        Dim objFileInfo As FileInfo
        Dim objDir As DirectoryInfo = New DirectoryInfo(DirPath)
        Dim objSubFolder As DirectoryInfo

        Try

            'add length of each file
            For Each objFileInfo In objDir.GetFiles()
                lngDirSize += objFileInfo.Length
            Next

            'call recursively to get sub folders
            'if you don't want this set optional
            'parameter to false 
            If IncludeSubFolders Then
                For Each objSubFolder In objDir.GetDirectories()
                    lngDirSize += GetFolderSize(objSubFolder.FullName)
                Next
            End If

        Catch Ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & Ex.Message, MsgBoxStyle.Critical, Application.Current.MainWindow.Title)
        End Try

        Return lngDirSize
    End Function

    Public Sub CancellaCacheTVDB()
        Dim Files As String() = System.IO.Directory.GetFiles(IO.Path.Combine(DataPath, "cache"), "*.*", IO.SearchOption.AllDirectories)
        For Each File As String In Files
            If IO.Path.GetFileName(File.ToLower) <> "languages.xml" Then
                IO.File.Delete(File)
            End If
        Next

        Dim SottoCartelle As String() = System.IO.Directory.GetDirectories(IO.Path.Combine(DataPath, "cache"))
        For Each Cartella As String In SottoCartelle
            System.IO.Directory.Delete(Cartella)
        Next
    End Sub

    Public Function FormatBYTE(ByVal b As Double, Optional NoDecimali As Boolean = False) As String
        Dim bSize(8) As String
        Dim i As Integer

        bSize(0) = "Bytes"
        bSize(1) = "KB" 'Kilobytes
        bSize(2) = "MB" 'Megabytes
        bSize(3) = "GB" 'Gigabytes
        bSize(4) = "TB" 'Terabytes
        bSize(5) = "PB" 'Petabytes
        bSize(6) = "EB" 'Exabytes
        bSize(7) = "ZB" 'Zettabytes
        bSize(8) = "YB" 'Yottabytes

        b = CDbl(b) ' Make sure var is a Double (not just variant)
        Dim risultato As String = "0 Byte"
        For i = UBound(bSize) To 0 Step -1
            If b >= (1024 ^ i) Then
                If NoDecimali Then
                    risultato = Format$(CInt((b / (1024 ^ i)))) & " " & bSize(i)
                Else
                    risultato = ThreeNonZeroDigits(b / (1024 ^ i)) & " " & bSize(i)
                End If
                Exit For
            End If
        Next

        Return risultato
    End Function
    Private Function ThreeNonZeroDigits(ByVal value As Double) As String
        If value >= 100 Then
            ' No digits after the decimal.
            Return Format$(CInt(value))
        ElseIf value >= 10 Then
            ' One digit after the decimal.
            Return Format$(value, "0.0")
        Else
            ' Two digits after the decimal.
            Return Format$(value, "0.00")
        End If
    End Function

    Public Sub CheckFilenameLenghtLimit(ByRef Filename As String)
        If Filename.Length > 260 Then
            Filename = Filename.Substring(0, 260)
        End If
    End Sub
End Module
