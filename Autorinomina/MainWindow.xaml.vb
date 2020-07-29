Imports System.ComponentModel
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.Win32

Class MainWindow
    Dim CM_Categoria As New ContextMenu()
    Dim PopUp_Tooltip As New CustomToolTip
    Dim MouseOnLV_ScrollBar As Boolean = False
    WithEvents Timer_AutoCattura As New Windows.Threading.DispatcherTimer
    WithEvents Timer_DelayLivePreview As New Windows.Threading.DispatcherTimer
    WithEvents BW_CheckNewSWVersion As New BackgroundWorker
    Dim Thread_RunLivePreview As New Thread(AddressOf Sub_Thread_RunLivePreview)
    Dim RowMenu As ContextMenu
    Dim CM_MM As ContextMenu
    Dim BloccoOperazioni As Boolean = False

#Region "Funzioni per Index On Item MouseOver su ListView"
    Delegate Function GetPositionDelegate(element As IInputElement) As Point

    Private Function GetListViewItem(index As Integer) As ListViewItem
        If LV_StrutturaNomeFile.ItemContainerGenerator.Status <> Primitives.GeneratorStatus.ContainersGenerated Then
            Return Nothing
        End If

        Return TryCast(LV_StrutturaNomeFile.ItemContainerGenerator.ContainerFromIndex(index), ListViewItem)
    End Function

    Private Function IsMouseOverTarget(target As Visual, getPosition As GetPositionDelegate) As Boolean
        Dim bounds As Rect = VisualTreeHelper.GetDescendantBounds(target)
        Dim mousePos As Point = getPosition(DirectCast(target, IInputElement))
        Return bounds.Contains(mousePos)
    End Function

    Private Function GetCurrentIndex(getPosition As GetPositionDelegate) As Integer
        Dim index As Integer = -1
        For i As Integer = 0 To LV_StrutturaNomeFile.Items.Count - 1
            Dim item As ListViewItem = GetListViewItem(i)
            If IsMouseOverTarget(item, getPosition) Then
                index = i
                Exit For
            End If
        Next
        Return index
    End Function
#End Region

#Region "Guida veloce"
    Private Sub SGuide_Click(sender As Object, e As MouseButtonEventArgs)
        Dim SGuide As SuggestionGuide = sender
        If SGuide.CurrentPage = 4 Then
            Grid_Principale.Children.Remove(SGuide)
        Else
            SGuide.CambiaPaginaGuida(SGuide.CurrentPage + 1)
        End If
    End Sub

    Private Sub VisualizzaGuida()
        If Grid_Principale.Children.OfType(Of SuggestionGuide).Count = 0 Then
            Dim SGuide As New SuggestionGuide
            AddHandler SGuide.PreviewMouseLeftButtonUp, AddressOf SGuide_Click
            Grid.SetRow(SGuide, 0)
            Grid.SetRowSpan(SGuide, 3)
            Grid_Principale.Children.Add(SGuide)
            SGuide.Focus()
        End If
    End Sub
#End Region

    Private Sub MainWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Try
            RowMenu = FindResource("RowMenu")
            CM_MM = FindResource("CM_MM_Menu")

            Dim ColonneVisibili As New ArrayList
            ColonneVisibili.AddRange(XMLSettings_Read("ColonneVisibili").Split("/"))
            Dim ColonneVisibiliWidth As New ArrayList
            ColonneVisibiliWidth.AddRange(XMLSettings_Read("ColonneVisibiliWidth").Split("/"))

            Dim PassaAvanti As Integer = 0
            For x As Integer = 2 To DG_Files.Columns.Count - 1
                Dim Col As DataGridColumn = DG_Files.Columns(x)

                Dim MI_Col As New MenuItem
                MI_Col.Header = Localization.Resource_WND_Main.ResourceManager.GetString("Column_" & Col.SortMemberPath)
                MI_Col.Tag = Col.SortMemberPath
                MI_Col.IsCheckable = True
                AddHandler MI_Col.Checked, AddressOf MI_Menu_MostraColonne_Action
                AddHandler MI_Col.Unchecked, AddressOf MI_Menu_MostraColonne_Action

                If ColonneVisibili.Contains(Col.SortMemberPath) Then
                    Col.DisplayIndex = ColonneVisibili.IndexOf(Col.SortMemberPath) + 2

                    Dim ColWidth As DataGridLength = New DataGridLength(Double.Parse(ColonneVisibiliWidth(ColonneVisibili.IndexOf(Col.SortMemberPath))), DataGridLengthUnitType.Pixel)
                    Col.Width = ColWidth
                    MI_Col.IsChecked = True
                Else
                    Col.DisplayIndex = ColonneVisibili.Count + PassaAvanti + 2
                    Col.Visibility = Visibility.Collapsed
                    PassaAvanti += 1
                End If

                MI_Menu_MostraColonne.Items.Add(MI_Col)
            Next

            Me.Height = Double.Parse(XMLSettings_Read("MainWindow_Height"))
            Me.Width = Double.Parse(XMLSettings_Read("MainWindow_Width"))

            Dim CVS As ListCollectionView = CollectionViewSource.GetDefaultView(Coll_Files)
            CVS.IsLiveSorting = True
            CVS.LiveSortingProperties.Add("Index")

            DG_Files.ItemsSource = Coll_Files

            LV_StrutturaNomeFile.ItemsSource = Coll_Struttura

            For Each SP As StackPanel In CB_Menu_Categoria.Items
                If SP.Tag.ToString = XMLSettings_Read("CategoriaSelezionata") Then
                    CB_Menu_Categoria.SelectedItem = SP
                    Exit For
                End If
            Next

            PopUp_Tooltip.Inizializza()
            AddHandler PopUp_Tooltip.UpdateLivePreviewEvent, AddressOf UpdateLivePreview
            PopUp_LVItem.Child = PopUp_Tooltip


            Dim PathSaved As String = XMLSettings_Read("FavoriteFolderPath_" & 1)
            If Not String.IsNullOrEmpty(PathSaved) AndAlso New IO.DirectoryInfo(PathSaved).Exists Then
                MI_Menu_AggiungiPreferite.Tag = PathSaved
                MI_Menu_AggiungiPreferite.Header = Localization.Resource_WND_Main.MI_AggiungiFilesDaCartellaPreferita & " '" & New IO.DirectoryInfo(PathSaved).Name & "'"
                BTN_Menu_AggiungiPreferiti.ToolTip = Localization.Resource_WND_Main.MI_AggiungiFilesDaCartellaPreferita & " '" & New IO.DirectoryInfo(PathSaved).Name & "'"
            Else
                MI_Menu_AggiungiPreferite.Tag = ""
                MI_Menu_AggiungiPreferite.Header = Localization.Resource_WND_Main.MI_AggiungiFilesDaCartellaPreferita & " '...'"
                BTN_Menu_AggiungiPreferiti.ToolTip = Localization.Resource_WND_Main.MI_AggiungiFilesDaCartellaPreferita & " '...'"
            End If


            MI_Menu_IncludiSottocartelle.IsChecked = Boolean.Parse(XMLSettings_Read("VerificaFilesIncludiSottocartelle"))
            MI_Menu_MostraBarra.IsChecked = Boolean.Parse(XMLSettings_Read("BarraStrumenti_Visibile"))
            MI_Menu_LivePreview.IsChecked = Boolean.Parse(XMLSettings_Read("MainWindow_LivePreview"))

            Timer_AutoCattura.Interval = New TimeSpan(0, 0, 1)
            If Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchFiles")) OrElse Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchPreset")) Then
                Timer_AutoCattura.Start()
            End If
            MI_Menu_CatturaFiles.IsChecked = Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchFiles"))
            MI_Menu_CatturaPreset.IsChecked = Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchPreset"))


            If Boolean.Parse(XMLSettings_Read("MainWindow_AlwaysOnTop")) AndAlso Boolean.Parse(XMLSettings_Read("MainWindow_AlwaysOnTop_Keep")) Then
                MI_Menu_AlwaysOnTop.IsChecked = True
            End If

            If Boolean.Parse(XMLSettings_Read("MainWindow_GuideOnStart")) Then
                XMLSettings_Save("MainWindow_GuideOnStart", "False")
                VisualizzaGuida()
            End If

            UpdateUI()

            Timer_DelayLivePreview.Interval = New TimeSpan(0, 0, 0, 0, 800)
            AddHandler GestioneXML.UpdateLivePreviewEvent, AddressOf UpdateLivePreview

        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & Space(1) & ex.Message, MsgBoxStyle.Critical, "Autorinomina")
            End
        End Try
    End Sub

    Private Sub UpdateLivePreview()
        'questo metodo viene richiamato ogni volta che si modifica in qualsiasi modo i dati della struttura (non è strettamente legato alla livepreview)
        If BTN_Anteprima.IsEnabled AndAlso BloccoOperazioni Then
            BloccoOperazioni = False

            BTN_Rinomina.IsEnabled = False
            BTN_RinominaSelezione.IsEnabled = False
            RowMenu.Items(0).IsEnabled = False

            BTN_Ripristina.IsEnabled = False
            BTN_RipristinaSelezione.IsEnabled = False
            RowMenu.Items(1).IsEnabled = False
        End If

        If Boolean.Parse(XMLSettings_Read("MainWindow_LivePreview")) AndAlso DG_Files.SelectedItem IsNot Nothing Then
            If Coll_Files.Count > 0 Then
                If Timer_DelayLivePreview.IsEnabled = False Then Timer_DelayLivePreview.Start()
            End If
        End If
    End Sub

    Private Sub Timer_DelayLivePreview_Tick(sender As Object, e As EventArgs) Handles Timer_DelayLivePreview.Tick
        If Coll_Struttura.Count <= 1 Then
            If Coll_Struttura.FirstOrDefault(Function(i) i.TipoDato.Contains("EXT_")) IsNot Nothing OrElse Coll_Struttura.Count = 0 Then
                Timer_DelayLivePreview.Stop()
                Return
            End If
        End If

        If Coll_Files.Count > 0 AndAlso BloccoOperazioni = False Then
            Try
                If Thread_RunLivePreview.IsAlive Then Thread_RunLivePreview.Abort()
                Thread_RunLivePreview = New Thread(AddressOf Sub_Thread_RunLivePreview)
                Thread_RunLivePreview.Start({DG_Files.Items, DG_Files.SelectedItems})
            Catch ex As Exception
                If Thread_RunLivePreview.IsAlive Then Thread_RunLivePreview.Abort()
            End Try
        End If
        Timer_DelayLivePreview.Stop()
    End Sub

    Private Sub Sub_Thread_RunLivePreview(ByVal param As Object)
        Try
            Dim AR_Preview As New AR_Core_RunPreview(Me)
            AR_Preview.StartPreview(CType(param(0), ItemCollection), CType(param(1), IList), True)
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub

    Private Sub MainWindow_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        BW_CheckNewSWVersion.RunWorkerAsync()
    End Sub

    Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Dim ColonneVisibili As String = ""
        Dim ColonneVisibiliWidth As String = ""

        For Each col In DG_Files.Columns.OrderBy(Function(c) c.DisplayIndex)
            If col.SortMemberPath = "Stato" OrElse col.SortMemberPath = "StatoInfo" Then Continue For

            If col.Visibility = Visibility.Visible Then
                ColonneVisibili &= col.SortMemberPath & "/"
                ColonneVisibiliWidth &= col.Width.Value.ToString & "/"
            End If
        Next
        XMLSettings_Save("ColonneVisibili", ColonneVisibili.Remove(ColonneVisibili.Length - 1, 1))
        XMLSettings_Save("ColonneVisibiliWidth", ColonneVisibiliWidth.Remove(ColonneVisibiliWidth.Length - 1, 1))

        XMLSettings_Save("MainWindow_Height", Me.Height.ToString)
        XMLSettings_Save("MainWindow_Width", Me.Width.ToString)

        XMLSettings_WriteFile()
        XMLStrutture_Save("Default")
    End Sub

    Private Sub CreaMenuCategoria()
        CM_Categoria = New ContextMenu()
        CM_Categoria.UseLayoutRounding = True

        Dim AL_TipoMenu As New ArrayList()
        Dim AL_UserData As New ArrayList()
        Dim AL_Color As New ArrayList()

        'Definizione dei menu da generare

        AL_TipoMenu.AddRange(New String() {"MI", "MI", "MI", "MI", "MI", "M"})
        AL_UserData.AddRange(New String() {"#MenuAddInfo", "STD_Testo", "STD_Separatore", "STD_ElencoSequenziale", "STD_NumerazioneSequenziale", "STD_FiltraNomeFile"})
        AL_Color.AddRange(New String() {"", "", "", "#4E2769", "#4E2769", "#2B66C3"})

        If XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_serietv") OrElse XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_video") Then
            AL_TipoMenu.Add("M")
            AL_UserData.Add("STD_InfoMultimediali")
            AL_Color.Add("#FF1410")
        End If

        If XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_serietv") Then
            AL_TipoMenu.Add("M")
            AL_UserData.Add("STD_TheTvdb")
            AL_Color.Add("#10B01A")
        End If

        If XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_audio") Then
            AL_TipoMenu.Add("M")
            AL_UserData.Add("STD_InfoAudio")
            AL_Color.Add("#FF1410")

            AL_TipoMenu.Add("M")
            AL_UserData.Add("STD_InfoID3Tag")
            AL_Color.Add("#FF1410")
        End If

        If XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_immagini") Then
            AL_TipoMenu.Add("M")
            AL_UserData.Add("STD_InfoImmagine")
            AL_Color.Add("#3F4CB6")

            AL_TipoMenu.Add("M")
            AL_UserData.Add("STD_InfoEXIFGeneriche")
            AL_Color.Add("#D9801C")

            AL_TipoMenu.Add("M")
            AL_UserData.Add("STD_InfoEXIFFotocamera")
            AL_Color.Add("#D9801C")

            AL_TipoMenu.Add("M")
            AL_UserData.Add("STD_InfoIPTC")
            AL_Color.Add("#25D136")
        End If

        AL_TipoMenu.Add("M")
        AL_UserData.Add("STD_ProprietaFile")
        AL_Color.Add("#FF00CE")

        AL_TipoMenu.Add("MI")
        AL_UserData.Add("EXT_Opzioni")
        AL_Color.Add("#FF00CE")

        'Creazione dei controlli
        For x As Integer = 0 To AL_UserData.Count - 1

            If AL_TipoMenu(x) = "MI" Then
                Dim MI_new As MenuItem
                If AL_UserData(x).ToString.Contains("#") Then
                    MI_new = New MenuItem
                    MI_new.Header = Localization.Resource_Struttura.ResourceManager.GetString(AL_UserData(x).ToString.Replace("#", ""))
                    MI_new.Style = FindResource("MICustomSeparator")
                Else
                    MI_new = New MenuItem
                    MI_new.Header = Localization.Resource_Struttura.ResourceManager.GetString(AL_UserData(x).ToString)
                    If AL_UserData(x).ToString.Contains("STD_NumerazioneSequenziale") OrElse AL_UserData(x).ToString.Contains("EXT_Opzioni") Then
                        MI_new.Tag = AL_UserData(x) 'permette di essere riabilitato dopo essere stato inserito
                    End If
                    AddHandler MI_new.Click, AddressOf CategoriaMenu_Click
                End If
                CM_Categoria.Items.Add(MI_new)

            Else

                Dim M As New MenuItem()
                M.Header = Localization.Resource_Struttura.ResourceManager.GetString(AL_UserData(x).ToString)
                Dim UserData As String() = New String() {""}

                If AL_UserData(x).Equals("STD_FiltraNomeFile") Then
                    If XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_serietv") Then
                        UserData = New String() {"AR_Numerazione", "AR_TitoloEpisodio", "AR_TitoloSerieTv", "AR_Data", "AR_RegexPattern"}

                    ElseIf XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_video") Then
                        UserData = New String() {"AR_NomeDelFile", "AR_Data", "AR_Anno", "AR_RegexPattern"}

                    ElseIf XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_immagini") Then
                        UserData = New String() {"AR_Numerazione", "AR_NomeDelFile", "AR_Data", "AR_Anno", "AR_RegexPattern"}

                    ElseIf XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_audio") Then
                        UserData = New String() {"AR_Numerazione", "AR_NomeDelFile", "AR_Data", "AR_Anno", "AR_RegexPattern"}

                    ElseIf XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_generica") Then
                        UserData = New String() {"AR_Numerazione", "AR_NomeDelFile", "AR_Data", "AR_Anno", "AR_RegexPattern"}

                    End If
                End If

                If AL_UserData(x).Equals("STD_InfoMultimediali") Then
                    UserData = New String() {"MI_Durata", "MI_Lingue", "MI_LingueSottotitoli", "MI_TotaleCapitoli", "#MI_Separator_Codifica", "MI_RisoluzioneFilmato",
                        "MI_AspectRatio", "MI_FrameRate", "MI_CodecVideo", "MI_EncodedLibraryName", "MI_CodecAudio", "MI_CodecAudioLingua"}
                End If

                If AL_UserData(x).Equals("STD_InfoAudio") Then
                    UserData = New String() {"MI_AudioDurata", "#MI_Separator_Codifica", "MI_AudioFrequenza", "MI_AudioBitRate", "MI_AudioModalitaBitRate", "MI_AudioNumeroCanali"}
                End If

                If AL_UserData(x).Equals("STD_InfoID3Tag") Then
                    UserData = New String() {"MI_AudioTagNumero", "MI_AudioTagArtista", "MI_AudioTagCompositore", "MI_AudioTagAlbum", "MI_AudioTagTitolo", "MI_AudioTagGenere",
                                         "MI_AudioTagDataRilascio", "MI_AudioTagDataRegistrazione", "MI_AudioTagPubliscer"}
                End If

                If AL_UserData(x).Equals("STD_TheTvdb") Then '"TVDB_NumeroStagioni" rimosso deprecato con TVDB API v2, mantengo impostazioni se in futuro viene reimplementato
                    UserData = New String() {"TVDB_TitoloSerieTv", "TVDB_TitoloEpisodio", "TVDB_DataPrimaTv", "TVDB_Creator", "TVDB_Director", "TVDB_Genere", "TVDB_Network", "TVDB_NumeroEpisodi"}
                End If

                If AL_UserData(x).Equals("STD_InfoImmagine") Then
                    UserData = New String() {"II_Dimensione", "II_AspectRatio", "II_DPI", "II_BitsPerPixel", "II_UniqueColors"}
                End If

                If AL_UserData(x).Equals("STD_InfoEXIFGeneriche") Then
                    UserData = New String() {"EXIF_MAIN_Artist", "EXIF_MAIN_Compression", "EXIF_MAIN_Copyright", "EXIF_MAIN_DateTime", "EXIF_MAIN_EquipmentModel", "EXIF_MAIN_ImageDescription",
                        "EXIF_MAIN_ImageHeight", "EXIF_MAIN_ImageWidth", "EXIF_MAIN_Make", "EXIF_MAIN_Orientation", "EXIF_MAIN_Software", "EXIF_MAIN_ResolutionUnit",
                        "EXIF_MAIN_XResolution", "EXIF_MAIN_YResolution"}
                End If

                If AL_UserData(x).Equals("STD_InfoEXIFFotocamera") Then
                    UserData = New String() {"EXIF_ApertureValue", "EXIF_BrightnessValue", "EXIF_ColorSpace", "EXIF_CompressedBitPerPixel", "EXIF_DateTimeDigitized", "EXIF_DateTimeOriginal",
                        "EXIF_DigitalZoomRatio", "EXIF_ExposureBiasValue", "EXIF_ExposureMode", "EXIF_ExposureTime", "EXIF_Flash", "EXIF_FlashEnergy",
                        "EXIF_FocalLenght", "EXIF_MaxApertureValue", "EXIF_PixelXDimension", "EXIF_PixelYDimension", "EXIF_Saturation", "EXIF_SceneCaptureType",
                        "EXIF_SensingMethod", "EXIF_Sharpness", "EXIF_ShutterSpeedValue", "EXIF_WhiteBalance"}
                End If

                If AL_UserData(x).Equals("STD_InfoIPTC") Then
                    UserData = New String() {"#IPTC_Separator_Contact", "IPTC_ByLine", "IPTC_ByLineTitle", "#IPTC_Separator_Details", "IPTC_DateCreated", "IPTC_City", "IPTC_ProvinceState", "IPTC_CountryPrimaryLocationName",
                        "IPTC_HeadLine", "IPTC_CaprionAbstract", "IPTC_WriterEditor", "#IPTC_Separator_State", "IPTC_ObjectName", "IPTC_OriginalTransmissionReference", "IPTC_SpecialInstructions", "IPTC_Category", "IPTC_Credit",
                        "IPTC_Source", "IPTC_CopyrightNotice"}
                End If

                If AL_UserData(x).Equals("STD_ProprietaFile") Then
                    UserData = New String() {"PF_NomeFile", "PF_DataCreazione", "PF_DataUltimaModifica", "PF_DataUltimoAccesso"}
                End If

                CreaSubMenu(M, UserData, AL_Color(x))
                CM_Categoria.Items.Add(M)
            End If


            Dim MI As MenuItem = CM_Categoria.Items.GetItemAt(CM_Categoria.Items.Count - 1)
            MI.Tag = AL_UserData(x)
            If String.IsNullOrEmpty(AL_Color(x)) = False Then
                MI.Foreground = New BrushConverter().ConvertFromString(AL_Color(x))
            End If

            If Not (AL_UserData(x).ToString.Contains("#")) Then
                Try
                    Dim uri As Uri = New Uri("pack://application:,,,/AutoRinomina;component/Immagini/iconeMenuCategorie/" & AL_UserData(x) & ".png", UriKind.Absolute)
                    MI.Icon = New Image() With {.Source = New BitmapImage(uri), .Height = 16, .Width = 16}
                Catch ex As Exception
                End Try
            End If
        Next
    End Sub

    Private Sub CreaSubMenu(ByRef M As MenuItem, ByRef List As String(), ByRef Colore As String)
        For x As Integer = 0 To List.Length - 1
            Dim MI As MenuItem

            If List(x).Contains("#") Then
                MI = New MenuItem
                MI.Header = Localization.Resource_Struttura.ResourceManager.GetString(List(x).ToString.Replace("#", ""))
                MI.Style = FindResource("MICustomSeparator")
            Else
                MI = New MenuItem
                MI.Header = Localization.Resource_Struttura.ResourceManager.GetString(List(x))
                MI.Tag = List(x)
                MI.Foreground = New BrushConverter().ConvertFromString(Colore)
                AddHandler MI.Click, AddressOf CategoriaMenu_Click
            End If
            M.Items.Add(MI)
        Next
    End Sub

    Private Sub CategoriaMenu_Click(sender As Object, e As RoutedEventArgs)
        Dim MI As MenuItem = sender
        Dim ContenutoDati As String = ""

        'Quando si modifica la struttura resetto a Default perchè non corrisponde più al Preset selezionato
        If Not XMLSettings_Read("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata")) = "Default" Then
            XMLSettings_Save("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata"), "Default")
        End If

        Select Case MI.Tag.ToString
            Case "STD_Testo"
                Dim CreaNuovo As Boolean = True

                While CreaNuovo

                    Dim WND As New DLG_Aggiungi_TXT(Nothing)
                    WND.Owner = Me
                    WND.ShowDialog()
                    If WND.DialogResult = True Then
                        If WND.Testo = "" Then Return

                        ContenutoDati = "<STD_Testo><Testo><![CDATA[" & WND.Testo & "]]></Testo></STD_Testo>"
                        AggiungiDatoStrutturaCollection(MI.Tag.ToString, ContenutoDati, False)
                        UpdateLivePreview()

                        CreaNuovo = WND.CreaNuovo
                    Else
                        Return
                    End If
                End While

            Case Else
                AggiungiDatoStrutturaCollection(MI.Tag.ToString, ContenutoDati, False)
                UpdateLivePreview()
        End Select
    End Sub

    Private Sub AggiornaMenuCategorie()
        'limina a 1 inserimento solo i campi aggiungi
        For Each item In Coll_Struttura
            If item.TipoDato.Contains("STD_NumerazioneSequenziale") OrElse
                item.TipoDato.Contains("TVDB_") OrElse
                item.TipoDato.Contains("MI_") OrElse
                item.TipoDato.Contains("AR_") OrElse
                item.TipoDato.Contains("PF_") OrElse
                item.TipoDato.Contains("EXIF_") OrElse
                item.TipoDato.Contains("IPTC_") OrElse
                item.TipoDato.Contains("II_") OrElse
                item.TipoDato.Contains("EXT_") Then

                RiabilitaMenuCategoria(item.TipoDato, False)
            End If
        Next
    End Sub

    Private Sub RiabilitaMenuCategoria(ByVal UserData As String, ByVal abilita As Boolean)
        FunzioniVarie.RiabilitaMenuCategoria(UserData, abilita, CM_Categoria)
    End Sub

    Private Sub BTN_AggiungiInfo_Click(sender As Object, e As RoutedEventArgs) Handles BTN_AggiungiInfo.Click
        AggiornaMenuCategorie()

        CM_Categoria.PlacementTarget = BTN_AggiungiInfo
        CM_Categoria.Placement = Primitives.PlacementMode.Bottom
        CM_Categoria.IsOpen = True
    End Sub

    Private Sub MI_Menu_CambiaCategoria_Click(sender As Object, e As RoutedEventArgs)
        Dim MI As MenuItem = sender

        If Not MI.Tag.ToString.Equals(XMLSettings_Read("CategoriaSelezionata")) Then
            For Each SP As StackPanel In CB_Menu_Categoria.Items
                If SP.Tag.ToString = MI.Tag.ToString Then
                    CB_Menu_Categoria.SelectedItem = SP
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub CB_Menu_Categoria_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CB_Menu_Categoria.SelectionChanged
        If Coll_Files.Count > 0 Then
            Dim Result As MsgBoxResult = MsgBox(Localization.Resource_WND_Main.CategoryUserChange.Replace("\n", vbCrLf), MsgBoxStyle.Question + MsgBoxStyle.YesNo, Localization.Resource_WND_Main.MI_CambiaCategoria)
            If Result = MsgBoxResult.Yes Then
                Coll_Files.Clear()
                UpdateUI()
            Else
                'workaround per annullare il selectionChanged, l'annullamento non è supportato
                RemoveHandler CB_Menu_Categoria.SelectionChanged, AddressOf CB_Menu_Categoria_SelectionChanged

                For Each SP As StackPanel In CB_Menu_Categoria.Items
                    If SP.Tag.ToString = XMLSettings_Read("CategoriaSelezionata") Then
                        CB_Menu_Categoria.SelectedItem = SP
                        Exit For
                    End If
                Next

                AddHandler CB_Menu_Categoria.SelectionChanged, AddressOf CB_Menu_Categoria_SelectionChanged

                Return
            End If
        End If

        XMLSettings_Save("CategoriaSelezionata", CType(CB_Menu_Categoria.SelectedItem, StackPanel).Tag)

        XMLStrutture_Read(Nothing)

        CreaMenuCategoria()
        PopUp_Tooltip.CM_Categoria = CM_Categoria

        Me.Title = "Autorinomina - " & Localization.Resource_WND_Main.ResourceManager.GetString("Nome_" & CType(CB_Menu_Categoria.SelectedItem, StackPanel).Tag.ToString)

        TB_Menu_Sostituisci.IsChecked = Boolean.Parse(XMLSettings_Read(XMLSettings_Read("CategoriaSelezionata") & "_config_SostituzioneTermini"))

        For Each ITEM As MenuItem In CM_MM.Items
            If ITEM.IsChecked Then ITEM.IsChecked = False
        Next
        CType(CM_MM.Items.GetItemAt(Integer.Parse(XMLSettings_Read(XMLSettings_Read("CategoriaSelezionata") & "_config_StileMaiuscoleMinuscole"))), MenuItem).IsChecked = True

        For Each ITEM As MenuItem In MI_Menu_MaiuscoleMinuscole.Items
            If ITEM.IsChecked Then ITEM.IsChecked = False
        Next
        CType(MI_Menu_MaiuscoleMinuscole.Items.GetItemAt(Integer.Parse(XMLSettings_Read(XMLSettings_Read("CategoriaSelezionata") & "_config_StileMaiuscoleMinuscole"))), MenuItem).IsChecked = True
    End Sub

    Private Sub LV_StrutturaNomeFile_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles LV_StrutturaNomeFile.SelectionChanged
        Dim LVI As ListViewItem = LV_StrutturaNomeFile.ItemContainerGenerator.ContainerFromItem(LV_StrutturaNomeFile.SelectedItem)
        If LVI Is Nothing OrElse MouseOnLV_ScrollBar Then
            PopUp_LVItem.IsOpen = False
            Return
        End If

        LV_StrutturaNomeFile.Tag = True 'serve per prevenire il caricamento multiplo dai dati causato dal nel mousemove 


        Dim popupSize As New Size(PopUp_LVItem.ActualWidth, PopUp_LVItem.ActualHeight)
        Dim Posizione As Point = LVI.TransformToAncestor(Me).Transform(New Point(0, 0))

        Posizione.X = (Posizione.X - (274 / 2)) + (LVI.ActualWidth / 2)
        Posizione.Y += 6

        PopUp_Tooltip.AvoidSave = True
        PopUp_Tooltip.AggiornaContenutoTT(LV_StrutturaNomeFile.SelectedItem)
        PopUp_LVItem.PlacementRectangle = New Rect(Posizione, popupSize)

        Dispatcher.BeginInvoke(Windows.Threading.DispatcherPriority.ApplicationIdle, Sub()
                                                                                         PopUp_LVItem.IsOpen = True
                                                                                         PopUp_Tooltip.AvoidSave = False
                                                                                         LV_StrutturaNomeFile.Tag = False 'serve per prevenire il caricamento multiplo dai dati causato dal nel mousemove 
                                                                                     End Sub)

    End Sub

    Private Sub LV_StrutturaNomeFile_PreviewMouseMove(sender As Object, e As MouseEventArgs) Handles LV_StrutturaNomeFile.PreviewMouseMove
        If Me.IsActive = False OrElse MouseOnLV_ScrollBar Then Return
        Dim Index As Integer = GetCurrentIndex(AddressOf e.GetPosition)
        If Index = -1 Then Return
        If LV_StrutturaNomeFile.Items(Index) Is Nothing Then Return


        If LV_StrutturaNomeFile.SelectedIndex = Index AndAlso LV_StrutturaNomeFile.Tag = False Then
            If PopUp_LVItem.IsOpen = False Then
                LV_StrutturaNomeFile_SelectionChanged(Me, Nothing)
            End If
        Else
            LV_StrutturaNomeFile.SelectedIndex = Index
        End If
    End Sub

    Private Async Sub PopUp_LVItem_MouseLeave(sender As Object, e As MouseEventArgs) Handles PopUp_LVItem.MouseLeave
        Await Task.Delay(1200)
        If LV_StrutturaNomeFile.IsMouseOver = False AndAlso PopUp_LVItem.IsMouseOver = False Then
            If PopUp_LVItem.IsOpen Then PopUp_LVItem.IsOpen = False
        End If
    End Sub

    Private Async Sub LV_StrutturaNomeFile_MouseLeave(sender As Object, e As MouseEventArgs) Handles LV_StrutturaNomeFile.MouseLeave
        Await Task.Delay(400)
        If LV_StrutturaNomeFile.IsMouseOver = False AndAlso PopUp_LVItem.IsMouseOver = False Then
            If PopUp_LVItem.IsOpen Then PopUp_LVItem.IsOpen = False
        End If
    End Sub

    Private Sub MI_TerminiBlackList_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_TerminiBlackList.Click
        Dim WND As New WND_Termini("BlackList")
        WND.Owner = Me
        WND.ShowDialog()
        UpdateLivePreview()
    End Sub

    Private Sub MI_SostituzioneTermini_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_SostituzioneTermini.Click
        Dim WND As New WND_Termini("Sostituzioni")
        WND.Owner = Me
        WND.ShowDialog()
        UpdateLivePreview()
    End Sub

    Private Sub MI_ElencoSequenziale_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_ElencoSequenziale.Click, BTN_Menu_ElencoSequenziale.Click
        Dim WND As New WND_ElencoSequenziale
        WND.Owner = Me
        WND.ShowDialog()
        UpdateLivePreview()
    End Sub

    Private Sub MI_CambiaLingua_Click(sender As Object, e As RoutedEventArgs) Handles MI_CambiaLingua.Click
        Dim WND As New WND_Languages
        WND.Owner = Me
        WND.ShowDialog()
    End Sub

    Private Sub MI_InformazioniSu_Click(sender As Object, e As RoutedEventArgs) Handles MI_InformazioniSu.Click
        Dim WND As New WND_About
        WND.Owner = Me
        WND.ShowDialog()
    End Sub

    Private Sub MI_CaricaPreset_Click(sender As Object, e As RoutedEventArgs) Handles MI_CaricaPreset.Click
        Dim WND As New DLG_Preset(False)
        WND.Owner = Me
        WND.ShowDialog()
        UpdateLivePreview()
    End Sub

    Private Sub MI_SalvaPreset_Click(sender As Object, e As RoutedEventArgs) Handles MI_SalvaPreset.Click
        Dim WND As New DLG_Preset(True)
        WND.Owner = Me
        WND.ShowDialog()
        UpdateLivePreview()
    End Sub

    Private Sub MI_CopiaPreset_Click(sender As Object, e As RoutedEventArgs) Handles MI_CopiaPreset.Click
        Clipboard.SetText(XMLStrutture_GetXML())
    End Sub

    Private Sub MI_IncollaPreset_Click(sender As Object, e As RoutedEventArgs) Handles MI_IncollaPreset.Click
        Try
            XMLStrutture_SetXML(Clipboard.GetText)
            MsgBox(Localization.Resource_WND_Main.PresetPasteOk, MsgBoxStyle.Information, MI_IncollaPreset.Header.ToString)
            UpdateLivePreview()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, MI_IncollaPreset.Header.ToString)
        End Try
    End Sub

    Private Sub BTN_Anteprima_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Anteprima.Click
        'distruggo la livepreview se si sta generando
        If Thread_RunLivePreview.IsAlive Then Thread_RunLivePreview.Abort()

        If Coll_Files.Count = 0 Then
            MsgBox(Localization.Resource_WND_Main.FileListEmpty, MsgBoxStyle.Information, Localization.Resource_WND_Main.Btn_Anteprima)
            Return
        End If
        If Coll_Struttura.Count <= 1 Then
            If Coll_Struttura.FirstOrDefault(Function(i) i.TipoDato.Contains("EXT_")) IsNot Nothing OrElse Coll_Struttura.Count = 0 Then
                MsgBox(Localization.Resource_WND_Main.Msg_FileStructureEmpty, MsgBoxStyle.Information, Localization.Resource_WND_Main.Btn_Anteprima)
                Return
            End If
        End If
        If Coll_Struttura.FirstOrDefault(Function(i) i.TipoDato.Contains("TVDB_")) IsNot Nothing Then
            If Not Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable Then
                'controlla solo se il computer è in rete (potrebbe essere in rete ma con router disconnesso)
                'attualmente api più precise richiedono funzioni più complesse, non strettamente necessarie
                MsgBox(Localization.Resource_WND_Main.Msg_NoInternetConnection, MsgBoxStyle.Information, Localization.Resource_WND_Main.Btn_Anteprima)
                Return
            End If
        End If


        Dim CVS As ListCollectionView = CollectionViewSource.GetDefaultView(Coll_Files)

        If CVS.CustomSort IsNot Nothing Then
            'cambio il sort solo nel caso della colonna NomeFileRinominato altrimenti quando si genera il nome cambia ordine alle righe
            If CType(CVS.CustomSort, SystemComparer).SortMemberPath.Equals("NomeFileRinominato") Then
                'rigenero l'index
                For x As Integer = 0 To DG_Files.Items.Count - 1
                    DG_Files.Items(x).Index = x
                Next

                For Each Col As DataGridColumn In DG_Files.Columns
                    If Col.SortDirection IsNot Nothing Then Col.SortDirection = Nothing ' this reset UI up/down arrow
                Next

                CVS.CustomSort = Nothing

                If CVS.SortDescriptions.Count = 0 Then
                    CVS.SortDescriptions.Add(New SortDescription("Index", ListSortDirection.Ascending))
                End If
            End If

        End If


        Try
            Dim WND As New DLG_Anteprima(DG_Files.Items)
            WND.Owner = Me
            WND.ShowDialog()
            If WND.DialogResult = True Then

                If (WND.PreviewErrorsResults(ENUM_PREVIEW_RESULT.ERRORI) <> 0) OrElse (WND.PreviewErrorsResults(ENUM_PREVIEW_RESULT.CONFLITTI) <> 0) OrElse (WND.PreviewErrorsResults(ENUM_PREVIEW_RESULT.NOANTEPRIMA) <> 0) Then
                    Dim WND_Riepilogo As New DLG_Riepilogo(0, WND.PreviewErrorsResults)
                    WND_Riepilogo.Owner = Me
                    WND_Riepilogo.ShowDialog()
                End If

                BTN_Rinomina.IsEnabled = True
                BTN_RinominaSelezione.IsEnabled = True
                RowMenu.Items(0).IsEnabled = True

                BTN_Ripristina.IsEnabled = False
                BTN_RipristinaSelezione.IsEnabled = True
                RowMenu.Items(1).IsEnabled = True

                BloccoOperazioni = True
            Else
                If Not String.IsNullOrEmpty(WND.ErrorExceptionMessage) Then
                    MsgBox(Localization.Resource_Common.Exception_General & Space(1) & WND.ErrorExceptionMessage, MsgBoxStyle.Critical, Localization.Resource_WND_Main.Btn_Anteprima)
                End If
                BTN_Rinomina.IsEnabled = False
                BTN_RinominaSelezione.IsEnabled = False
                RowMenu.Items(0).IsEnabled = False

                BTN_Ripristina.IsEnabled = False
                BTN_RipristinaSelezione.IsEnabled = False
                RowMenu.Items(1).IsEnabled = False

                BloccoOperazioni = False
            End If

        Catch ex As Exception
            BTN_Rinomina.IsEnabled = False
            BTN_RinominaSelezione.IsEnabled = False
            RowMenu.Items(0).IsEnabled = False

            BTN_Ripristina.IsEnabled = False
            BTN_RipristinaSelezione.IsEnabled = False
            RowMenu.Items(1).IsEnabled = False

            BloccoOperazioni = False

            MsgBox(Localization.Resource_Common.Exception_General & Space(1) & ex.Message, MsgBoxStyle.Critical, Localization.Resource_WND_Main.Btn_Anteprima)
        End Try

        DG_Files.Items.Refresh() 'Update UI after items edited by code

        If Coll_BlackList_New.Count > 0 Then
            Badged_BlackListTips.Visibility = Visibility.Visible
            Badged_BlackListTips.Badge = Coll_BlackList_New.Count 'TODO
        End If
    End Sub

    Private Sub BTN_Rinomina_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Rinomina.Click
        If Coll_Files.Count = 0 Then
            MsgBox(Localization.Resource_WND_Main.FileListEmpty, MsgBoxStyle.Information, Localization.Resource_WND_Main.Btn_RinominaTutto)
            Return
        End If


        Try
            Dim WND As New DLG_Rinomina(Coll_Files, False)
            WND.Owner = Me
            WND.ShowDialog()
            If WND.DialogResult = True Then
                BTN_Rinomina.IsEnabled = False
                BTN_Ripristina.IsEnabled = True

                BTN_Anteprima.IsEnabled = False


                Dim WND_Riepilogo As New DLG_Riepilogo(1, WND.RenameErrorsResults)
                WND_Riepilogo.Owner = Me
                WND_Riepilogo.ShowDialog()
                If WND_Riepilogo.DialogResult = False Then
                    Coll_Files.Clear()
                    UpdateUI()
                End If

            Else
                'operazione annullata
                BTN_Rinomina.IsEnabled = True
                BTN_Ripristina.IsEnabled = True

                BTN_Anteprima.IsEnabled = True
            End If

        Catch ex As Exception
            BTN_Rinomina.IsEnabled = False
            BTN_RinominaSelezione.IsEnabled = False
            RowMenu.Items(0).IsEnabled = False

            BTN_Ripristina.IsEnabled = False
            BTN_RipristinaSelezione.IsEnabled = False
            RowMenu.Items(1).IsEnabled = False

            BTN_Anteprima.IsEnabled = True
            MsgBox(Localization.Resource_Common.Exception_General & Space(1) & ex.Message, MsgBoxStyle.Critical, Localization.Resource_WND_Main.Btn_RinominaTutto)
        End Try

        DG_Files.Items.Refresh() 'Update UI after items edited by code
    End Sub

    Private Sub BTN_RinominaSelezione_Click(sender As Object, e As RoutedEventArgs) Handles BTN_RinominaSelezione.Click
        Coll_Files_OnPropertyChanged_Enabled = True
        Try
            Dim WND As New DLG_Rinomina(DG_Files.SelectedItems, False)
            WND.Owner = Me
            WND.ShowDialog()

        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & Space(1) & ex.Message, MsgBoxStyle.Critical, Localization.Resource_WND_Main.Btn_Selezione)
        End Try
        Coll_Files_OnPropertyChanged_Enabled = False
    End Sub

    Private Sub BTN_RipristinaSelezione_Click(sender As Object, e As RoutedEventArgs) Handles BTN_RipristinaSelezione.Click
        Coll_Files_OnPropertyChanged_Enabled = True
        Try

            Dim WND As New DLG_Rinomina(DG_Files.SelectedItems, True)
            WND.Owner = Me
            WND.ShowDialog()

        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & Space(1) & ex.Message, MsgBoxStyle.Critical, Localization.Resource_WND_Main.Btn_Selezione)
        End Try
        Coll_Files_OnPropertyChanged_Enabled = False
    End Sub

    Private Sub BTN_Ripristina_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Ripristina.Click
        If Coll_Files.Count = 0 Then
            MsgBox(Localization.Resource_WND_Main.FileListEmpty, MsgBoxStyle.Information, Localization.Resource_WND_Main.Btn_RipristinaTutto)
            Return
        End If

        Try
            Dim WND As New DLG_Rinomina(Coll_Files, True)
            WND.Owner = Me
            WND.ShowDialog()
            If WND.DialogResult = True Then
                BTN_Rinomina.IsEnabled = True
                BTN_Ripristina.IsEnabled = False

                BTN_Anteprima.IsEnabled = True


                Dim WND_Riepilogo As New DLG_Riepilogo(2, WND.RenameErrorsResults)
                WND_Riepilogo.Owner = Me
                WND_Riepilogo.ShowDialog()
                If WND_Riepilogo.DialogResult = False Then
                    Coll_Files.Clear()
                    UpdateUI()
                End If

            Else
                'operazione annullata
                BTN_Rinomina.IsEnabled = False
                BTN_Ripristina.IsEnabled = True

                BTN_Anteprima.IsEnabled = False
            End If

        Catch ex As Exception
            BTN_Rinomina.IsEnabled = False
            BTN_RinominaSelezione.IsEnabled = False
            RowMenu.Items(0).IsEnabled = False

            BTN_Ripristina.IsEnabled = False
            BTN_RipristinaSelezione.IsEnabled = False
            RowMenu.Items(1).IsEnabled = False

            BTN_Anteprima.IsEnabled = True
            MsgBox(Localization.Resource_Common.Exception_General & Space(1) & ex.Message, MsgBoxStyle.Critical, Localization.Resource_WND_Main.Btn_RipristinaTutto)
        End Try

        DG_Files.Items.Refresh() 'Update UI after items edited by code
    End Sub

    Private Sub DG_Files_Sorting(sender As Object, e As DataGridSortingEventArgs) Handles DG_Files.Sorting
        Dim column As DataGridColumn = e.Column
        Dim direction As ListSortDirection = If((column.SortDirection <> ListSortDirection.Ascending), ListSortDirection.Ascending, ListSortDirection.Descending)

        column.SortDirection = direction

        Dim CVS As ListCollectionView = CollectionViewSource.GetDefaultView(Coll_Files)
        If CVS.SortDescriptions.Count = 0 Then
            CVS.SortDescriptions.Clear()
        End If

        CVS.CustomSort = New SystemComparer(direction, column.SortMemberPath)

        e.Handled = True
    End Sub

    Private Sub DG_Files_DragEnter(sender As Object, e As DragEventArgs) Handles DG_Files.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effects = DragDropEffects.Copy
        End If
    End Sub

    Private Sub DG_Files_Drop(sender As Object, e As DragEventArgs) Handles DG_Files.Drop
        If e.Data.GetDataPresent((DataFormats.FileDrop)) Then
            AvviaVerificaFiles(CType(e.Data.GetData(DataFormats.FileDrop), Array))
        End If
    End Sub

    Private Sub AvviaVerificaFiles(ElencoFiles As Array)
        Using Coll_Files.DelayNotifications
            Try
                Dim WND As New DLG_VerificaFiles(ElencoFiles)
                WND.Owner = Me
                WND.ShowDialog()
                If WND.DialogResult = True Then

                    If Not String.IsNullOrEmpty(WND.Message) Then
                        'No risultati - Categoria sbagliata?
                        MsgBox(WND.Message, MsgBoxStyle.Information, Localization.Resource_Common_Dialogs.DLG_VerificationFiles_Desc)
                    Else
                        'aggiungo i risultati
                        For Each item In WND.ResultCollItems
                            Coll_Files.Add(item)
                        Next
                    End If

                Else
                    If Not String.IsNullOrEmpty(WND.Message) Then
                        MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & WND.Message, MsgBoxStyle.Critical, Localization.Resource_Common_Dialogs.DLG_VerificationFiles_Desc)
                    End If
                End If

            Catch ex As Exception
                MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Critical, Localization.Resource_Common_Dialogs.DLG_VerificationFiles_Desc)
            End Try

        End Using
        UpdateUI()

        If DG_Files.SelectedItem Is Nothing Then
            If DG_Files.Items.Count > 0 Then DG_Files.SelectedItem = DG_Files.Items(0)
        End If

        UpdateLivePreview()
    End Sub

    Private Sub TB_Menu_Sostituisci_Checked(sender As Object, e As RoutedEventArgs) Handles TB_Menu_Sostituisci.Checked
        TB_Menu_Sostituisci.ToolTip = Localization.Resource_WND_Main.MI_SostituzioneTermini & " - " & Localization.Resource_WND_Main.Attivato
    End Sub

    Private Sub TB_Menu_Sostituisci_Unchecked(sender As Object, e As RoutedEventArgs) Handles TB_Menu_Sostituisci.Unchecked
        TB_Menu_Sostituisci.ToolTip = Localization.Resource_WND_Main.MI_SostituzioneTermini & " - " & Localization.Resource_WND_Main.Disattivato
    End Sub

    Private Sub TB_Menu_Sostituisci_Click(sender As Object, e As RoutedEventArgs) Handles TB_Menu_Sostituisci.Click, MI_Menu_SostituzioneTermini_Checkbox.Click
        XMLSettings_Save(XMLSettings_Read("CategoriaSelezionata") & "_config_SostituzioneTermini", TB_Menu_Sostituisci.IsChecked.ToString)
    End Sub

    Private Sub MI_Menu_AggiungiCartelle_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Menu_AggiungiCartella.Click, MI_Menu_AggiungiCartelle.Click
        If BloccoOperazioni Then MsgBox(Autorinomina.Localization.Resource_WND_Main.Msg_AppBusy, MsgBoxStyle.Exclamation, BTN_Menu_AggiungiCartella.ToolTip.ToString) : Return

        Using DLG As New Windows.Forms.FolderBrowserDialog

            DLG.Description = Localization.Resource_WND_Main.Desc_AggiungiCartella
            DLG.ShowNewFolderButton = False
            DLG.RootFolder = Environment.SpecialFolder.Desktop

            If DLG.ShowDialog = Forms.DialogResult.OK Then
                AvviaVerificaFiles(New String() {DLG.SelectedPath}.ToArray)
            End If

        End Using
    End Sub

    Private Sub MI_Menu_AggiungiFiles_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Menu_AggiungiFiles.Click, MI_Menu_AggiungiFiles.Click
        If BloccoOperazioni Then MsgBox(Localization.Resource_WND_Main.Msg_AppBusy, MsgBoxStyle.Exclamation, BTN_Menu_AggiungiFiles.ToolTip.ToString) : Return

        Dim DlgOpen As New OpenFileDialog()
        Dim GetExtsResult() As Object = GetExtensions(XMLSettings_Read("CategoriaSelezionata"))

        DlgOpen.Filter = GetExtsResult(2) & "|" & GetExtsResult(0)
        DlgOpen.FilterIndex = 0
        DlgOpen.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer).ToString
        DlgOpen.Multiselect = True
        DlgOpen.RestoreDirectory = True
        DlgOpen.Title = Localization.Resource_WND_Main.MI_AggiungiDeiFiles

        If DlgOpen.ShowDialog Then
            AvviaVerificaFiles(DlgOpen.FileNames)
        End If
    End Sub

    Private Sub MI_Menu_RimuoviSelezionati_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Menu_Elimina.Click, MI_Menu_RimuoviSelezionati.Click
        If DG_Files.SelectedItems.Count > 0 Then
            Dim Result As MsgBoxResult
            Result = MsgBox(Localization.Resource_WND_Main.Msg_RimuovereSelezionati, MsgBoxStyle.Question + MsgBoxStyle.YesNo, MI_Menu_RimuoviSelezionati.Header.ToString)

            If Result = MsgBoxResult.Yes Then
                For Nitm As Integer = DG_Files.SelectedItems.Count - 1 To 0 Step -1
                    Dim item As ItemFileData = DG_Files.SelectedItems(Nitm)
                    Coll_Files.Remove(item)
                Next
                UpdateUI()
            End If
        End If
    End Sub

    Private Sub MI_Menu_RimuoviTutto_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Menu_EliminaTutto.Click, MI_Menu_RimuoviTutto.Click
        If DG_Files.Items.Count > 0 Then
            Dim Result As MsgBoxResult
            Result = MsgBox(Localization.Resource_WND_Main.Msg_RimuovereTutti, MsgBoxStyle.Question + MsgBoxStyle.YesNo, MI_Menu_RimuoviTutto.Header.ToString)

            If Result = MsgBoxResult.Yes Then
                Coll_Files.Clear()
                UpdateUI()
            End If
        End If
    End Sub

    Private Sub MI_Menu_OpzioniAvanzate_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_OpzioniAvanzate.Click
        Dim WND As New WND_OpzioniAvanzate()
        WND.Owner = Me
        WND.ShowDialog()
        UpdateLivePreview()
    End Sub

    Private Sub MI_Menu_Impostazioni_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_Impostazioni.Click
        Dim WND As New WND_Impostazioni()
        WND.Owner = Me
        WND.ShowDialog()

        MI_Menu_IncludiSottocartelle.IsChecked = Boolean.Parse(XMLSettings_Read("VerificaFilesIncludiSottocartelle"))
    End Sub

    Private Sub UpdateUI()
        If Coll_Files.Count = 0 Then
            MI_Menu_RimuoviSelezionati.IsEnabled = False
            MI_Menu_RimuoviTutto.IsEnabled = False

            MI_Menu_SpostaSu.IsEnabled = False
            MI_Menu_SpostaGiu.IsEnabled = False

            BTN_Ripristina.IsEnabled = False
            BTN_Anteprima.IsEnabled = True
            BTN_Rinomina.IsEnabled = False

            BTN_RinominaSelezione.IsEnabled = False
            RowMenu.Items(0).IsEnabled = False 'cm menu rinomina selez.
            BTN_RipristinaSelezione.IsEnabled = False
            RowMenu.Items(1).IsEnabled = False 'cm menu ripristina selez.

            Dim VB As New VisualBrush
            VB.TileMode = TileMode.None
            VB.Stretch = Stretch.None
            Dim R As New Run
            R.Foreground = Brushes.Gray
            R.Text = Localization.Resource_WND_Main.Desc_ElencoVuoto
            R.FontStyle = FontStyles.Italic
            R.FontSize = 13
            VB.Visual = New TextBlock(R)
            DG_Files.Background = VB

            BloccoOperazioni = False

            Badged_BlackListTips.Visibility = Visibility.Collapsed
        Else
            MI_Menu_RimuoviSelezionati.IsEnabled = True
            MI_Menu_RimuoviTutto.IsEnabled = True

            MI_Menu_SpostaSu.IsEnabled = True
            MI_Menu_SpostaGiu.IsEnabled = True

            DG_Files.Background = Brushes.White
        End If
    End Sub

    Private Sub MI_Menu_SpostaSu_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Menu_SpostaSu.Click, MI_Menu_SpostaSu.Click
        If DG_Files.SelectedItem Is Nothing Then Return
        Dim CVS As ListCollectionView = CollectionViewSource.GetDefaultView(Coll_Files)
        Coll_Files_OnPropertyChanged_Enabled = True

        If CVS.CustomSort IsNot Nothing Then
            'rigenero l'index
            For x As Integer = 0 To DG_Files.Items.Count - 1
                DG_Files.Items(x).Index = x
            Next

            For Each Col As DataGridColumn In DG_Files.Columns
                If Col.SortDirection IsNot Nothing Then Col.SortDirection = Nothing ' this reset UI up/down arrow
            Next

            CVS.CustomSort = Nothing
        End If

        If CVS.SortDescriptions.Count = 0 Then
            CVS.SortDescriptions.Add(New SortDescription("Index", ListSortDirection.Ascending))
        End If

        Using CVS.DeferRefresh
            Dim index As Integer = DG_Files.SelectedIndex

            Dim IFD1 As ItemFileData = Coll_Files.First(Function(x) x.Index = index)
            Dim IFD2 As ItemFileData = Coll_Files.FirstOrDefault(Function(x) x.Index = (index - 1))
            If IFD2 IsNot Nothing Then
                IFD1.Index = IFD1.Index - 1
                IFD2.Index = IFD2.Index + 1
            End If
        End Using
        DG_Files.ScrollIntoView(DG_Files.SelectedItem)
        GetContainerFromIndex(Of DataGridRow)(DG_Files, DG_Files.SelectedIndex).MoveFocus(New TraversalRequest(FocusNavigationDirection.Next))

        Coll_Files_OnPropertyChanged_Enabled = False
    End Sub

    Private Sub MI_Menu_SpostaGiu_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Menu_SpostaGiu.Click, MI_Menu_SpostaGiu.Click
        If DG_Files.SelectedItem Is Nothing Then Return
        Dim CVS As ListCollectionView = CollectionViewSource.GetDefaultView(Coll_Files)
        Coll_Files_OnPropertyChanged_Enabled = True

        If CVS.CustomSort IsNot Nothing Then
            'rigenero l'index
            For x As Integer = 0 To DG_Files.Items.Count - 1
                DG_Files.Items(x).Index = x
            Next

            For Each Col As DataGridColumn In DG_Files.Columns
                If Col.SortDirection IsNot Nothing Then Col.SortDirection = Nothing ' this reset UI up/down arrow
            Next

            CVS.CustomSort = Nothing
        End If

        If CVS.SortDescriptions.Count = 0 Then
            CVS.SortDescriptions.Add(New SortDescription("Index", ListSortDirection.Ascending))
        End If

        Using CVS.DeferRefresh
            Dim index As Integer = DG_Files.SelectedIndex

            Dim IFD1 As ItemFileData = Coll_Files.First(Function(x) x.Index = index)
            Dim IFD2 As ItemFileData = Coll_Files.FirstOrDefault(Function(x) x.Index = (index + 1))
            If IFD2 IsNot Nothing Then
                IFD1.Index = IFD1.Index + 1
                IFD2.Index = IFD2.Index - 1
            End If
        End Using
        DG_Files.ScrollIntoView(DG_Files.SelectedItem)
        GetContainerFromIndex(Of DataGridRow)(DG_Files, DG_Files.SelectedIndex).MoveFocus(New TraversalRequest(FocusNavigationDirection.Next))

        Coll_Files_OnPropertyChanged_Enabled = False
    End Sub

    Private Sub MI_Menu_MostraColonne_Action(sender As Object, e As RoutedEventArgs)
        Dim MI As MenuItem = sender
        Dim Col As DataGridColumn = DG_Files.Columns.FirstOrDefault(Function(c) c.SortMemberPath = MI.Tag.ToString)
        If Col IsNot Nothing Then
            If MI.IsChecked Then
                Col.Visibility = Visibility.Visible
            Else
                Col.Visibility = Visibility.Collapsed
            End If
        End If
    End Sub

    Private Sub MI_Menu_MostraBarra_Action(sender As Object, e As RoutedEventArgs) Handles MI_Menu_MostraBarra.Checked, MI_Menu_MostraBarra.Unchecked
        If MI_Menu_MostraBarra.IsChecked Then
            BR_BarraStrumenti.Visibility = Visibility.Visible
            XMLSettings_Save("BarraStrumenti_Visibile", "True")
        Else
            BR_BarraStrumenti.Visibility = Visibility.Collapsed
            XMLSettings_Save("BarraStrumenti_Visibile", "False")
        End If
    End Sub

    Private Sub Timer_AutoCattura_Tick(sender As Object, e As EventArgs) Handles Timer_AutoCattura.Tick
        If Me.OwnedWindows.Count <= 4 Then 'disabilito temporaneamente se sono presenti altre finestre aperte, la finestra principale da come valore count: 4

            If Clipboard.ContainsFileDropList AndAlso BloccoOperazioni = False Then
                'Clipboard contiene files/cartelle

                Try
                    If Clipboard.GetFileDropList.Count > 0 Then
                        If Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchFiles")) = False Then Return

                        Dim TempArray(Clipboard.GetFileDropList.Count - 1) As String
                        Clipboard.GetFileDropList.CopyTo(TempArray, 0)

                        If XMLSettings_Read("VerificaFiles_AutoCapture_Files_OnTop") AndAlso Me.Topmost = False Then
                            Me.Topmost = True
                            Me.Topmost = False
                        End If

                        AvviaVerificaFiles(TempArray)

                    End If
                Catch ex As Exception
                    Debug.Print(ex.Message)
                Finally
                    Clipboard.Clear()
                End Try


            ElseIf Clipboard.ContainsText Then
                'Clipboard contiene struttura xml?
                Try
                    If Not String.IsNullOrEmpty(Clipboard.GetText.Trim) Then
                        If Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchPreset")) = False Then Return

                        If Clipboard.GetText.Contains("<CATEGORIA_") AndAlso Clipboard.GetText.Contains("nome=""Clipboard"">") Then
                            MI_IncollaPreset_Click(Me, Nothing)
                        End If
                    End If

                Catch ex As Exception
                    Debug.Print(ex.Message)
                Finally
                    Clipboard.Clear()
                End Try
            End If
        End If
    End Sub

    Private Sub MI_Menu_CatturaFiles_Action(sender As Object, e As RoutedEventArgs) Handles MI_Menu_CatturaFiles.Checked, MI_Menu_CatturaFiles.Unchecked
        XMLSettings_Save("MainWindow_AutoCatchFiles", MI_Menu_CatturaFiles.IsChecked.ToString)

        If Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchFiles")) OrElse Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchPreset")) Then
            If Timer_AutoCattura.IsEnabled = False Then Timer_AutoCattura.Start()
        End If

        If Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchFiles")) = False AndAlso Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchPreset")) = False Then
            Timer_AutoCattura.Stop()
        End If
    End Sub

    Private Sub MI_Menu_CatturaPreset_Action(sender As Object, e As RoutedEventArgs) Handles MI_Menu_CatturaPreset.Checked, MI_Menu_CatturaPreset.Unchecked
        XMLSettings_Save("MainWindow_AutoCatchPreset", MI_Menu_CatturaPreset.IsChecked.ToString)

        If Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchFiles")) OrElse Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchPreset")) Then
            If Timer_AutoCattura.IsEnabled = False Then Timer_AutoCattura.Start()
        End If

        If Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchFiles")) = False AndAlso Boolean.Parse(XMLSettings_Read("MainWindow_AutoCatchPreset")) = False Then
            Timer_AutoCattura.Stop()
        End If
    End Sub

    Private Sub BTN_WindowCmd_AlwaysOnTop_Click(sender As Object, e As RoutedEventArgs) Handles BTN_WindowCmd_AlwaysOnTop.Click
        MI_Menu_AlwaysOnTop.IsChecked = Not MI_Menu_AlwaysOnTop.IsChecked
    End Sub

    Private Sub MI_Menu_AlwaysOnTop_Checked(sender As Object, e As RoutedEventArgs) Handles MI_Menu_AlwaysOnTop.Checked, MI_Menu_AlwaysOnTop.Unchecked
        If BTN_WindowCmd_AlwaysOnTop.Tag.ToString.Equals("False") Then
            Me.Topmost = True
            BTN_WindowCmd_AlwaysOnTop.Tag = "True"
            XMLSettings_Save("MainWindow_AlwaysOnTop", "True")
        Else
            Me.Topmost = False
            BTN_WindowCmd_AlwaysOnTop.Tag = "False"
            XMLSettings_Save("MainWindow_AlwaysOnTop", "False")
        End If
    End Sub

    Private Sub MI_Menu_SelezionaTutto_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_SelezionaTutto.Click
        DG_Files.SelectAll()
    End Sub

    Private Sub MI_Menu_DeselezionaTutto_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_DeselezionaTutto.Click
        DG_Files.UnselectAll()
    End Sub

    Private Sub BlockDefaultCommandBinding(sender As Object, e As CanExecuteRoutedEventArgs)
        'for now this method disable only datagrid row copy to clipboard
        e.CanExecute = False
        e.Handled = True
    End Sub

    Private Sub MI_Menu_ApriFile(sender As Object, e As RoutedEventArgs)
        Dim SelItem As ItemFileData = DG_Files.SelectedItem
        Dim PercorsoFile As String = ""

        If IO.File.Exists(IO.Path.Combine(SelItem.Percorso, SelItem.NomeFile)) = False Then
            PercorsoFile = IO.Path.Combine(SelItem.Percorso, SelItem.NomeFileRinominato)
        Else
            PercorsoFile = IO.Path.Combine(SelItem.Percorso, SelItem.NomeFile)
        End If


        Dim process As System.Diagnostics.Process = Nothing
        Dim processStartInfo As New System.Diagnostics.ProcessStartInfo()

        processStartInfo.FileName = PercorsoFile

        '  If System.Environment.OSVersion.Version.Major >= 6 Then ' Windows Vista or higher
        'processStartInfo.Verb = "runas"
        'Else
        '' No need to prompt to run as admin
        'End If

        processStartInfo.Arguments = ""
        processStartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
        processStartInfo.UseShellExecute = True

        Try
            process = System.Diagnostics.Process.Start(processStartInfo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Title)
        Finally

            If Not (process Is Nothing) Then
                process.Dispose()
            End If

        End Try
    End Sub

    Private Sub MI_Menu_ProprietaFile(sender As Object, e As RoutedEventArgs)
        Try
            Dim SelItem As ItemFileData = DG_Files.SelectedItem
            Dim PercorsoFile As String = ""

            If IO.File.Exists(IO.Path.Combine(SelItem.Percorso, SelItem.NomeFile)) = False Then
                PercorsoFile = IO.Path.Combine(SelItem.Percorso, SelItem.NomeFileRinominato)
            Else
                PercorsoFile = IO.Path.Combine(SelItem.Percorso, SelItem.NomeFile)
            End If

            Interaction.ShowFileProperties(PercorsoFile)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Title)
        End Try
    End Sub

    Private Sub MI_Menu_VisualizzaStatoFile(sender As Object, e As RoutedEventArgs)
        Dim SelItem As ItemFileData = DG_Files.SelectedItem
        If Not String.IsNullOrEmpty(SelItem.StatoInfo) Then
            MsgBox(SelItem.NomeFile & vbCrLf & vbCrLf & SelItem.StatoInfo, MsgBoxStyle.OkOnly, Localization.Resource_WND_Main.MI_VisualizzaStato)
        End If
    End Sub

#Region "Check new version"

    Private Sub BW_CheckNewSWVersion_DoWork(sender As Object, e As DoWorkEventArgs) Handles BW_CheckNewSWVersion.DoWork
        Dim Result_SoftwareVersion As String = ""
        Dim DataUltimoControllo As Date = Date.ParseExact(XMLSettings_Read("MainWindow_CheckNewVersion_LastCheck"), "dd/MM/yyyy", New CultureInfo("en-US"))
        Dim Msg As Boolean = (e.Argument IsNot Nothing)

        If DataUltimoControllo.AddDays(XMLSettings_Read("MainWindow_CheckNewVersion_Days")) < Now.Date Then

            If Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() Then

                XMLSettings_Save("MainWindow_CheckNewVersion_LastCheck", Format(Date.Now, "dd/MM/yyyy"))
                Try
                    Dim wRemote As Net.WebRequest
                    wRemote = Net.WebRequest.Create("http://www.autorinomina.it/download/app_version.txt")
                    '  wRemote.Credentials = New Net.NetworkCredential("name", "pass")

                    Dim myWebResponse As Net.WebResponse = wRemote.GetResponse
                    Dim sChunks As IO.Stream = myWebResponse.GetResponseStream
                    sChunks.ReadTimeout = 200000

                    Dim srRead As System.IO.StreamReader
                    srRead = New System.IO.StreamReader(sChunks)
                    Result_SoftwareVersion = srRead.ReadToEnd.ToString

                Catch ex As Exception
                    Debug.Print(ex.Message)
                    Result_SoftwareVersion = ""
                    Msg = False
                End Try
            End If
        End If

        e.Result = New Object() {Result_SoftwareVersion, Msg}
    End Sub

    Private Sub BW_CheckNewSWVersion_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BW_CheckNewSWVersion.RunWorkerCompleted
        Dim Result As Object() = e.Result
        Dim NuovaVersione As String = ""

        If Not String.IsNullOrEmpty(Result(0)) Then
            Dim VersioneApp As Version = My.Application.Info.Version
            Dim VersioneWeb As New Version(Result(0))

            If VersioneWeb > VersioneApp Then
                BTN_NewVersion.Visibility = Visibility.Visible
                BTN_NewVersion.Tag = VersioneWeb.ToString
            End If
        End If

        If Result(1) = True Then
            If NuovaVersione.Length > 0 Then
                MsgBoxNewVersion(NuovaVersione)
            Else
                MsgBox(Localization.Resource_WND_Main.Msg_NoNewVersion, MsgBoxStyle.Information, Localization.Resource_WND_Main.MI_ControlloAggiornamenti)
            End If
        End If
    End Sub

    Private Sub BTN_NewVersion_Click(sender As Object, e As RoutedEventArgs)
        MsgBoxNewVersion(sender.Tag.ToString)
    End Sub

    Private Sub MsgBoxNewVersion(ver As String)
        Dim ResultMsg As MsgBoxResult = MsgBox(String.Format(Localization.Resource_WND_Main.Msg_NewVersion, ver), MsgBoxStyle.Question + MsgBoxStyle.YesNo, Localization.Resource_WND_Main.MI_ControlloAggiornamenti)

        If ResultMsg = MsgBoxResult.Yes Then
            Try
                System.Diagnostics.Process.Start("http://www.autorinomina.it/index.php/autorinomina/scarica")
            Catch ex As Exception
                MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Me.Title)
            End Try
        End If
    End Sub

    Private Sub MI_Menu_Aggiornamenti_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_Aggiornamenti.Click
        XMLSettings_Save("MainWindow_CheckNewVersion_LastCheck", FormatDateTime(Now.Date.AddYears(-1), DateFormat.ShortDate))
        BW_CheckNewSWVersion.RunWorkerAsync("-") 'Un contenuto qualsiasi sull'argomento fa si che al completamento mostra all'utente msgbox di risposta
    End Sub
#End Region



    Private Sub MainWindow_KeyDown(sender As Object, e As KeyEventArgs) Handles DG_Files.PreviewKeyDown, Me.KeyDown
        If Grid_Principale.Children.OfType(Of SuggestionGuide).Count > 0 Then
            If e.Key = Key.Escape Then
                Dim SGuide As SuggestionGuide = Grid_Principale.Children.OfType(Of SuggestionGuide).First
                Grid_Principale.Children.Remove(SGuide)
            End If

            If e.Key = Key.Space Then
                Dim SGuide As SuggestionGuide = Grid_Principale.Children.OfType(Of SuggestionGuide).First
                If SGuide.CurrentPage = 4 Then
                    Grid_Principale.Children.Remove(SGuide)
                Else
                    SGuide.CambiaPaginaGuida(SGuide.CurrentPage + 1)
                End If
            End If
            e.Handled = True
            Return
        End If

        If e.Key = Key.F AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then MI_Menu_AggiungiFiles_Click(Me, e)
        If e.Key = Key.D AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then MI_Menu_AggiungiCartelle_Click(Me, e)
        If e.Key = Key.P AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then MI_Menu_AggiungiPreferite_Click(Me, e)
        If e.Key = Key.Q AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then Close()

        If e.Key = Key.Z AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then MI_Menu_DeselezionaTutto_Click(Me, e)
        If e.Key = Key.F3 Then MI_Menu_CatturaFiles.IsChecked = Not MI_Menu_CatturaFiles.IsChecked
        If e.Key = Key.F4 Then MI_Menu_CatturaPreset.IsChecked = Not MI_Menu_CatturaPreset.IsChecked
        If e.Key = Key.F9 Then MI_Menu_Impostazioni_Click(Me, e)

        If e.Key = Key.F5 Then MI_TerminiBlackList_Click(Me, e)
        If e.Key = Key.F6 Then MI_SostituzioneTermini_Click(Me, e)
        If e.Key = Key.F7 Then MI_ElencoSequenziale_Click(Me, e)
        If e.Key = Key.F8 Then MI_Menu_OpzioniAvanzate_Click(Me, e)

        If e.Key = Key.F1 Then VisualizzaGuida()


        If DG_Files.SelectedItem Is Nothing OrElse DG_Files.IsKeyboardFocusWithin = False Then Return
        Dim SelectedRow = GetContainerFromIndex(Of DataGridRow)(DG_Files, DG_Files.SelectedIndex)
        If SelectedRow IsNot Nothing Then
            If SelectedRow.IsEditing Then
                If e.Key = Key.Escape Then DG_Files.CancelEdit()
                Return 'per non interferire con key qui sotto, mentre si modifica manualmente il testo
            End If
        End If

        If e.Key = Key.Delete Then MI_Menu_RimuoviSelezionati_Click(Me, e)
        If e.Key = Key.Back Then MI_Menu_RimuoviTutto_Click(Me, e)

        If e.Key = Key.Subtract Then MI_Menu_SpostaSu_Click(Me, e) : e.Handled = True
        If e.Key = Key.Add Then MI_Menu_SpostaGiu_Click(Me, e) : e.Handled = True

    End Sub

    Public Function GetContainerFromIndex(Of TContainer As DependencyObject)(ByVal itemsControl As ItemsControl, ByVal index As Integer) As TContainer
        Return DirectCast(itemsControl.ItemContainerGenerator.ContainerFromIndex(index), TContainer)
    End Function

    Private Sub MI_Menu_GuidaRapida_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_GuidaRapida.Click
        VisualizzaGuida()
    End Sub

    Private Sub MI_Menu_AggiungiPreferite_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_AggiungiPreferite.Click, BTN_Menu_AggiungiPreferiti.Click
        If BloccoOperazioni Then MsgBox(Autorinomina.Localization.Resource_WND_Main.Msg_AppBusy, MsgBoxStyle.Exclamation, MI_Menu_AggiungiPreferite.Header.ToString) : Return

        Dim NomeCartella As String = ""
        Dim PathCartella As String = MI_Menu_AggiungiPreferite.Tag.ToString
        Dim MI_Index As Integer = 1

        If String.IsNullOrEmpty(MI_Menu_AggiungiPreferite.Tag.ToString) Then
            NomeCartella = "..."
        Else
            NomeCartella = New IO.DirectoryInfo(MI_Menu_AggiungiPreferite.Tag.ToString).Name
        End If

        Dim WND As New DLG_CartellePreferite(PathCartella, NomeCartella, MI_Index)
        WND.Owner = Me
        WND.ShowDialog()

        If WND.DialogResult = True Then
            If Not PathCartella.Equals(XMLSettings_Read("FavoriteFolderPath_" & MI_Index)) Then
                MI_Menu_AggiungiPreferite.Header = Localization.Resource_WND_Main.MI_AggiungiFilesDaCartellaPreferita & " '" & New IO.DirectoryInfo(XMLSettings_Read("FavoriteFolderPath_" & MI_Index)).Name & "'"
                BTN_Menu_AggiungiPreferiti.ToolTip = Localization.Resource_WND_Main.MI_AggiungiFilesDaCartellaPreferita & " '" & New IO.DirectoryInfo(XMLSettings_Read("FavoriteFolderPath_" & MI_Index)).Name & "'"
                MI_Menu_AggiungiPreferite.Tag = XMLSettings_Read("FavoriteFolderPath_" & MI_Index)
            End If

            Try
                AvviaVerificaFiles(WND.ElencoCartelle)
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
        End If
    End Sub

    Private Sub BTN_Menu_MaiuscoleMinuscole_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Menu_MaiuscoleMinuscole.Click
        CM_MM.PlacementTarget = BTN_Menu_MaiuscoleMinuscole
        CM_MM.Placement = Primitives.PlacementMode.Bottom
        CM_MM.IsOpen = True
    End Sub

    Private Sub MI_MM_Click(sender As Object, e As RoutedEventArgs)
        Dim index As Integer
        If CType(sender, MenuItem).Parent.GetType = GetType(ContextMenu) Then
            'menu dalla barra strumenti
            index = CType(CType(sender, MenuItem).Parent, ContextMenu).Items.IndexOf(sender)
        Else
            'menu
            index = CType(CType(sender, MenuItem).Parent, MenuItem).Items.IndexOf(sender)
        End If


        For i As Integer = 0 To CM_MM.Items.Count - 1
            If i = index Then
                CM_MM.Items(i).IsChecked = True
                MI_Menu_MaiuscoleMinuscole.Items(i).IsChecked = True
            Else
                CM_MM.Items(i).IsChecked = False
                MI_Menu_MaiuscoleMinuscole.Items(i).IsChecked = False
            End If
        Next

        XMLSettings_Save(XMLSettings_Read("CategoriaSelezionata") & "_config_StileMaiuscoleMinuscole", index)
        UpdateLivePreview()
    End Sub

    Private Sub MI_Menu_Esci_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_Esci.Click
        Close()
    End Sub

    Private Sub LV_StrutturaNomeFile_PreviewDrop(sender As Object, e As DragEventArgs) Handles LV_StrutturaNomeFile.PreviewDrop
        UpdateLivePreview()
    End Sub

    Private Sub MI_Menu_LivePreview_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_LivePreview.Click, TB_Menu_LivePreview.Click
        XMLSettings_Save("MainWindow_LivePreview", MI_Menu_LivePreview.IsChecked.ToString)
        If MI_Menu_LivePreview.IsChecked Then UpdateLivePreview()
    End Sub

    Private Sub MainWindow_Deactivated(sender As Object, e As EventArgs) Handles Me.Deactivated
        If PopUp_LVItem.IsOpen Then PopUp_LVItem.IsOpen = False
    End Sub

    Private Sub DG_Files_BeginningEdit(sender As Object, e As DataGridBeginningEditEventArgs) Handles DG_Files.BeginningEdit
        DG_Files.Tag = CType(e.Row.Item, ItemFileData).NomeFileRinominato

        Dim CVS As ListCollectionView = CollectionViewSource.GetDefaultView(Coll_Files)
        If CVS.CustomSort IsNot Nothing Then
            'rigenero l'index
            For x As Integer = 0 To DG_Files.Items.Count - 1
                DG_Files.Items(x).Index = x
            Next

            For Each Col As DataGridColumn In DG_Files.Columns
                If Col.SortDirection IsNot Nothing Then Col.SortDirection = Nothing ' this reset UI up/down arrow
            Next

            CVS.CustomSort = Nothing
        End If

        If CVS.SortDescriptions.Count = 0 Then
            CVS.SortDescriptions.Add(New SortDescription("Index", ListSortDirection.Ascending))
        End If
    End Sub

    Private Sub DG_Files_CellEditEnding(sender As Object, e As DataGridCellEditEndingEventArgs) Handles DG_Files.CellEditEnding
        Coll_Files_OnPropertyChanged_Enabled = True
        If e.EditAction = DataGridEditAction.Cancel Then
            Dim IFD As ItemFileData = e.Row.Item
            'IFD.Stato = FileDataInfoStato.NULLO
            'IFD.StatoInfo = ""
            IFD.NomeFileRinominato = DG_Files.Tag
        Else
            Dim IFD As ItemFileData = e.Row.Item
            If String.IsNullOrEmpty(IFD.NomeFileRinominato.Trim) Then
                IFD.NomeFileRinominato = DG_Files.Tag
            Else

                If System.IO.Path.HasExtension(IFD.NomeFile) Then
                    If IFD.NomeFileRinominato.Length >= IO.Path.GetExtension(IFD.NomeFile).Length Then
                        Dim ExtPresent As Boolean = (IFD.NomeFileRinominato.Substring(IFD.NomeFileRinominato.Length - IO.Path.GetExtension(IFD.NomeFile).Length)).ToLower.Equals(IO.Path.GetExtension(IFD.NomeFile).ToLower)
                        If Not ExtPresent Then
                            IFD.NomeFileRinominato = IFD.NomeFileRinominato & IO.Path.GetExtension(IFD.NomeFile)
                        End If
                    Else
                        IFD.NomeFileRinominato = IFD.NomeFileRinominato & IO.Path.GetExtension(IFD.NomeFile)
                    End If
                End If
                IFD.NomeFileRinominato = ValidateChar_Replace(IFD.NomeFileRinominato)
                IFD.Stato = FileDataInfoStato.ANTEPRIMA_OK
                IFD.StatoInfo = Localization.Resource_Common.StateFileInfo_ManualEditing
            End If
        End If
        Coll_Files_OnPropertyChanged_Enabled = False
    End Sub

    Private Function FindVisualParent(Of parentItem As DependencyObject)(obj As DependencyObject) As parentItem
        Dim parent As DependencyObject = VisualTreeHelper.GetParent(obj)
        While parent IsNot Nothing AndAlso Not parent.[GetType]().Equals(GetType(parentItem))
            parent = VisualTreeHelper.GetParent(parent)
        End While
        Return TryCast(parent, parentItem)
    End Function

    Private Sub LV_StrutturaNomeFile_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles LV_StrutturaNomeFile.PreviewMouseDown
        Dim original As Object = e.OriginalSource

        If Not original.GetType.Equals(GetType(ScrollViewer)) Then
            If FindVisualParent(Of Primitives.ScrollBar)(TryCast(original, DependencyObject)) IsNot Nothing Then
                If PopUp_LVItem.IsOpen Then PopUp_LVItem.IsOpen = False
                MouseOnLV_ScrollBar = True
            End If
        End If
    End Sub

    Private Sub LV_StrutturaNomeFile_PreviewMouseUp(sender As Object, e As MouseButtonEventArgs) Handles LV_StrutturaNomeFile.PreviewMouseUp
        Dim original As Object = e.OriginalSource

        If Not original.GetType.Equals(GetType(ScrollViewer)) Then
            If FindVisualParent(Of Primitives.ScrollBar)(TryCast(original, DependencyObject)) IsNot Nothing Then
                MouseOnLV_ScrollBar = False
            End If
        End If
    End Sub

    Private Sub MI_Menu_Donazione_Click(sender As Object, e As RoutedEventArgs) Handles MI_Menu_Donazione.Click
        Try
            System.Diagnostics.Process.Start("http://www.autorinomina.it/index.php/contribuisci/donazione")
        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Me.Title)
        End Try
    End Sub

    Private Sub DG_TextBox_NomeFileRinominato_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        e.Handled = ValidateNoSpecialChar(e.Text, False)
    End Sub

    Private Sub HY_BlackListTips_Click(sender As Object, e As RoutedEventArgs)
        Badged_BlackListTips.Visibility = Visibility.Collapsed
        MI_TerminiBlackList_Click(sender, e)
    End Sub
End Class
