Imports System.Text.RegularExpressions
Imports System.Windows.Threading
Imports MahApps.Metro.Controls

Public Class CustomToolTip
    Public CM_Categoria As ContextMenu
    Public Event UpdateLivePreviewEvent As EventHandler 'ugly hack to make live preview, i don't know a solution to check obscoll items changed
    Public AvoidSave As Boolean = True
    Private ItemSel As ItemStruttura

    Public Sub Inizializza()

        'Text changed handler
        TB_Numerazione_Prefisso.AddHandler(Primitives.TextBoxBase.TextChangedEvent, New TextChangedEventHandler(
                                Sub(ByVal sender As Object, ByVal args As TextChangedEventArgs)
                                    If AvoidSave Then Return
                                    ItemSel.Opzioni = ModificaStringOpzione("NumerazionePrefisso", TB_Numerazione_Prefisso.Text, ItemSel)
                                End Sub))
        TB_Numerazione_Suffisso.AddHandler(Primitives.TextBoxBase.TextChangedEvent, New TextChangedEventHandler(
                                Sub(ByVal sender As Object, ByVal args As TextChangedEventArgs)
                                    If AvoidSave Then Return
                                    ItemSel.Opzioni = ModificaStringOpzione("NumerazioneSuffisso", TB_Numerazione_Suffisso.Text, ItemSel)
                                End Sub))
        TB_Testo.AddHandler(Primitives.TextBoxBase.TextChangedEvent, New TextChangedEventHandler(
                               Sub(ByVal sender As Object, ByVal args As TextChangedEventArgs)
                                   If AvoidSave Then Return
                                   ItemSel.Opzioni = ModificaStringOpzione("Testo", TB_Testo.Text, ItemSel)
                                   ItemSel.Testo = TB_Testo.Text
                               End Sub))
        CB_FormatoDurata.AddHandler(Primitives.TextBoxBase.TextChangedEvent, New TextChangedEventHandler(
                               Sub(ByVal sender As Object, ByVal args As TextChangedEventArgs)
                                   Try
                                       TB_FormatoDurata_Anteprima.Text = Autorinomina.Localization.Resource_BalloonToolTip.Text_Example & Space(1) & CB_FormatoDurata.Text.Replace("%H", "1").Replace("%M", "22").Replace("%S", "06")
                                   Catch ex As Exception
                                       TB_FormatoDurata_Anteprima.Text = Autorinomina.Localization.Resource_BalloonToolTip.Error_FormatNotSupported
                                   End Try
                                   If AvoidSave Then Return
                                   ItemSel.Opzioni = ModificaStringOpzione("FormatoDurataStile", CB_FormatoDurata.Text, ItemSel)
                               End Sub))
        CB_FormatoData.AddHandler(Primitives.TextBoxBase.TextChangedEvent, New TextChangedEventHandler(
                               Sub(ByVal sender As Object, ByVal args As TextChangedEventArgs)
                                   Try
                                       TB_FormatoData_Anteprima.Text = Autorinomina.Localization.Resource_BalloonToolTip.Text_Example & Space(1) & Format(Date.Now, CB_FormatoData.Text)
                                   Catch ex As Exception
                                       TB_FormatoData_Anteprima.Text = Autorinomina.Localization.Resource_BalloonToolTip.Error_FormatNotSupported
                                   End Try
                                   If AvoidSave Then Return
                                   ItemSel.Opzioni = ModificaStringOpzione("FormatoDataStile", CB_FormatoData.Text, ItemSel)
                               End Sub))
        TB_Estensioni_Sostituisci.AddHandler(Primitives.TextBoxBase.TextChangedEvent, New TextChangedEventHandler(
                              Sub(ByVal sender As Object, ByVal args As TextChangedEventArgs)
                                  If AvoidSave Then Return
                                  ItemSel.Opzioni = ModificaStringOpzione("EXTSostituisci", TB_Estensioni_Sostituisci.Text, ItemSel)
                              End Sub))
        CB_StileDimensione.AddHandler(Primitives.TextBoxBase.TextChangedEvent, New TextChangedEventHandler(
                             Sub(ByVal sender As Object, ByVal args As TextChangedEventArgs)
                                 Try
                                     TB_StileDimensione_Anteprima.Text = Autorinomina.Localization.Resource_BalloonToolTip.Text_Example & Space(1) & CB_StileDimensione.Text.Replace("%W", "600").Replace("%H", "400")
                                 Catch ex As Exception
                                     TB_StileDimensione_Anteprima.Text = Autorinomina.Localization.Resource_BalloonToolTip.Error_FormatNotSupported
                                 End Try

                                 If AvoidSave Then Return
                                 ItemSel.Opzioni = ModificaStringOpzione("SizeStyle", CB_StileDimensione.Text, ItemSel)
                             End Sub))
        TB_RegexPattern.AddHandler(Primitives.TextBoxBase.TextChangedEvent, New TextChangedEventHandler(
                               Sub(ByVal sender As Object, ByVal args As TextChangedEventArgs)
                                   If AvoidSave Then Return
                                   ItemSel.Opzioni = ModificaStringOpzione("Pattern", TB_RegexPattern.Text, ItemSel)
                               End Sub))


        If XMLSettings_Read("DisableExtensionWarning").Equals("True") Then CB_Estensioni_Abilita.IsEnabled = True
    End Sub

    Private Sub PannelliVisibili(ParamArray Grids() As Grid)
        SP_Contenuto.Children.Clear()

        For Each grid As Grid In Grids
            SP_Contenuto.Children.Add(grid)
        Next
    End Sub

    Public Sub AggiornaContenutoTT(ByRef _ItemSel As ItemStruttura)
        ItemSel = _ItemSel
        ' CB_Separatore.RemoveHandler(Primitives.TextBoxBase.TextChangedEvent,
        '  New TextChangedEventHandler(AddressOf CB_Separatore_TextChanged))

        Select Case ItemSel.TipoDato
            Case "TVDB_NumerazioneEpisodio", "AR_Numerazione", "MI_AudioTagNumero"
                TB_Titolo.Text = Localization.Resource_BalloonToolTip.Title_ByFilter
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_StileNumerazione)

                If XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_serietv") Then
                    CB_StileNumerazione.SelectedIndex = Integer.Parse(LeggiStringOpzioni("NumerazioneStile", ItemSel.Opzioni, "5"))
                Else
                    CB_StileNumerazione.SelectedIndex = Integer.Parse(LeggiStringOpzioni("NumerazioneStile", ItemSel.Opzioni))
                End If

            Case "STD_NumerazioneSequenziale"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_NumSeq
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_StileNumerazione, Grid_NumerazioneSequenziale)

                If XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_serietv") Then
                    CB_StileNumerazione.SelectedIndex = Integer.Parse(LeggiStringOpzioni("NumerazioneStile", ItemSel.Opzioni, "5"))
                Else
                    CB_StileNumerazione.SelectedIndex = Integer.Parse(LeggiStringOpzioni("NumerazioneStile", ItemSel.Opzioni))
                End If

                NUD_Numerazione_Ep.Value = Double.Parse(LeggiStringOpzioni("NumerazioneIniziaDa", ItemSel.Opzioni))
                NUD_Numerazione_Stag.Value = Double.Parse(LeggiStringOpzioni("NumerazioneStagione", ItemSel.Opzioni))
                CB_AutoPaddingZero.IsChecked = Boolean.Parse(LeggiStringOpzioni("NumerazioneAutoPadding", ItemSel.Opzioni))
                NUD_PaddingZero.Value = Double.Parse(LeggiStringOpzioni("NumerazionePadding", ItemSel.Opzioni))
                TB_Numerazione_Prefisso.Text = LeggiStringOpzioni("NumerazionePrefisso", ItemSel.Opzioni)
                TB_Numerazione_Suffisso.Text = LeggiStringOpzioni("NumerazioneSuffisso", ItemSel.Opzioni)

            Case "AR_TitoloEpisodio"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_TitleOfEpisode
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_TitoloEpisodio)

                CB_TitoloEpisodio_Concatena.IsChecked = Boolean.Parse(LeggiStringOpzioni("ConcatenaTitoloSerieTv", ItemSel.Opzioni))

            Case "AR_RegexPattern"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_RegexPattern
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_RegexPattern)

                TB_RegexPattern.Text = LeggiStringOpzioni("Pattern", ItemSel.Opzioni)
                CB_RemoveInsteadInsert.IsChecked = Boolean.Parse(LeggiStringOpzioni("RemoveInsteadInsert", ItemSel.Opzioni))

            Case "AR_Anno"
                TB_Titolo.Text = ItemSel.Testo
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_RemoveInsteadInsert)

                CB_RemoveInsteadInsert.IsChecked = Boolean.Parse(LeggiStringOpzioni("RemoveInsteadInsert", ItemSel.Opzioni))

            Case "AR_Data"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_FormatDateTime
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_FormatoData, Grid_RemoveInsteadInsert)

                CB_FormatoData.Text = LeggiStringOpzioni("FormatoDataStile", ItemSel.Opzioni)
                CB_RemoveInsteadInsert.IsChecked = Boolean.Parse(LeggiStringOpzioni("RemoveInsteadInsert", ItemSel.Opzioni))

            Case "STD_Testo"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_TextBox
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_Testo)

                TB_Testo.Text = LeggiStringOpzioni("Testo", ItemSel.Opzioni)

            Case "STD_ElencoSequenziale"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_SeqList
                PannelliVisibili(Grid_Titolo, Grid_Modifica)

            Case "STD_Separatore"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_Separator
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_Separatore)

                CB_Separatore.Text = LeggiStringOpzioni("SeparatoreCarattere", ItemSel.Opzioni)
                CB_Separatore_Spaziatura.IsChecked = Boolean.Parse(LeggiStringOpzioni("SeparatoreSpaziatura", ItemSel.Opzioni))

            Case "MI_AudioDurata", "MI_Durata"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_FormatLenght
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_FormatoDurata)

                CB_FormatoDurata.Text = LeggiStringOpzioni("FormatoDurataStile", ItemSel.Opzioni)

            Case "II_Dimensione"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_SizeOptions
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_UnitaMisura, Grid_StileDimensione)

                Dim UnitaDefault As String = LeggiStringOpzioni("UnitMeasure", ItemSel.Opzioni)
                For Each item As ComboBoxItem In CB_UnitaMisura.Items
                    If item.Tag.ToString.Equals(UnitaDefault) Then
                        CB_UnitaMisura.SelectedItem = item
                        Exit For
                    End If
                Next
                CB_StileDimensione.Text = LeggiStringOpzioni("SizeStyle", ItemSel.Opzioni)
                CB_StileDimensione_ArrotondaValori.IsChecked = Boolean.Parse(LeggiStringOpzioni("UnitMeasureRounded", ItemSel.Opzioni))

            Case "II_DPI"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_DPIOptions
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_StileDimensione)

                CB_StileDimensione.Text = LeggiStringOpzioni("SizeStyle", ItemSel.Opzioni)
                CB_StileDimensione_ArrotondaValori.IsChecked = Boolean.Parse(LeggiStringOpzioni("UnitMeasureRounded", ItemSel.Opzioni))

            Case "TVDB_TitoloEpisodio"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_TVDB_TitleOfEpisode
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_TVDB_ordineEpisodi)

                CB_TVDB_ordineEpisodi.SelectedIndex = Integer.Parse(LeggiStringOpzioni("TVDBOrdineEpisodi", ItemSel.Opzioni))

            Case "PF_DataCreazione", "PF_DataUltimaModifica", "PF_DataUltimoAccesso", "TVDB_DataPrimaTv", "IPTC_DateCreated", "EXIF_DateTimeDigitized", "EXIF_DateTimeOriginal", "EXIF_MAIN_DateTime", "MI_AudioTagDataRilascio", "MI_AudioTagDataRegistrazione"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_FormatDateTime
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_FormatoData)

                CB_FormatoData.Text = LeggiStringOpzioni("FormatoDataStile", ItemSel.Opzioni)

            Case "EXT_Opzioni"
                TB_Titolo.Text = Autorinomina.Localization.Resource_BalloonToolTip.Title_Extensions
                PannelliVisibili(Grid_Titolo, Grid_Modifica, Grid_Estensioni)

                CB_Estensioni_MM.SelectedIndex = Integer.Parse(LeggiStringOpzioni("EXTMaiuscoloMinisucolo", ItemSel.Opzioni))
                TB_Estensioni_Sostituisci.Text = LeggiStringOpzioni("EXTSostituisci", ItemSel.Opzioni)

            Case Else
                TB_Titolo.Text = ItemSel.Testo
                PannelliVisibili(Grid_Titolo, Grid_Modifica)
        End Select

        'CB_Separatore.AddHandler(Primitives.TextBoxBase.TextChangedEvent,
        '               New TextChangedEventHandler(AddressOf CB_Separatore_TextChanged))

    End Sub


    Private Sub HY_Rimuovi_Click(sender As Object, e As RoutedEventArgs) Handles HY_Rimuovi.Click
        FunzioniVarie.RiabilitaMenuCategoria(ItemSel.TipoDato, True, CM_Categoria)
        Coll_Struttura.Remove(ItemSel)
        RaiseEvent UpdateLivePreviewEvent(Nothing, Nothing)

        'Quando si modifica la struttura resetto a Default perchè non corrisponde più al Preset selezionato
        If Not XMLSettings_Read("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata")) = "Default" Then
            XMLSettings_Save("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata"), "Default")
        End If
    End Sub

    Private Sub HY_RimuoviTutto_Click(sender As Object, e As RoutedEventArgs) Handles HY_RimuoviTutto.Click
        FunzioniVarie.RiabilitaMenuCategoria(CM_Categoria)
        Coll_Struttura.Clear()
        RaiseEvent UpdateLivePreviewEvent(Nothing, Nothing)

        'Quando si modifica la struttura resetto a Default perchè non corrisponde più al Preset selezionato
        If Not XMLSettings_Read("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata")) = "Default" Then
            XMLSettings_Save("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata"), "Default")
        End If
    End Sub


    Private Sub CB_StileNumerazione_SelectionChanged(sender As System.Object, e As System.Windows.Controls.SelectionChangedEventArgs)

        If CB_StileNumerazione.SelectedIndex = 5 OrElse CB_StileNumerazione.SelectedIndex = 6 OrElse CB_StileNumerazione.SelectedIndex = 7 Then
            TBlock_Numerazione_Stag.Visibility = Windows.Visibility.Visible
            NUD_Numerazione_Stag.Visibility = Windows.Visibility.Visible
            TBlock_Numerazione_Ep.Text = Localization.Resource_BalloonToolTip.Grid_NumSeq_NumEpisodio
        Else
            TBlock_Numerazione_Stag.Visibility = Windows.Visibility.Collapsed
            NUD_Numerazione_Stag.Visibility = Windows.Visibility.Collapsed
            TBlock_Numerazione_Ep.Text = Localization.Resource_BalloonToolTip.Grid_NumSeq_NumEpisodio_IniziaDa
        End If

        If CB_StileNumerazione.SelectedIndex = 1 Then
            CB_AutoPaddingZero.Visibility = Windows.Visibility.Visible
            NUD_PaddingZero.Visibility = Windows.Visibility.Visible
        Else
            CB_AutoPaddingZero.Visibility = Windows.Visibility.Collapsed
            NUD_PaddingZero.Visibility = Windows.Visibility.Collapsed
        End If

        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("NumerazioneStile", CB_StileNumerazione.SelectedIndex.ToString, ItemSel)
    End Sub


    Private Sub CB_Separatore_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs)
        Dim TB As TextBox = sender.Template.FindName("PART_EditableTextBox", sender)
        TB.MaxLength = 1
        TB.TextAlignment = TextAlignment.Center

        If AvoidSave Then Return
        ItemSel.Testo = CB_Separatore.Text
        ItemSel.Opzioni = ModificaStringOpzione("SeparatoreCarattere", CB_Separatore.Text, ItemSel)
    End Sub

    Private Sub CB_Separatore_Spaziatura_Action(sender As Object, e As RoutedEventArgs) Handles CB_Separatore_Spaziatura.Checked, CB_Separatore_Spaziatura.Unchecked
        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("SeparatoreSpaziatura", CB_Separatore_Spaziatura.IsChecked.ToString, ItemSel)
    End Sub

    Private Sub CB_RemoveInsteadInsert_Action(sender As Object, e As RoutedEventArgs) Handles CB_RemoveInsteadInsert.Checked, CB_RemoveInsteadInsert.Unchecked
        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("RemoveInsteadInsert", CB_RemoveInsteadInsert.IsChecked.ToString, ItemSel)
    End Sub

    Private Sub BTN_Separatore_ApplicaTutti_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Separatore_ApplicaTutti.Click
        If AvoidSave Then Return
        Dim Testo As String = ItemSel.Testo
        Dim DatiXML As String = ItemSel.Opzioni

        For Each item In Coll_Struttura
            If item.TipoDato = "STD_Separatore" Then
                item.Testo = Testo
                item.Opzioni = DatiXML
            End If
        Next
    End Sub

    Private Sub CB_AutoPaddingZero_Action(sender As Object, e As RoutedEventArgs) Handles CB_AutoPaddingZero.Checked, CB_AutoPaddingZero.Unchecked
        If CB_AutoPaddingZero.IsChecked Then
            NUD_PaddingZero.IsEnabled = False
        Else
            NUD_PaddingZero.IsEnabled = True
        End If

        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("NumerazioneAutoPadding", CB_AutoPaddingZero.IsChecked.ToString, ItemSel)
    End Sub

    Private Sub CB_TitoloEpisodio_Concatena_Action(sender As Object, e As RoutedEventArgs) Handles CB_TitoloEpisodio_Concatena.Checked, CB_TitoloEpisodio_Concatena.Unchecked
        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("ConcatenaTitoloSerieTv", CB_TitoloEpisodio_Concatena.IsChecked.ToString, ItemSel)
    End Sub

    Private Sub CB_TVDB_ordineEpisodi_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CB_TVDB_ordineEpisodi.SelectionChanged
        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("TVDBOrdineEpisodi", CB_TVDB_ordineEpisodi.SelectedIndex.ToString, ItemSel)
    End Sub

    Private Sub CB_Estensioni_Abilita_Action(sender As Object, e As RoutedEventArgs) Handles CB_Estensioni_Abilita.Checked, CB_Estensioni_Abilita.Unchecked
        If CB_Estensioni_Abilita.IsChecked Then
            If XMLSettings_Read("DisableExtensionWarning").Equals("False") Then MsgBox(Localization.Resource_BalloonToolTip.Grid_Extensions_Warning.Replace("\n", vbCrLf), MsgBoxStyle.Exclamation, Application.Current.MainWindow.Title)
        Else
            CB_Estensioni_MM.SelectedIndex = 0
            TB_Estensioni_Sostituisci.Text = ""
        End If
    End Sub

    Private Sub CB_Estensioni_MM_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CB_Estensioni_MM.SelectionChanged
        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("EXTMaiuscoloMinisucolo", CB_Estensioni_MM.SelectedIndex.ToString, ItemSel)
    End Sub

    Private Sub CB_UnitaMisura_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CB_UnitaMisura.SelectionChanged
        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("UnitMeasure", CType(CB_UnitaMisura.SelectedItem, ComboBoxItem).Tag.ToString, ItemSel)
    End Sub

    Private Sub CB_StileDimensione_ArrotondaValori_Action(sender As Object, e As RoutedEventArgs) Handles CB_StileDimensione_ArrotondaValori.Checked, CB_StileDimensione_ArrotondaValori.Unchecked
        If AvoidSave Then Return
        ItemSel.Opzioni = ModificaStringOpzione("UnitMeasureRounded", CB_StileDimensione_ArrotondaValori.IsChecked.ToString, ItemSel)
    End Sub


    Private Sub NUD_Numerazione_Ep_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double?)) Handles NUD_Numerazione_Ep.ValueChanged
        If AvoidSave OrElse String.IsNullOrEmpty(e.NewValue.ToString) Then Return
        ItemSel.Opzioni = ModificaStringOpzione("NumerazioneIniziaDa", e.NewValue.ToString, ItemSel)
    End Sub

    Private Sub NUD_Numerazione_Stag_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double?)) Handles NUD_Numerazione_Stag.ValueChanged
        If AvoidSave OrElse String.IsNullOrEmpty(e.NewValue.ToString) Then Return
        ItemSel.Opzioni = ModificaStringOpzione("NumerazioneStagione", e.NewValue.ToString, ItemSel)
    End Sub

    Private Sub NUD_PaddingZero_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double?)) Handles NUD_PaddingZero.ValueChanged
        If AvoidSave OrElse String.IsNullOrEmpty(e.NewValue.ToString) Then Return
        ItemSel.Opzioni = ModificaStringOpzione("NumerazionePadding", e.NewValue.ToString, ItemSel)
    End Sub




    Private Sub ValidationOnlyNumber(sender As Object, e As TextCompositionEventArgs)
        Dim regex As New Regex("[^0-9]+")
        e.Handled = regex.IsMatch(e.Text)
    End Sub

    Private Sub ValidationNoSpecialChar(sender As Object, e As TextCompositionEventArgs)
        e.Handled = ValidateNoSpecialChar(e.Text, False)
    End Sub

End Class
