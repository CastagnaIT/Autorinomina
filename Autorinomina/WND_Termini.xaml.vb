Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Public Class WND_Termini

    Dim Coll_Termini_SerieTv As New CollItemsTermini
    Dim Coll_Termini_Video As New CollItemsTermini
    Dim Coll_Termini_Audio As New CollItemsTermini
    Dim Coll_Termini_Immagini As New CollItemsTermini
    Dim Coll_Termini_Generica As New CollItemsTermini
    Dim WND_Type As String
    Dim CategoriaSelezionata As String

    Public Sub New(ByVal _WND_Type As String)

        ' La chiamata è richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        WND_Type = _WND_Type
    End Sub

    Private Sub WND_BlackList_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Coll_Termini_SerieTv.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & WND_Type & "_CATEGORIA_serietv.xml", WND_Type))
        Coll_Termini_Video.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & WND_Type & "_CATEGORIA_video.xml", WND_Type))
        Coll_Termini_Audio.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & WND_Type & "_CATEGORIA_audio.xml", WND_Type))
        Coll_Termini_Immagini.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & WND_Type & "_CATEGORIA_immagini.xml", WND_Type))
        Coll_Termini_Generica.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DataPath & "\" & WND_Type & "_CATEGORIA_generica.xml", WND_Type))

        If WND_Type = "Sostituzioni" Then
            DG_Termini.Columns(1).Visibility = Visibility.Visible
            Me.Title = Localization.Resource_WND_Main.MI_SostituzioneTermini
            LB_Descrizione.Content = Localization.Resource_WND_Termini.Sostituzioni_desc

        Else
            Me.Title = Localization.Resource_WND_Main.MI_TerminiInBlackList
            LB_Descrizione.Content = Localization.Resource_WND_Termini.BlackList_desc
        End If

        Select Case XMLSettings_Read("CategoriaSelezionata")
            Case "CATEGORIA_serietv"
                CType(SP_Categorie.Children(0), Primitives.ToggleButton).IsChecked = True
            Case "CATEGORIA_video"
                CType(SP_Categorie.Children(1), Primitives.ToggleButton).IsChecked = True
            Case "CATEGORIA_audio"
                CType(SP_Categorie.Children(2), Primitives.ToggleButton).IsChecked = True
            Case "CATEGORIA_immagini"
                CType(SP_Categorie.Children(3), Primitives.ToggleButton).IsChecked = True
            Case "CATEGORIA_generica"
                CType(SP_Categorie.Children(4), Primitives.ToggleButton).IsChecked = True
        End Select
    End Sub

    Private Sub WND_Termini_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        'Content Rendered permette di visualizzare l'animazione del badge in apertura della finestra
        If XMLSettings_Read("CategoriaSelezionata").Equals("CATEGORIA_serietv") Then
            If Coll_BlackList_New.Count > 0 Then Badged_BTN_Suggeriti.Badge = Coll_BlackList_New.Count & Space(1) & "suggeriti"
        End If
    End Sub

    Private Sub ToggleCategoria_Checked(sender As Object, e As RoutedEventArgs)
        Dim TB As Primitives.ToggleButton = sender

        For Each ChildTB As Primitives.ToggleButton In SP_Categorie.Children
            If Not ChildTB.Tag = TB.Tag Then
                ChildTB.IsChecked = False
            End If
        Next

        CategoriaSelezionata = TB.Tag.ToString

        Select Case TB.Tag.ToString
            Case "CATEGORIA_serietv"
                DG_Termini.ItemsSource = Coll_Termini_SerieTv
                If Coll_BlackList_New.Count > 0 Then Badged_BTN_Suggeriti.Badge = Coll_BlackList_New.Count & Space(1) & "suggeriti"
            Case "CATEGORIA_video"
                DG_Termini.ItemsSource = Coll_Termini_Video
                Badged_BTN_Suggeriti.Badge = Nothing
            Case "CATEGORIA_audio"
                DG_Termini.ItemsSource = Coll_Termini_Audio
                Badged_BTN_Suggeriti.Badge = Nothing
            Case "CATEGORIA_immagini"
                DG_Termini.ItemsSource = Coll_Termini_Immagini
                Badged_BTN_Suggeriti.Badge = Nothing
            Case "CATEGORIA_generica"
                DG_Termini.ItemsSource = Coll_Termini_Generica
                Badged_BTN_Suggeriti.Badge = Nothing
        End Select

        TS_OrdineTermini.IsChecked = Boolean.Parse(XMLSettings_Read(WND_Type & "_" & CategoriaSelezionata & "_ordineAuto"))
    End Sub

    Private Sub TickImage_PreviewMouseUp(sender As Object, e As MouseButtonEventArgs)
        Dim selItem As ItemTermine = DG_Termini.SelectedItem
        selItem.CaseSensitive = (Not Boolean.Parse(selItem.CaseSensitive)).ToString
    End Sub

    Private Sub BTN_InserisciTermine_Click(sender As Object, e As RoutedEventArgs) Handles BTN_InserisciTermine.Click
        Dim CountIniziale_CollBlackListNew As Integer = Coll_BlackList_New.Count
        Dim CountIniziale_CollTermini As Integer = CType(DG_Termini.ItemsSource, CollItemsTermini).Count


        Dim WND As New DLG_AggiungiTermine(WND_Type, DG_Termini.ItemsSource, CategoriaSelezionata)
        WND.Owner = Me
        WND.ShowDialog()

        If CType(DG_Termini.ItemsSource, CollItemsTermini).Count <> CountIniziale_CollTermini Then
            If XMLSettings_Read(WND_Type & "_" & CategoriaSelezionata & "_ordineAuto").Equals("True") Then
                Riordina(DG_Termini.ItemsSource)
            End If
        End If

        If CountIniziale_CollBlackListNew <> Coll_BlackList_New.Count Then
            Badged_BTN_Suggeriti.Badge = Coll_BlackList_New.Count & Space(1) & "suggeriti"
        ElseIf Coll_BlackList_New.Count = 0 Then
            Badged_BTN_Suggeriti.Badge = ""
        End If
    End Sub

    Private Sub BTN_EliminaTermine_Click(sender As Object, e As RoutedEventArgs) Handles BTN_EliminaTermine.Click
        If DG_Termini.SelectedItems.Count > 0 Then

            Dim Result As MsgBoxResult = MsgBox(Localization.Resource_WND_Termini.Msg_RemoveSelected, MsgBoxStyle.Question + MsgBoxStyle.YesNo, BTN_EliminaTermine.Content.ToString)

            If Result = MsgBoxResult.Yes Then
                Dim items As System.Collections.IList = DirectCast(DG_Termini.SelectedItems, System.Collections.IList)
                Dim collection = items.Cast(Of ItemTermine)

                Dim CollAttuale As CollItemsTermini = DG_Termini.ItemsSource
                CollAttuale.RemoveRange(collection)
            End If
        End If
    End Sub

    Public Sub Riordina(ByRef CollTermini As CollItemsTermini)
        Dim ListaRiordinata As New CollItemsTermini
        Dim TempColl As New CollItemsTermini
        Dim Temp_ParoleSpostateSimiliAggiunte As New ArrayList

        '---------ottengo SOLO item con valore true in CASE SENSITIVE
        Dim Lista_Case = CollTermini.Where(Function(item) (item.CaseSensitive.Equals("True")))

        'riordino in modo decrescente per termine
        Lista_Case = Lista_Case.OrderByDescending(Function(item) (item.Termine))

        For Each item As ItemTermine In Lista_Case

            'cerco se l'elemento attuale è già esistente in tempcoll
            If (TempColl.Where(Function(w) (w.Termine.Equals(item.Termine))).Count = 0) Then

                'ricerco parole simili da aggiungere
                Dim ElencoSimili = From s In Lista_Case
                                   Where Regex.IsMatch(s.Termine, "(^)" & Regex.Escape(item.Termine) & "|" & Regex.Escape(item.Termine) & "($)", RegexOptions.None)

                'riordino in crescente le parole simili
                ElencoSimili = ElencoSimili.OrderBy(Function(s) (s.Termine))

                For Each item_simile As ItemTermine In ElencoSimili
                    'se la parola simile è diversa
                    If Not item.Termine.Equals(item_simile.Termine) Then
                        'aggiungi quella simile o sposta quelle aggiunte
                        'cerca se la simile è già stata aggiunta
                        If (TempColl.Where(Function(w) (w.Termine.Equals(item_simile.Termine))).Count = 0) Then
                            TempColl.Add(item_simile)
                            Temp_ParoleSpostateSimiliAggiunte.Add(item_simile.Termine)
                        Else
                            If Not Temp_ParoleSpostateSimiliAggiunte.Contains(item_simile.Termine) Then
                                'se non è una parola spostata o simile già aggiunta
                                TempColl.Remove(item_simile)
                                TempColl.Add(item_simile)
                                Temp_ParoleSpostateSimiliAggiunte.Add(item_simile.Termine)
                            End If
                        End If
                    End If
                Next

                'se non è ancora stato aggiungo l'item corrente lo aggiungo ora
                If (TempColl.Where(Function(w) (w.Termine.Equals(item.Termine))).Count = 0) Then
                    TempColl.Add(item)
                End If
            End If
        Next

        ListaRiordinata.AddRange(TempColl)

        Temp_ParoleSpostateSimiliAggiunte.Clear()
        TempColl = New CollItemsTermini

        '---------ottengo SOLO item con valore true in CASE  NON SENSITIVE
        Lista_Case = CollTermini.Where(Function(item) (item.CaseSensitive.Equals("False")))
        'riordino in modo decrescente per termine
        Lista_Case = Lista_Case.OrderByDescending(Function(item) item.Termine.ToLower,
                                     Comparer(Of String).Create(Function(key1, key2) key1.CompareTo(key2)))

        For Each item As ItemTermine In Lista_Case

            'cerco se l'elemento attuale è già esistente in tempcoll
            If (TempColl.Where(Function(w) (w.Termine.Equals(item.Termine))).Count = 0) Then

                'ricerco parole simili da aggiungere
                Dim ElencoSimili = From s In Lista_Case
                                   Where Regex.IsMatch(s.Termine, "(^)" & Regex.Escape(item.Termine) & "|" & Regex.Escape(item.Termine) & "($)", RegexOptions.IgnoreCase)

                'riordino in crescente le parole simili
                ElencoSimili = ElencoSimili.OrderBy(Function(a) a.Termine.ToLower,
                                     Comparer(Of String).Create(Function(key1, key2) key1.CompareTo(key2)))

                For Each item_simile As ItemTermine In ElencoSimili
                    'se la parola simile è diversa
                    If Not item.Termine.Equals(item_simile.Termine, StringComparison.OrdinalIgnoreCase) Then
                        'aggiungi quella simile o sposta quelle aggiunte
                        'cerca se la simile è già stata aggiunta
                        If (TempColl.Where(Function(w) (w.Termine.Equals(item_simile.Termine, StringComparison.OrdinalIgnoreCase))).Count = 0) Then
                            TempColl.Add(item_simile)
                            Temp_ParoleSpostateSimiliAggiunte.Add(item_simile.Termine.ToLower)
                        Else
                            If Temp_ParoleSpostateSimiliAggiunte.Contains(item_simile.Termine.ToLower) Then
                                'se non è una parola spostata o simile già aggiunta
                                TempColl.Remove(item_simile)
                                TempColl.Add(item_simile)
                                Temp_ParoleSpostateSimiliAggiunte.Add(item_simile.Termine.ToLower)
                            End If
                        End If
                    End If
                Next

                'se non è ancora stato aggiungo l'item corrente lo aggiungo ora
                If (TempColl.Where(Function(w) (w.Termine.Equals(item.Termine))).Count = 0) Then
                    TempColl.Add(item)
                End If
            End If
        Next

        ListaRiordinata.AddRange(TempColl)

        CollTermini = ListaRiordinata
    End Sub

    Private Sub TS_OrdineTermini_Click(sender As Object, e As RoutedEventArgs) Handles TS_OrdineTermini.Click
        XMLSettings_Save(WND_Type & "_" & CategoriaSelezionata & "_ordineAuto", TS_OrdineTermini.IsChecked.ToString)
        If TS_OrdineTermini.IsChecked Then Riordina(DG_Termini.ItemsSource)
    End Sub

    Private Sub TS_OrdineTermini_Action(sender As Object, e As RoutedEventArgs) Handles TS_OrdineTermini.Checked, TS_OrdineTermini.Unchecked
        BTN_SpostaGiu.IsEnabled = Not TS_OrdineTermini.IsChecked
        BTN_SpostaSu.IsEnabled = Not TS_OrdineTermini.IsChecked
    End Sub

    Private Sub BTN_ImportaTermini_Click(sender As Object, e As RoutedEventArgs) Handles BTN_ImportaTermini.Click
        Dim DlgOpen As New OpenFileDialog()
        DlgOpen.DefaultExt = "*.xml"
        DlgOpen.Filter = "XML eXtensible Markup Language (*.XML)|*.xml"
        DlgOpen.FilterIndex = 0
        DlgOpen.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer).ToString
        DlgOpen.Multiselect = False
        DlgOpen.Title = BTN_ImportaTermini.Content.ToString

        If DlgOpen.ShowDialog Then
            Try
                Dim NAggiunti As Integer = 0
                Dim NEsistenti As String = 0

                Select Case IO.Path.GetExtension(DlgOpen.FileName).ToLower
                    Case ".xml"
                        Dim Coll_Termini_File As New CollItemsTermini
                        Dim Coll_Attuale As CollItemsTermini = DG_Termini.ItemsSource
                        Coll_Termini_File.AddRange(XmlSerialization.ReadFromXmlFile(Of ItemTermine)(DlgOpen.FileName, WND_Type))

                        For Each item In Coll_Termini_File

                            If (Coll_Attuale.Where(Function(w) (w.Termine.Equals(item.Termine) And (w.CaseSensitive.Equals(item.CaseSensitive)))).Count = 0) Then
                                Coll_Attuale.Add(item)
                                NAggiunti += 1
                            Else
                                NEsistenti += 1
                            End If
                        Next

                    Case ".txt"

                        'TODO: import from txt
                End Select

                MsgBox(String.Format(Localization.Resource_WND_Termini.Msg_ImportFile, vbCrLf, NAggiunti & vbCrLf, Space(1) & NEsistenti), MsgBoxStyle.Information)

            Catch ex As Exception
                MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Private Sub BTN_EsportaTermini_Click(sender As Object, e As RoutedEventArgs) Handles BTN_EsportaTermini.Click
        Dim DlgSave As New SaveFileDialog()
        DlgSave.DefaultExt = "*.xml"
        DlgSave.Filter = "XML eXtensible Markup Language (*.XML)|*.xml"
        DlgSave.FilterIndex = 0
        DlgSave.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer).ToString
        DlgSave.Title = BTN_EsportaTermini.Content.ToString
        DlgSave.FileName = WND_Type & "_" & CategoriaSelezionata

        If DlgSave.ShowDialog Then

            Try
                Select Case IO.Path.GetExtension(DlgSave.FileName).ToLower
                    Case ".xml"
                        Dim Coll_Attuale As CollItemsTermini = DG_Termini.ItemsSource
                        XmlSerialization.WriteToXmlFile(DlgSave.FileName, Coll_Attuale, WND_Type, False)

                End Select

                MsgBox(Localization.Resource_WND_Termini.Msg_ExportFile, MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Private Sub BTN_SpostaSu_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SpostaSu.Click
        If DG_Termini.SelectedIndex > 0 AndAlso DG_Termini.SelectedItems.Count = 1 Then

            Dim Coll_Attuale As CollItemsTermini = DG_Termini.ItemsSource

            Dim index As Integer = DG_Termini.SelectedIndex

            Coll_Attuale.Move(index, index - 1)
            DG_Termini.ScrollIntoView(DG_Termini.SelectedItem)
        End If
    End Sub

    Private Sub BTN_SpostaGiu_Click(sender As Object, e As RoutedEventArgs) Handles BTN_SpostaGiu.Click
        If DG_Termini.SelectedIndex < DG_Termini.Items.Count - 1 AndAlso DG_Termini.SelectedItems.Count = 1 Then

            Dim Coll_Attuale As CollItemsTermini = DG_Termini.ItemsSource

            Dim index As Integer = DG_Termini.SelectedIndex

            Coll_Attuale.Move(index, index + 1)
            DG_Termini.ScrollIntoView(DG_Termini.SelectedItem)
        End If
    End Sub

    Private Sub BTN_Salva_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Salva.Click

        ' qui il riordino dei termini non sarebbe necessario, ma se si modificano i case sensitive
        ' attraverso la tabella, poi l'elenco deve essere riordinato

        If XMLSettings_Read(WND_Type & "_CATEGORIA_serietv_ordineAuto").Equals("True") Then Riordina(Coll_Termini_SerieTv)
        If XMLSettings_Read(WND_Type & "_CATEGORIA_video_ordineAuto").Equals("True") Then Riordina(Coll_Termini_Video)
        If XMLSettings_Read(WND_Type & "_CATEGORIA_audio_ordineAuto").Equals("True") Then Riordina(Coll_Termini_Audio)
        If XMLSettings_Read(WND_Type & "_CATEGORIA_immagini_ordineAuto").Equals("True") Then Riordina(Coll_Termini_Immagini)
        If XMLSettings_Read(WND_Type & "_CATEGORIA_generica_ordineAuto").Equals("True") Then Riordina(Coll_Termini_Generica)

        XmlSerialization.WriteToXmlFile(DataPath & "\" & WND_Type & "_CATEGORIA_serietv.xml", Coll_Termini_SerieTv, WND_Type, False)
        XmlSerialization.WriteToXmlFile(DataPath & "\" & WND_Type & "_CATEGORIA_video.xml", Coll_Termini_Video, WND_Type, False)
        XmlSerialization.WriteToXmlFile(DataPath & "\" & WND_Type & "_CATEGORIA_audio.xml", Coll_Termini_Audio, WND_Type, False)
        XmlSerialization.WriteToXmlFile(DataPath & "\" & WND_Type & "_CATEGORIA_immagini.xml", Coll_Termini_Immagini, WND_Type, False)
        XmlSerialization.WriteToXmlFile(DataPath & "\" & WND_Type & "_CATEGORIA_generica.xml", Coll_Termini_Generica, WND_Type, False)

        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Annulla.Click
        Close()
    End Sub

End Class
