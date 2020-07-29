Public Class DLG_Preset
    Dim ModalitaSalva As Boolean = False
    Dim StrutturaAttualeBackup As New CollItemsStruttura
    Dim AvoidPreview As Boolean = True

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(_ModalitaSalva As Boolean)

        ' La chiamata è richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        ModalitaSalva = _ModalitaSalva
    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If ModalitaSalva Then
            Me.Title = Autorinomina.Localization.Resource_Common_Dialogs.DLG_Presets_Title_Save
            BTN_Carica.Visibility = Windows.Visibility.Collapsed
            TB_NomePreset.Focus()
        Else
            Me.Title = Autorinomina.Localization.Resource_Common_Dialogs.DLG_Presets_Title_Load
            BTN_Salva.Visibility = Windows.Visibility.Collapsed

            Grid_Contenuto.RowDefinitions(0).Height = New GridLength(0)
            TB_NomePreset.IsEnabled = False
        End If

        For Each NomeStruttura As String In XMLStrutture_Elenca()
            LB_Presets.Items.Add(NomeStruttura)

            If XMLSettings_Read("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata")).Equals(NomeStruttura) Then
                LB_Presets.SelectedValue = NomeStruttura
            End If
        Next

        StrutturaAttualeBackup.AddRange(Coll_Struttura)
        AvoidPreview = False
    End Sub

    Private Sub LB_Presets_SelectionChanged(sender As System.Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles LB_Presets.SelectionChanged
        If ModalitaSalva Then
            TB_NomePreset.Text = LB_Presets.SelectedItem
        Else
            If Not AvoidPreview Then XMLStrutture_Read(LB_Presets.SelectedItem)
        End If
    End Sub

    Private Sub BTN_Carica_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Carica.Click
        If LB_Presets.SelectedIndex = -1 Then
            MsgBox(AutoRinomina.Localization.Resource_Common_Dialogs.DLG_Presets_SelectPreset, MsgBoxStyle.Exclamation, BTN_Carica.Content.ToString)
            Return
        End If

        XMLSettings_Save("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata"), LB_Presets.SelectedItem)

        Me.DialogResult = True
        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Annulla.Click
        Coll_Struttura.Clear()
        For Each item In StrutturaAttualeBackup
            Coll_Struttura.Add(item)
        Next

        Me.DialogResult = False
        Close()
    End Sub

    Private Sub BTN_Salva_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Salva.Click
        If String.IsNullOrEmpty(TB_NomePreset.Text.Trim) Then
            MsgBox(Autorinomina.Localization.Resource_Common_Dialogs.DLG_Presets_NeedPresetName, MsgBoxStyle.Exclamation, BTN_Salva.Content.ToString)
            TB_NomePreset.Focus()
        Else
            If LB_Presets.Items.Contains(TB_NomePreset.Text) Then
                Dim Result As MsgBoxResult = MsgBox(Localization.Resource_Common_Dialogs.DLG_Presets_OverwritePreset, MsgBoxStyle.Question + MsgBoxStyle.YesNo, BTN_Salva.Content.ToString)
                If Result = MsgBoxResult.No Then Return
            End If

            XMLStrutture_Save(TB_NomePreset.Text)
            Close()
        End If
    End Sub

    Private Sub BTN_Elimina_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Elimina.Click
        If LB_Presets.SelectedIndex = -1 Then
            MsgBox(Autorinomina.Localization.Resource_Common_Dialogs.DLG_Presets_SelectPreset, MsgBoxStyle.Exclamation, BTN_Elimina.Content.ToString)
            Return
        Else

            Dim Result As MsgBoxResult = MsgBox(Localization.Resource_Common_Dialogs.DLG_Presets_DeletePreset, MsgBoxStyle.Question + MsgBoxStyle.YesNo, BTN_Elimina.Content.ToString)
            If Result = MsgBoxResult.No Then Return

            XMLStrutture_Delete(LB_Presets.SelectedItem)

            If XMLSettings_Read("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata")) = LB_Presets.SelectedItem Then
                XMLStrutture_Save("Default")
                XMLSettings_Save("StrutturaSelezionata_" & XMLSettings_Read("CategoriaSelezionata"), "Default")
            End If

            LB_Presets.Items.Remove(LB_Presets.SelectedItem)
        End If
    End Sub

    Private Sub ListBoxItem_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        If ModalitaSalva Then
            BTN_Salva_Click(sender, e)
        Else
            BTN_Carica_Click(sender, e)
        End If
    End Sub
End Class
