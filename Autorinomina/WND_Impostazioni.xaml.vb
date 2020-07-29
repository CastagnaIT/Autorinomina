Public Class WND_Impostazioni

    Private Sub WND_Impostazioni_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        'Varie
        CB_IncludiFiles_Sottocartelle.IsChecked = Boolean.Parse(XMLSettings_Read("VerificaFilesIncludiSottocartelle"))
        CB_IncludiFiles_Nascosti.IsChecked = Boolean.Parse(XMLSettings_Read("VerificaFilesIncludiNascosti"))
        CB_Generica_EscludiDB.IsChecked = Boolean.Parse(XMLSettings_Read("VerificaFiles_CATEGORIA_generica_EscludiDB"))
        CB_CatturaAppunti_Files_OnTop.IsChecked = Boolean.Parse(XMLSettings_Read("VerificaFiles_AutoCapture_Files_OnTop"))
        CB_AlwaysOnTop.IsChecked = Boolean.Parse(XMLSettings_Read("MainWindow_AlwaysOnTop_Keep"))
        CB_CheckNewVersion.IsChecked = Boolean.Parse(XMLSettings_Read("MainWindow_CheckNewVersion"))
        CB_PreviewCheckLocalFile.IsChecked = Boolean.Parse(XMLSettings_Read("PreviewCheckLocalFile"))
        NUD_CheckNewVersion_Days.Value = Double.Parse(XMLSettings_Read("MainWindow_CheckNewVersion_Days"))

        'Programma esterno
        CB_PrgEsterno_Attivato.IsChecked = Boolean.Parse(XMLSettings_Read("ExternalSoftware_Enabled"))
        CB_PrgEsterno_EseguiRidotto.IsChecked = Boolean.Parse(XMLSettings_Read("ExternalSoftware_RunMinimized"))
        For Each SP As ComboBoxItem In CB_PrgEsterno_Categoria.Items
            If SP.Tag.ToString = XMLSettings_Read("ExternalSoftware_Category") Then
                CB_PrgEsterno_Categoria.SelectedItem = SP
                Exit For
            End If
        Next
        TB_PrgEsterno_Percorso.Text = XMLSettings_Read("ExternalSoftware_ExePath")
        TB_PrgEsterno_ArgomentiLineaComando.Text = XMLSettings_Read("ExternalSoftware_CommandLine")

        'estensioni
        TB_Estensioni_Video.Text = XMLSettings_Read("Extensions_video")
        TB_Estensioni_Audio.Text = XMLSettings_Read("Extensions_audio")
        TB_Estensioni_Immagini.Text = XMLSettings_Read("Extensions_images")
        TB_Estensioni_MetadataImmagini.Text = XMLSettings_Read("Extensions_images_metadata")

    End Sub

    Private Sub BTN_Salva_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Salva.Click
        'Varie
        XMLSettings_Save("VerificaFilesIncludiSottocartelle", CB_IncludiFiles_Sottocartelle.IsChecked.ToString)
        XMLSettings_Save("VerificaFilesIncludiNascosti", CB_IncludiFiles_Nascosti.IsChecked.ToString)
        XMLSettings_Save("VerificaFiles_CATEGORIA_generica_EscludiDB", CB_Generica_EscludiDB.IsChecked.ToString)
        XMLSettings_Save("VerificaFiles_AutoCapture_Files_OnTop", CB_CatturaAppunti_Files_OnTop.IsChecked.ToString)
        XMLSettings_Save("MainWindow_AlwaysOnTop_Keep", CB_AlwaysOnTop.IsChecked.ToString)
        XMLSettings_Save("PreviewCheckLocalFile", CB_PreviewCheckLocalFile.IsChecked.ToString)
        XMLSettings_Save("MainWindow_CheckNewVersion", CB_CheckNewVersion.IsChecked.ToString)
        XMLSettings_Save("MainWindow_CheckNewVersion_Days", NUD_CheckNewVersion_Days.Value.ToString)

        'Programma esterno
        XMLSettings_Save("ExternalSoftware_Enabled", CB_PrgEsterno_Attivato.IsChecked.ToString)
        XMLSettings_Save("ExternalSoftware_RunMinimized", CB_PrgEsterno_EseguiRidotto.IsChecked.ToString)
        XMLSettings_Save("ExternalSoftware_Category", CB_PrgEsterno_Categoria.SelectedItem.Tag.ToString)
        XMLSettings_Save("ExternalSoftware_ExePath", TB_PrgEsterno_Percorso.Text)
        XMLSettings_Save("ExternalSoftware_CommandLine", TB_PrgEsterno_ArgomentiLineaComando.Text)

        'estensioni
        XMLSettings_Save("Extensions_video", TB_Estensioni_Video.Text)
        XMLSettings_Save("Extensions_audio", TB_Estensioni_Audio.Text)
        XMLSettings_Save("Extensions_images", TB_Estensioni_Immagini.Text)
        XMLSettings_Save("Extensions_images_metadata", TB_Estensioni_MetadataImmagini.Text)

        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Annulla.Click
        Close()
    End Sub

    Private Sub BTN_PrgEsterno_CambiaPercorso_Click(sender As Object, e As RoutedEventArgs) Handles BTN_PrgEsterno_CambiaPercorso.Click
        Dim DlgOpen As New System.Windows.Forms.OpenFileDialog()
        DlgOpen.DefaultExt = "*.exe"
        DlgOpen.Filter = "Application (*.EXE)|*.exe"
        DlgOpen.FilterIndex = 0
        DlgOpen.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer).ToString
        DlgOpen.Multiselect = False
        DlgOpen.Title = Localization.Resource_WND_Impostazioni.Msg_SelectExternalSoftware
        DlgOpen.FileName = ""

        If DlgOpen.ShowDialog = Forms.DialogResult.OK Then
            TB_PrgEsterno_Percorso.Text = DlgOpen.FileName
        End If
    End Sub
End Class